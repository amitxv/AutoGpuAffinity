import sys
import ctypes
import os
import textwrap
import time
import subprocess
import winreg
import csv
import math
import wmi
from tabulate import tabulate

subprocess_null = {"stdout": subprocess.DEVNULL, "stderr": subprocess.DEVNULL}
ntdll = ctypes.WinDLL("ntdll.dll")

def parse_config(config_path) -> dict:
    """parse a simple configuration file and return a dict of the settings/values"""
    config = {}
    with open(config_path, "r", encoding="UTF-8") as f:
        for line in f:
            if "//" not in line:
                line = line.strip("\n")
                setting, _equ, value = line.rpartition("=")
                if setting != "" and value != "":
                    if value.isdigit():
                        value = int(value)
                    config[setting] = value
    return config

def create_lava_cfg(enable_fullscren, x_resolution, y_resolution) -> None:
    """creates the lava-triangle configuration file"""
    lava_triangle_folder = f"{os.environ['USERPROFILE']}\\AppData\\Roaming\\liblava\\lava triangle"
    os.makedirs(lava_triangle_folder, exist_ok=True)
    lava_triangle_config = f"{lava_triangle_folder}\\window.json"

    if os.path.exists(lava_triangle_config):
        os.remove(lava_triangle_config)

    config_content = [
        "{",
        '    "default": {',
        '        "decorated": true,',
        '        "floating": false,',
        f'        "fullscreen": {"true" if enable_fullscren else "false"},',
        f'        "height": {y_resolution},',
        '        "maximized": false,',
        '        "monitor": 0,',
        '        "resizable": true,',
        f'        "width": {x_resolution},',
        '        "x": 0,',
        '        "y": 0',
        "    }",
        "}"
    ]
    with open(lava_triangle_config, "a", encoding="UTF-8") as f:
        for i in config_content:
            f.write(f"{i}\n")

def start_afterburner(path, profile) -> None:
    """starts msi afterburner and loads a profile"""
    subprocess.Popen([path, f"-Profile{profile}"])
    time.sleep(7)
    kill_processes("MSIAfterburner.exe")

def kill_processes(*targets) -> None:
    """kill windows processes"""
    for process in targets:
        subprocess.run(["taskkill", "/F", "/IM", process], **subprocess_null, check=False)

def write_key(path, value_name, data_type, value_data) -> None:
    """write keys to windows registry"""
    with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
        winreg.SetValueEx(key, value_name, 0, data_type, value_data)

def delete_key(path, value_name) -> None:
    """delete keys in windows registry"""
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY) as key:
            try:
                winreg.DeleteValue(key, value_name)
            except FileNotFoundError:
                pass
    except FileNotFoundError:
        pass

def convert_affinity(cpu) -> int:
    """convert cpu affinity to the decimal representation"""
    affinity = 0
    affinity |= 1 << cpu
    return affinity

def apply_affinity(hwids, action, dec_affinity = -1) -> None:
    """apply interrupt affinity policy to the graphics driver"""
    for hwid in hwids:
        policy_path = f"SYSTEM\\ControlSet001\\Enum\\{hwid}\\Device Parameters\\Interrupt Management\\Affinity Policy"
        if action == 1 and dec_affinity > -1:
            bin_affinity = bin(dec_affinity).replace("0b", "")
            le_hex = int(bin_affinity, 2).to_bytes(8, "little").rstrip(b"\x00")
            write_key(policy_path, "DevicePolicy", 4, 4)
            write_key(policy_path, "AssignmentSetOverride", 3, le_hex)
        elif action == 0:
            delete_key(policy_path, "DevicePolicy")
            delete_key(policy_path, "AssignmentSetOverride")

    subprocess.run(["bin\\restart64\\restart64.exe", "/q"], check=False)

def compute_frametimes(frametime_data, metric, value = -1) -> float:
    """calculate various metrics based on frametime data"""
    result = 0
    if metric == "Max":
        result = frametime_data["min"]
    elif metric == "Avg":
        result = frametime_data["sum"] / frametime_data["len"]
    elif metric == "Min":
        result = frametime_data["max"]
    elif metric == "Percentile" and value > -1:
        result = frametime_data["frametimes"][math.ceil(value / 100 * frametime_data["len"]) - 1]
    elif metric == "Lows" and value > -1:
        current_total = 0
        for present in frametime_data["frametimes"]:
            current_total += present
            if current_total >= value / 100 * frametime_data["sum"]:
                result = present
                break
    elif metric == "STDEV":
        mean = frametime_data["sum"] / frametime_data["len"]
        dev = [x - mean for x in frametime_data["frametimes"]]
        dev2 = [x * x for x in dev]
        result = math.sqrt(sum(dev2) / frametime_data["len"])
    return 1000 / result

def timer_resolution(enabled) -> int:
    """
    sets the process timer-resolution to 1000hz
    """
    min_res = ctypes.c_ulong()
    max_res = ctypes.c_ulong()
    curr_res = ctypes.c_ulong()

    ntdll.NtQueryTimerResolution(ctypes.byref(min_res), ctypes.byref(max_res), ctypes.byref(curr_res))

    if max_res.value <= 10000 and ntdll.NtSetTimerResolution(10000, int(enabled), ctypes.byref(curr_res)) == 0:
        return 0
    return 1

def main() -> int:
    """program entrypoint"""

    if not ctypes.windll.shell32.IsUserAnAdmin():
        print("error: administrator privileges required")
        return 1

    if getattr(sys, "frozen", False):
        os.chdir(os.path.dirname(sys.executable))
    elif __file__:
        os.chdir(os.path.dirname(__file__))

    version = "0.13.3"
    cfg = parse_config("config.txt")
    present_mon = "PresentMon-1.6.0-x64.exe"
    cpu_count = os.cpu_count()
    user32 = ctypes.windll.user32
    master_table = [[""]]
    videocontroller_hwids = [x.PnPDeviceID for x in wmi.WMI().Win32_VideoController()]

    if sys.getwindowsversion().major >= 10:
        present_mon = "PresentMon-1.8.0-x64.exe"

    if cpu_count is not None:
        cpu_count -= 1
    else:
        print("error: unable to get CPU count")
        return 1

    if len(videocontroller_hwids) == 0:
        print("error: no graphics card found")
        return 1

    if not all(os.path.exists(f"bin\\{x}") for x in ["liblava\\lava-triangle.exe", "liblava\\res.zip", f"PresentMon\\{present_mon}", "restart64\\restart64.exe"]):
        print("error: missing binaries")
        return 1

    if (cfg["cache_duration"] < 0) or (cfg["duration"] <= 0):
        print("error: invalid durations in config")
        return 1

    if cfg["dpcisr"] and not os.path.exists(cfg["xperf_path"]):
        cfg["dpcisr"] = 0

    if cfg["afterburner_profile"] and not os.path.exists(cfg["afterburner_path"]):
        cfg["afterburner_profile"] = 0

    if not (cfg["custom_cores"].startswith("[") and cfg["custom_cores"].endswith("]")):
        print("error: surrounding brackets for custom_cores value not found")
        return 1

    cfg["custom_cores"] = [int(x) for x in cfg["custom_cores"][1:-1].replace(" ", "").split(",") if x != "" and 0 <= int(x) <= cpu_count]
    cfg["custom_cores"] = list(dict.fromkeys(cfg["custom_cores"]))
    cfg["custom_cores"].sort()

    cfg["colored_output"] = cfg["colored_output"] and sys.getwindowsversion().major >= 10

    os.makedirs("captures", exist_ok=True)
    output_path = f"captures\\AutoGpuAffinity-{time.strftime('%d%m%y%H%M%S')}"

    runtime_info = f"""
    AutoGpuAffinity v{version}
    GitHub - https://github.com/amitxv

        {"Session Directory" : <24}.\\{output_path}
        {"Cache Duration" : <24}{cfg["cache_duration"]} seconds
        {"Benchmark Duration" : <24}{cfg["duration"]} seconds
        {"Benchmark CPUs" : <24}{"All" if cfg["custom_cores"] == [] else str(cfg["custom_cores"])[1:-1].replace(" ", "")}
        {"Estimated Time" : <24}{round((cpu_count * (10 + (7 * cfg['afterburner_profile']) + cfg["cache_duration"] + (cfg["duration"] + 5)))/60)} minutes approx
        {"Load Afterburner" : <24}{bool(cfg["afterburner_profile"])} {f"(profile {cfg['afterburner_profile']})" if cfg["afterburner_profile"] else ""}
        {"DPC/ISR Logging" : <24}{bool(cfg['dpcisr'])}
        {"Save ETLs" : <24}{bool(cfg["save_etls"])}
        {"Colored Output" : <24}{bool(cfg["colored_output"])}
        {"Fullscreen" : <24}{bool(cfg["fullscreen"])} ({f"{user32.GetSystemMetrics(0)}x{user32.GetSystemMetrics(1)}" if cfg["fullscreen"] else f"{cfg['x_res']}x{cfg['y_res']}"})
        {"Sync Affinity" : <24}{bool(cfg["sync_liblava_affinity"])}
    """

    print(textwrap.dedent(runtime_info))
    input("info: press enter to start benchmarking...")

    print("info: generating and preparing prerequisites")
    create_lava_cfg(cfg["fullscreen"], cfg["x_res"], cfg["y_res"])
    os.mkdir(output_path)
    os.mkdir(f"{output_path}\\CSVs")

    if cfg["maximum"]:
        master_table[0].append("Max")

    if cfg["avgerage"]:
        master_table[0].append("Avg")

    if cfg["minimum"]:
        master_table[0].append("Min")

    if cfg["stdev"]:
        master_table[0].append("STDEV")

    if cfg["percentile"]:
        for metric in [0.1, 0.01, 0.005]:
            master_table[0].append(f"{metric} %ile")

    if cfg["lows"]:
        for metric in [0.1, 0.01, 0.005]:
            master_table[0].append(f"{metric}% Low")

    if cfg["dpcisr"]:
        os.mkdir(f"{output_path}\\xperf")

    # stop any existing trace sessions and processes
    if cfg["dpcisr"]:
        subprocess.run([cfg["xperf_path"], "-stop"], **subprocess_null, check=False)
    kill_processes("xperf.exe", "lava-triangle.exe", present_mon)

    timer_resolution(True)

    for cpu in range(0, cpu_count + 1):
        if cfg["custom_cores"] != [] and cpu not in cfg["custom_cores"]:
            continue

        print(f"info: benchmarking CPU {cpu}")

        dec_affinity = convert_affinity(cpu)
        apply_affinity(videocontroller_hwids, 1, dec_affinity)
        time.sleep(5)

        if cfg["afterburner_profile"]:
            start_afterburner(cfg["afterburner_path"], cfg["afterburner_profile"])

        affinity_args = []
        if cfg["sync_liblava_affinity"]:
            affinity_args = ["/affinity", str(dec_affinity)]

        subprocess.run(["start", *affinity_args, "bin\\liblava\\lava-triangle.exe"], shell=True, check=False)
        time.sleep(5)

        if cfg["cache_duration"] != 0:
            time.sleep(cfg["cache_duration"])

        if cfg["dpcisr"]:
            subprocess.run([cfg["xperf_path"], "-on", "base+interrupt+dpc"], check=False)

        subprocess.Popen([
            f"bin\\PresentMon\\{present_mon}",
            "-stop_existing_session",
            "-no_top",
            "-timed", str(cfg["duration"]),
            "-process_name", "lava-triangle.exe",
            "-output_file", f"{output_path}\\CSVs\\CPU-{cpu}.csv",
        ], **subprocess_null)

        time.sleep(cfg["duration"] + 5)

        if cfg["dpcisr"]:
            subprocess.run([
                cfg["xperf_path"],
                "-d", f"{output_path}\\xperf\\CPU-{cpu}.etl"
            ], **subprocess_null, check=False)

            if not os.path.exists(f"{output_path}\\xperf\\CPU-{cpu}.etl"):
                print("error: xperf etl log unsuccessful")
                os.rmdir(output_path)
                return 1

        kill_processes("xperf.exe", "lava-triangle.exe", present_mon)

        if not os.path.exists(f"{output_path}\\CSVs\\CPU-{cpu}.csv"):
            print("error: csv log unsuccessful, this may be due to a missing dependency or windows component")
            os.rmdir(output_path)
            return 1

    for cpu in range(0, cpu_count + 1):
        if cfg["custom_cores"] != [] and cpu not in cfg["custom_cores"]:
            continue

        print(f"info: parsing data for CPU {cpu}")

        frametimes = []
        with open(f"{output_path}\\CSVs\\CPU-{cpu}.csv", "r", encoding="UTF-8") as f:
            for row in csv.DictReader(f):
                if (milliseconds := row.get("MsBetweenPresents")) is not None:
                    frametimes.append(float(milliseconds))
                elif (milliseconds := row.get("msBetweenPresents")) is not None:
                    frametimes.append(float(milliseconds))
        frametimes = sorted(frametimes, reverse=True)

        frametime_data = {}
        frametime_data["frametimes"] = frametimes
        frametime_data["min"] = min(frametimes)
        frametime_data["max"] = max(frametimes)
        frametime_data["sum"] = sum(frametimes)
        frametime_data["len"] = len(frametimes)

        fps_data = []
        fps_data.append(f"CPU {cpu}")

        for metric in ["Max", "Avg", "Min"]:
            if metric in master_table[0]:
                fps_data.append(f"{compute_frametimes(frametime_data, metric):.2f}")

        if cfg["stdev"]:
            fps_data.append(f"-{compute_frametimes(frametime_data, 'STDEV'):.2f}")

        if cfg["percentile"]:
            for value in [0.1, 0.01, 0.005]:
                fps_data.append(f"{compute_frametimes(frametime_data, 'Percentile', value):.2f}")

        if cfg["lows"]:
            for value in [0.1, 0.01, 0.005]:
                fps_data.append(f"{compute_frametimes(frametime_data, 'Lows', value):.2f}")

        master_table.append(fps_data)

        if cfg["dpcisr"]:
            subprocess.run([
                cfg["xperf_path"],
                "-quiet",
                "-i", f"{output_path}\\xperf\\CPU-{cpu}.etl",
                "-o", f"{output_path}\\xperf\\CPU-{cpu}.txt",
                "-a", "dpcisr"
                ], check=False)

            if not os.path.exists(f"{output_path}\\xperf\\CPU-{cpu}.txt"):
                print("error: unable to generate dpcisr report")
                os.rmdir(output_path)
                return 1

            if not cfg["save_etls"]:
                os.remove(f"{output_path}\\xperf\\CPU-{cpu}.etl")

    if cfg["colored_output"]:
        green = "\x1b[92m"
        default = "\x1b[0m"
        os.system("color")
    else:
        green = ""
        default = ""

    os.system("cls")
    os.system("mode 300, 1000")
    apply_affinity(videocontroller_hwids, 0)

    if os.path.exists("C:\\kernel.etl"):
        os.remove("C:\\kernel.etl")

    for column in range(1, len(master_table[0])):
        best_value = float(master_table[1][column])
        for row in range(1, len(master_table)):
            fps = float(master_table[row][column])
            if fps > best_value:
                best_value = fps

        # iterate over the entire row again and find matches
        # this way we can append a * or green text to all duplicate values
        # as it is only fair to do so
        for row in range(1, len(master_table)):
            fps = abs(float(master_table[row][column]))
            master_table[row][column] = f"{fps:.2f}"
            if fps == abs(best_value):
                new_value = f"{green}*{float(master_table[row][column]):.2f}{default}"
                master_table[row][column] = new_value

    print(textwrap.dedent(runtime_info))
    print(tabulate(master_table, headers="firstrow", tablefmt="fancy_grid", floatfmt=".2f") + "\n")

    # remove color codes from tables
    if cfg["colored_output"]:
        for outer_index, outer_value in enumerate(master_table):
            for inner_index, inner_value in enumerate(outer_value):
                if green in str(inner_value) or default in str(inner_value):
                    new_value = str(inner_value).replace(green, "").replace(default, "")
                    master_table[outer_index][inner_index] = new_value

    with open(f"{output_path}\\report.txt", "a", encoding="UTF-8") as f:
        f.write(textwrap.dedent(runtime_info) + "\n")
        f.write(tabulate(master_table, headers="firstrow", tablefmt="fancy_grid", floatfmt=".2f"))

    return 0

if __name__ == "__main__":
    sys.exit(main())
