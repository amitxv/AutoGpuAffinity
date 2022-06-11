from __future__ import annotations
import winreg
import os
import time
import subprocess
import csv
import math
import sys
from tabulate import tabulate
import psutil
import wmi

gpu_info = wmi.WMI().Win32_VideoController()
subprocess_null = {"stdout": subprocess.DEVNULL, "stderr": subprocess.DEVNULL}


def kill_processes(*targets: str) -> None:
    """Kill windows processes"""
    for process in psutil.process_iter():
        if process.name() in targets:
            process.kill()


def calc(frametime_data: dict, metric: str, value: float = -1) -> float:
    """Calculate various metrics based on framedata"""
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
    return 1000 / result


def write_key(path: str, value_name: str, data_type: int, value_data: int | bytes) -> None:
    """Write keys to Windows Registry"""
    with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
        winreg.SetValueEx(key, value_name, 0, data_type, value_data)  # type: ignore


def delete_key(path: str, value_name: str) -> None:
    """Delete keys in Windows Registry"""
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY) as key:
            try:
                winreg.DeleteValue(key, value_name)
            except FileNotFoundError:
                pass
    except FileNotFoundError:
        pass


def apply_affinity(action: str, thread: int = -1) -> None:
    """Apply interrupt affinity policy to graphics driver"""
    for item in gpu_info:
        policy_path = f"SYSTEM\\ControlSet001\\Enum\\{item.PnPDeviceID}\\Device Parameters\\Interrupt Management\\Affinity Policy"
        if action == "write" and thread > -1:
            dec_affinity = 0
            dec_affinity |= 1 << thread
            bin_affinity = bin(dec_affinity).replace("0b", "")
            le_hex = int(bin_affinity, 2).to_bytes(8, "little").rstrip(b"\x00")
            write_key(policy_path, "DevicePolicy", 4, 4)
            write_key(policy_path, "AssignmentSetOverride", 3, le_hex)
        elif action == "delete":
            delete_key(policy_path, "DevicePolicy")
            delete_key(policy_path, "AssignmentSetOverride")

    subprocess.run(["bin\\restart64\\restart64.exe", "/q"], check=False)


def create_lava_cfg() -> None:
    """Creates the lava-triangle configuration file"""
    lavatriangle_folder = (f"{os.environ['USERPROFILE']}\\AppData\\Roaming\\liblava\\lava triangle")
    os.makedirs(lavatriangle_folder, exist_ok=True)
    lavatriangle_config = f"{lavatriangle_folder}\\window.json"

    if os.path.exists(lavatriangle_config):
        os.remove(lavatriangle_config)

    lavatriangle_content = [
        "{",
        '    "default": {',
        '        "decorated": true,',
        '        "floating": false,',
        '        "fullscreen": true,',
        '        "height": 1080,',
        '        "maximized": false,',
        '        "monitor": 0,',
        '        "resizable": true,',
        '        "width": 1920,',
        '        "x": 0,',
        '        "y": 0',
        "    }",
        "}",
    ]
    with open(lavatriangle_config, "a", encoding="UTF-8") as f:
        for i in lavatriangle_content:
            f.write(f"{i}\n")


def start_afterburner(path: str, profile: int) -> None:
    """Starts afterburner and loads a profile"""
    print(f"loading afterburner profile {profile}")
    try:
        subprocess.run([path, f"-Profile{profile}"], timeout=7, check=False)
    except subprocess.TimeoutExpired:
        pass
    kill_processes("MSIAfterburner.exe")


def aggregate(files: list, output_file: str) -> None:
    """Aggregates PresentMon CSV files"""
    aggregated = []
    for file in files:
        with open(file, "r", encoding="UTF-8") as csv_f:
            lines = csv_f.readlines()
            aggregated.extend(lines)

    with open(output_file, "a", encoding="UTF-8") as csv_f:
        column_names = aggregated[0]
        csv_f.write(column_names)

        for line in aggregated:
            if line != column_names:
                csv_f.write(line)


def main() -> int:
    """CLI Entrypoint"""
    version = "0.6.1"

    # change directory to location of program
    program_path = ""
    if getattr(sys, 'frozen', False):
        program_path = os.path.dirname(sys.executable)
    elif __file__:
        program_path = os.path.dirname(__file__)
    os.chdir(program_path)

    config = {}
    with open("config.txt", "r", encoding="UTF-8") as f:
        for line in f:
            if "//" not in line:
                line = line.strip("\n")
                setting, _equ, value = line.rpartition("=")
                if setting != "" and value != "":
                    config[setting] = value

    trials = int(config["trials"])
    duration = int(config["duration"])
    dpcisr = int(config["dpcisr"])
    xperf_path = str(config["xperf_path"])
    cache_trials = int(config["cache_trials"])
    afterburner_path = str(config["afterburner_path"])
    afterburner_profile = int(config["afterburner_profile"])
    custom_cores = str(config["custom_cores"])
    colored_output = int(config["colored_output"])

    total_cpus = psutil.cpu_count()

    if trials <= 0 or cache_trials < 0 or duration <= 0:
        print("invalid trials, cache_trials or duration in config")
        return 1

    if custom_cores.startswith("[") and custom_cores.endswith("]"):
        custom_cores = custom_cores[1:-1].replace(" ", "").split(",")
        if custom_cores != [""]:
            custom_cores = list(dict.fromkeys(custom_cores))
            for i in custom_cores:
                if not 0 <= int(i) <= total_cpus:
                    print("invalid custom_cores value in config")
                    return 1
    else:
        print("surrounding brackets for custom_cores value not found")
        return 1

    has_xperf = dpcisr != 0 and os.path.exists(xperf_path)

    has_afterburner = 1 <= afterburner_profile <= 5 and os.path.exists(afterburner_path)

    seconds_per_trial = 10 + (7 if has_afterburner else 0) + (cache_trials + trials) * (duration + 5)
    estimated_time = seconds_per_trial * (total_cpus if custom_cores == [""] else len(custom_cores))

    os.makedirs("captures", exist_ok=True)
    output_path = f"captures\\AutoGpuAffinity-{time.strftime('%d%m%y%H%M%S')}"
    print_info = f"""
    AutoGpuAffinity v{version} Command Line

        Trials: {trials}
        Trial Duration: {duration} sec
        Benchmark CPUs: {"All" if custom_cores == [""] else ",".join(custom_cores)}
        Total CPUs: {total_cpus - 1}
        Hyperthreading: {total_cpus > psutil.cpu_count(logical=False)}
        Log dpc/isr with xperf: {has_xperf}
        Load MSI Afterburner : {has_afterburner}
        Cache trials: {cache_trials}
        Time for completion: {estimated_time/60:.2f} min
        Session Working directory: \\{output_path}\\
    """
    print(print_info)
    input("    Press enter to start benchmarking...\n")
    create_lava_cfg()

    os.mkdir(output_path)
    os.mkdir(f"{output_path}\\CSVs")
    if has_xperf:
        os.mkdir(f"{output_path}\\xperf")

    main_table = []
    main_table.append([
        "", "Max", "Avg", "Min",
        "1 %ile", "0.1 %ile", "0.01 %ile", "0.005 %ile",
        "1% Low", "0.1% Low", "0.01% Low", "0.005% Low"
    ])

    # kill all processes before loop
    if has_xperf:
        subprocess.run([xperf_path, "-stop"], **subprocess_null, check=False)
    kill_processes("xperf.exe", "lava-triangle.exe", "PresentMon.exe")

    for active_thread in range(0, total_cpus):

        if custom_cores != [""] and str(active_thread) not in custom_cores:
            continue

        apply_affinity("write", active_thread)
        time.sleep(5)

        if has_afterburner:
            start_afterburner(afterburner_path, afterburner_profile)

        subprocess.Popen(["bin\\liblava\\lava-triangle.exe"], **subprocess_null)
        time.sleep(5)

        if cache_trials > 0:
            for trial in range(1, cache_trials + 1):
                print(f"CPU {active_thread} - Cache Trial: {trial}/{cache_trials}")
                time.sleep(duration + 5)

        for trial in range(1, trials + 1):
            file_name = f"CPU-{active_thread}-Trial-{trial}"
            print(f"CPU {active_thread} - Recording Trial: {trial}/{trials}")

            if has_xperf:
                subprocess.run([xperf_path, "-on", "base+interrupt+dpc"], check=False)

            try:
                subprocess.run([
                    "bin\\PresentMon\\PresentMon.exe",
                    "-stop_existing_session",
                    "-no_top",
                    "-verbose",
                    "-timed", str(duration),
                    "-process_name", "lava-triangle.exe",
                    "-output_file", f"{output_path}\\CSVs\\{file_name}.csv",
                    ], timeout=duration + 5, **subprocess_null, check=False)
            except subprocess.TimeoutExpired:
                pass

            if not os.path.exists(f"{output_path}\\CSVs\\{file_name}.csv"):
                kill_processes("xperf.exe", "lava-triangle.exe", "PresentMon.exe")
                print("CSV log unsuccessful, this is due to a missing dependency/ windows component")
                return 1

            if has_xperf:
                subprocess.run([xperf_path, "-stop"], **subprocess_null, check=False)
                subprocess.run([
                    xperf_path,
                    "-i", "C:\\kernel.etl",
                    "-o", f"{output_path}\\xperf\\{file_name}.txt",
                    "-a", "dpcisr"
                    ], check=False)

        kill_processes("xperf.exe", "lava-triangle.exe", "PresentMon.exe")

    for active_thread in range(0, total_cpus):

        if custom_cores != [""] and str(active_thread) not in custom_cores:
            continue

        CSVs = []
        for trial in range(1, trials + 1):
            CSVs.append(f"{output_path}\\CSVs\\CPU-{active_thread}-Trial-{trial}.csv")

        aggregated_csv = f"{output_path}\\CSVs\\CPU-{active_thread}-Aggregated.csv"
        aggregate(CSVs, aggregated_csv)

        frametimes = []
        with open(
            f"{output_path}\\CSVs\\CPU-{active_thread}-Aggregated.csv", "r", encoding="UTF-8"
        ) as f:
            for row in csv.DictReader(f):
                if row["MsBetweenPresents"] is not None:
                    frametimes.append(float(row["MsBetweenPresents"]))
        frametimes = sorted(frametimes, reverse=True)

        frametime_data = {}
        frametime_data["frametimes"] = frametimes
        frametime_data["min"] = min(frametimes)
        frametime_data["max"] = max(frametimes)
        frametime_data["sum"] = sum(frametimes)
        frametime_data["len"] = len(frametimes)

        data = []
        data.append(f"CPU {active_thread}")
        for metric in ("Max", "Avg", "Min"):
            data.append(f"{calc(frametime_data, metric):.2f}")

        for metric in ("Percentile", "Lows"):
            for value in (1, 0.1, 0.01, 0.005):
                data.append(f"{calc(frametime_data, metric, value):.2f}")
        main_table.append(data)

    if os.path.exists("C:\\kernel.etl"):
        os.remove("C:\\kernel.etl")

    if colored_output:
        green = "\x1b[92m"
        default = "\x1b[0m"
        os.system("color")
    else:
        green = ""
        default = ""

    os.system("cls")
    os.system("mode 300, 1000")
    apply_affinity("delete")

    for column in range(1, len(main_table[0])):
        highest_fps = 0
        row_index = 0
        for row in range(1, len(main_table)):
            fps = float(main_table[row][column])
            if fps > highest_fps:
                highest_fps = fps
                row_index = row
        new_value = f"{green}*{float(main_table[row_index][column]):.2f}{default}"
        main_table[row_index][column] = new_value

    print_result_info = """
        > Drag and drop the aggregated data (located in the working directory) \
into https://boringboredom.github.io/Frame-Time-Analysis for a graphical representation of the data.
        > Affinities for all GPUs have been reset to the Windows default (none).
        > Consider running this tool a few more times to see if the same core is consistently performant.
        > If you see absurdly low values for 0.005% Lows, you should discard the results and re-run the tool.
    """

    print(print_info)
    print(tabulate(main_table, headers="firstrow", tablefmt="fancy_grid", floatfmt=".2f"), "\n")
    print(print_result_info)

    return 0


if __name__ == "__main__":
    sys.exit(main())
