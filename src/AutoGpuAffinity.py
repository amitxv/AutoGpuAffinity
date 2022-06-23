from __future__ import annotations
import winreg
import os
import time
import subprocess
import csv
import math
import sys
import ctypes
from tabulate import tabulate

ntdll = ctypes.WinDLL("ntdll.dll")
subprocess_null = {"stdout": subprocess.DEVNULL, "stderr": subprocess.DEVNULL}
enum_pci_path = "SYSTEM\\ControlSet001\\Enum\\PCI"


def kill_processes(*targets: str) -> None:
    """Kill windows processes"""
    for process in targets:
        subprocess.call(["taskkill", "/F", "/IM", process], **subprocess_null)


def compute_frametimes(frametime_data: dict, metric: str, value: float = -1) -> float:
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


def read_value(path: str, value_name: str) -> list | None:
    """Read keys in Windows Registry"""
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY
        ) as key:
            try:
                return winreg.QueryValueEx(key, value_name)[0]
            except FileNotFoundError:
                return None
    except FileNotFoundError:
        return None


def apply_affinity(instances: list, action: str, thread: int = -1) -> None:
    """Apply interrupt affinity policy to graphics driver"""
    for instance in instances:
        policy_path = f"{enum_pci_path}\\{instance}\\Device Parameters\\Interrupt Management\\Affinity Policy"
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


def create_lava_cfg(fullscr: bool, x_resolution: int, y_resolution: int) -> None:
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
        f'        "fullscreen": {"true" if fullscr else "false"},',
        f'        "height": {y_resolution},',
        '        "maximized": false,',
        '        "monitor": 0,',
        '        "resizable": true,',
        f'        "width": {x_resolution},',
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


def timer_resolution(enabled: bool) -> int:
    """
    Sets the kernel timer-resolution to 1ms

    This function does not affect other processes on Windows 10 2004+
    """
    min_res = ctypes.c_ulong()
    max_res = ctypes.c_ulong()
    curr_res = ctypes.c_ulong()

    ntdll.NtQueryTimerResolution(
        ctypes.byref(min_res), ctypes.byref(max_res), ctypes.byref(curr_res)
    )

    if max_res.value <= 10000 and ntdll.NtSetTimerResolution(
        10000, int(enabled), ctypes.byref(curr_res)
    ) == 0:
        return 0
    else:
        return 1


def gpu_instance_paths() -> list:
    """Returns a list of the device instance paths for all present NVIDIA/AMD GPUs"""
    dev_inst_path = []
    # iterate over Enum\PCI\X
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, enum_pci_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as pci:
        for a in range(winreg.QueryInfoKey(pci)[0]):
            pci_subkeys = winreg.EnumKey(pci, a)

            # iterate over Enum\PCI\X\Y
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, f"{enum_pci_path}\\{pci_subkeys}", 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as subkey:
                for b in range(winreg.QueryInfoKey(subkey)[0]):
                    sub_keys = f"{pci_subkeys}\\{winreg.EnumKey(subkey, b)}"

                    # read DeviceDesc inside Enum\PCI\X\Y
                    driver_desc = read_value(f"{enum_pci_path}\\{sub_keys}", "DeviceDesc")
                    if driver_desc is not None:
                        driver_desc = str(driver_desc).upper()
                        if "NVIDIA_DEV" in driver_desc or ("AMD" in driver_desc and "RADEON" in driver_desc):
                            dev_inst_path.append(sub_keys)
    return dev_inst_path


def parse_config(config_path: str) -> dict:
    """Parse a simple configuration file and return a dict of the settings/values"""
    config = {}
    with open(config_path, "r", encoding="UTF-8") as config_file:
        for line in config_file:
            if "//" not in line:
                line = line.strip("\n")
                setting, _equ, value = line.rpartition("=")
                if setting != "" and value != "":
                    config[setting] = value
    return config
    

def main() -> int:
    """CLI Entrypoint"""
    version = "0.10.0"

    # change directory to location of program
    program_path = ""
    if getattr(sys, "frozen", False):
        program_path = os.path.dirname(sys.executable)
    elif __file__:
        program_path = os.path.dirname(__file__)
    os.chdir(program_path)

    config = parse_config("config.txt")

    trials = int(config["trials"])
    duration = int(config["duration"])
    cache_trials = int(config["cache_trials"])
    dpcisr = bool(int(config["dpcisr"]))
    xperf_path = str(config["xperf_path"])
    save_etls = bool(int(config["save_etls"]))
    afterburner_profile = int(config["afterburner_profile"])
    afterburner_path = str(config["afterburner_path"])
    custom_cores = str(config["custom_cores"])
    colored_output = bool(int(config["colored_output"]))
    fullscreen = bool(int(config["fullscreen"]))
    x_res = int(config["x_res"])
    y_res = int(config["y_res"])

    if (total_cpus := os.cpu_count()) is None:
        print("error: unable to get cpu count")
        return 1

    if (trials <= 0) or (cache_trials < 0) or (duration <= 0):
        print("error: invalid trials, cache_trials or duration in config")
        return 1

    if custom_cores.startswith("[") and custom_cores.endswith("]"):
        # strip [] and remove white spaces from string then split values into list
        custom_cores = custom_cores[1:-1].replace(" ", "").split(",")
        # remove duplicates in list
        custom_cores = list(dict.fromkeys(custom_cores))
        # convert contents of list into list[int]
        custom_cores = [int(x) for x in custom_cores if x != ""]
        # sort list in ascending order
        custom_cores.sort()
        if custom_cores != []:
            for i in custom_cores:
                if not 0 <= i <= (total_cpus - 1):
                    print("error: invalid custom_cores value in config")
                    return 1
    else:
        print("error: surrounding brackets for custom_cores value not found")
        return 1

    if (instance_paths := gpu_instance_paths()) == []:
        print("error: no graphics card found")
        return 1

    has_xperf = dpcisr == 1 and os.path.exists(xperf_path)

    has_afterburner = (1 <= afterburner_profile <= 5) and (os.path.exists(afterburner_path))

    seconds_per_trial = 10 + (7 * has_afterburner) + (cache_trials + trials) * (duration + 5)
    estimated_time = seconds_per_trial * (total_cpus if custom_cores == [] else len(custom_cores))

    os.makedirs("captures", exist_ok=True)
    output_path = f"captures\\AutoGpuAffinity-{time.strftime('%d%m%y%H%M%S')}"
    print_info = f"""
    AutoGpuAffinity v{version} Command Line

        Trials: {trials}
        Trial Duration: {duration} sec
        Benchmark CPUs: {"All" if custom_cores == [] else str(custom_cores).strip("[]")}
        Total CPUs: {total_cpus - 1}
        Log dpc/isr with xperf: {has_xperf}
        Load MSI Afterburner : {has_afterburner}
        Cache trials: {cache_trials}
        Time for completion: {estimated_time/60:.2f} min
        Session Working directory: \\{output_path}\\
        Fullscreen: {fullscreen} {f"({x_res}x{y_res})" if not fullscreen else ""}
    """
    print(print_info)
    input("    Press enter to start benchmarking...\n")

    print("info: creating liblava config file")
    create_lava_cfg(fullscreen, x_res, y_res)

    if timer_resolution(True) != 0:
        print("info: unable to set timer-resolution")

    os.mkdir(output_path)
    os.mkdir(f"{output_path}\\CSVs")

    main_table = []
    main_table.append([
        "", "Max", "Avg", "Min",
        "1 %ile", "0.1 %ile", "0.01 %ile", "0.005 %ile",
        "1% Low", "0.1% Low", "0.01% Low", "0.005% Low"
    ])

    # kill all processes before loop and prepare xperf related data
    if has_xperf:
        os.mkdir(f"{output_path}\\xperf")
        os.mkdir(f"{output_path}\\xperf\\merged")
        os.mkdir(f"{output_path}\\xperf\\raw")

        dpc_table = []
        dpc_table.append([
            "", "95 %ile", "96 %ile", "97 %ile", "98 %ile", "99 %ile",
            "99.1 %ile", "99.2 %ile", "99.3 %ile", "99.4 %ile", "99.5 %ile", "99.6 %ile",
            "99.7 %ile", "99.8 %ile", "99.9 %ile"
        ])
        isr_table = dpc_table.copy()

        subprocess.run([xperf_path, "-stop"], **subprocess_null, check=False)

    kill_processes("xperf.exe", "lava-triangle.exe", "PresentMon.exe")

    for cpu in range(0, total_cpus):
        if custom_cores != [] and cpu not in custom_cores:
            continue

        print("info: applying affinity")
        apply_affinity(instance_paths, "write", cpu)
        time.sleep(5)

        if has_afterburner:
            print(f"info: loading afterburner profile {afterburner_profile}")
            start_afterburner(afterburner_path, afterburner_profile)

        subprocess.Popen(["bin\\liblava\\lava-triangle.exe"], **subprocess_null)
        time.sleep(5)

        if cache_trials > 0:
            for trial in range(1, cache_trials + 1):
                print(f"info: cpu {cpu} - cache trial: {trial}/{cache_trials}")
                time.sleep(duration + 5)

        for trial in range(1, trials + 1):
            file_name = f"CPU-{cpu}-Trial-{trial}"
            print(f"info: cpu {cpu} - recording trial: {trial}/{trials}")

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
                if has_xperf:
                    subprocess.run([xperf_path, "-stop"], **subprocess_null, check=False)
                kill_processes("xperf.exe", "lava-triangle.exe", "PresentMon.exe")
                print("error: csv log unsuccessful, this is due to a missing dependency/ windows component")
                return 1

            if has_xperf:
                subprocess.run([
                    xperf_path,
                    "-d", f"{output_path}\\xperf\\raw\\{file_name}.etl"
                ], **subprocess_null, check=False)

                if not os.path.exists(f"{output_path}\\xperf\\raw\\{file_name}.etl"):
                    kill_processes("xperf.exe", "lava-triangle.exe", "PresentMon.exe")
                    print("error: xperf etl log unsuccessful")
                    return 1

        kill_processes("xperf.exe", "lava-triangle.exe", "PresentMon.exe")

    print("info: begin parsing data, this may take a few minutes...")
    for cpu in range(0, total_cpus):
        if custom_cores != [] and cpu not in custom_cores:
            continue

        # begin aggregating CSVs and ETLs
        print(f"info: cpu {cpu} - aggregating frametime data")

        CSVs = []
        for trial in range(1, trials + 1):
            CSVs.append(f"{output_path}\\CSVs\\CPU-{cpu}-Trial-{trial}.csv")

        aggregated_csv = f"{output_path}\\CSVs\\CPU-{cpu}-Aggregated.csv"
        aggregate(CSVs, aggregated_csv)
        if not os.path.exists(f"{output_path}\\CSVs\\CPU-{cpu}-Aggregated.csv"):
            print("error: csv aggregation unsuccessful")
            return 1

        if has_xperf:
            # merge etls
            ETLs = []
            for trial in range(1, trials + 1):
                ETLs.append(f"{output_path}\\xperf\\raw\\CPU-{cpu}-Trial-{trial}.etl")

            subprocess.run([
                xperf_path,
                "-merge", *ETLs,
                f"{output_path}\\xperf\\merged\\CPU-{cpu}-Merged.etl"
            ], **subprocess_null, check=False)

            if not os.path.exists(f"{output_path}\\xperf\\merged\\CPU-{cpu}-Merged.etl"):
                print("error: etl merge unsuccessful")
                return 1

            # generate a report based on the merged etl
            subprocess.run([
                xperf_path,
                "-quiet",
                "-i", f"{output_path}\\xperf\\merged\\CPU-{cpu}-Merged.etl",
                "-o", f"{output_path}\\xperf\\merged\\CPU-{cpu}-Merged.txt",
                "-a", "dpcisr"
                ], check=False)

            if not os.path.exists(f"{output_path}\\xperf\\merged\\CPU-{cpu}-Merged.txt"):
                print("error: unable to generate dpcisr report")
                return 1

            if not save_etls:
                os.remove(f"{output_path}\\xperf\\merged\\CPU-{cpu}-Merged.etl")
                for trial in range(1, trials + 1):
                    os.remove(f"{output_path}\\xperf\\raw\\CPU-{cpu}-Trial-{trial}.etl")

        # begin parsing frametime data
        print(f"info: cpu {cpu} - parsing frametime data")

        frametimes = []
        with open(
            f"{output_path}\\CSVs\\CPU-{cpu}-Aggregated.csv", "r", encoding="UTF-8"
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

        fps_data = []
        fps_data.append(f"CPU {cpu}")
        for metric in ("Max", "Avg", "Min"):
            fps_data.append(f"{compute_frametimes(frametime_data, metric):.2f}")

        for metric in ("Percentile", "Lows"):
            for value in (1, 0.1, 0.01, 0.005):
                fps_data.append(f"{compute_frametimes(frametime_data, metric, value):.2f}")
        main_table.append(fps_data)

        # begin parsing dpc/isr data
        if has_xperf:
            print(f"info: cpu {cpu} - parsing dpc/isr data")
            with open(
                f"{output_path}\\xperf\\merged\\CPU-{cpu}-Merged.txt", "r", encoding="UTF-8"
            ) as report:
                report_lines = [x.strip("\n") for x in report]

            dpcs = 0
            for i in range(len(report_lines)):
                if "for module dxgkrnl.sys" in report_lines[i]:
                    usec_data = []
                    dpcs = not dpcs
                    i += 1
                    while "Total," not in report_lines[i]:
                        line = report_lines[i]
                        line = line.replace(" ", "")
                        line = line.strip("ElapsedTime,>")
                        line = line.replace("AND<=", ",")
                        line = line.replace("usecs", "")
                        line = line.split(",")[1:-1]
                        # convert to int
                        line = [int(x) for x in line]
                        for _ in range(line[1] + 1):
                            usec_data.append(line[0])
                        i += 1

                    length = len(usec_data)
                    dpc_isrdata = []
                    dpc_isrdata.append(f"CPU {cpu} {'DPCs' if dpcs else 'ISRs'}")
                    for metric in (95, 96, 97, 98, 99, 99.1, 99.2, 99.3, 99.4, 99.5, 99.6, 99.7, 99.8, 99.9):
                        dpc_isrdata.append(f"<={sorted(usec_data)[int(math.ceil((length * metric) / 100)) - 1]} μs")
                    
                    dpc_table.append(dpc_isrdata) if dpcs else isr_table.append(dpc_isrdata)

    # usually gets created with xperf -stop
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
    apply_affinity(instance_paths, "delete")

    timer_resolution(False)

    for column in range(1, len(main_table[0])):
        highest_fps = 0
        for row in range(1, len(main_table)):
            fps = float(main_table[row][column])
            if fps > highest_fps:
                highest_fps = fps

        # iterate over the entire row again and find matches
        # this way we can append a * or green text to all duplicate values
        # as it is only fair to do so
        for row in range(1, len(main_table)):
            fps = float(main_table[row][column])
            if fps == highest_fps:
                new_value = f"{green}*{float(main_table[row][column]):.2f}{default}"
                main_table[row][column] = new_value

    if has_xperf:
        for table in [dpc_table, isr_table]:
            for column in range(1, len(table[0])):
                lowest_usecs = 9999
                for row in range(1, len(table)):
                    usecs = float(table[row][column].strip("<= μs"))
                    if usecs < lowest_usecs:
                        lowest_usecs = usecs

                for row in range(1, len(table)):
                    usecs = float(table[row][column].strip("<= μs"))
                    if usecs == lowest_usecs:
                        new_value = f"{green}*<={int(table[row][column].strip('<= μs'))} μs{default}"
                        table[row][column] = new_value

    frametime_analysis_url = "https://boringboredom.github.io/Frame-Time-Analysis"
    print_result_info = f"""
        > Drag and drop the aggregated CSVs into {frametime_analysis_url} for a graphical representation of the data.
        > Affinities for all GPUs have been reset to the Windows default (none).
        > Consider running this tool a few more times to see if the same core is consistently performant.
        > If you see absurdly low values for 0.005% lows, you should discard the results and re-run the tool.
    """

    print(print_info)
    print("   FPS/frametime data:\n")
    print(tabulate(main_table, headers="firstrow", tablefmt="fancy_grid", floatfmt=".2f"))

    if has_xperf:
        print("\n    DPC and ISR latency data for dxgkrnl.sys:\n")
        print(tabulate(dpc_table, headers="firstrow", tablefmt="fancy_grid"))
        print(tabulate(isr_table, headers="firstrow", tablefmt="fancy_grid"))

    print(print_result_info)

    return 0


if __name__ == "__main__":
    sys.exit(main())
