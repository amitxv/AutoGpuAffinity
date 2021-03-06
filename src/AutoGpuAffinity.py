from __future__ import annotations
import winreg
import os
import time
import subprocess
import csv
import math
import sys
import ctypes
import platform
from tabulate import tabulate

ntdll = ctypes.WinDLL("ntdll.dll")
subprocess_null = {"stdout": subprocess.DEVNULL, "stderr": subprocess.DEVNULL}
enum_pci_path = "SYSTEM\\ControlSet001\\Enum\\PCI"


def kill_processes(*targets: str) -> None:
    """Kill windows processes"""
    for process in targets:
        subprocess.run(["taskkill", "/F", "/IM", process], **subprocess_null, check=False)


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
    elif metric == "STDEV":
        mean = frametime_data["sum"] / frametime_data["len"]
        dev = [x - mean for x in frametime_data["frametimes"]]
        dev2 = [x * x for x in dev]
        result = math.sqrt(sum(dev2) / frametime_data["len"])

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
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
            try:
                return winreg.QueryValueEx(key, value_name)[0]
            except FileNotFoundError:
                return None
    except FileNotFoundError:
        return None


def convert_affinity(cpu: int) -> int:
    """Convert CPU affinity to the decimal representation"""
    affinity = 0
    affinity |= 1 << cpu
    return affinity


def apply_affinity(instances: list, action: str, affinity: int = -1) -> None:
    """
    Apply interrupt affinity policy to graphics driver

    Accepts affinity as the decimal representation
    """
    for instance in instances:
        policy_path = f"{enum_pci_path}\\{instance}\\Device Parameters\\Interrupt Management\\Affinity Policy"
        if action == "write" and affinity > -1:
            bin_affinity = bin(affinity).replace("0b", "")
            le_hex = int(bin_affinity, 2).to_bytes(8, "little").rstrip(b"\x00")
            write_key(policy_path, "DevicePolicy", 4, 4)
            write_key(policy_path, "AssignmentSetOverride", 3, le_hex)
        elif action == "delete":
            delete_key(policy_path, "DevicePolicy")
            delete_key(policy_path, "AssignmentSetOverride")

    subprocess.run(["bin\\restart64\\restart64.exe", "/q"], check=False)


def create_lava_cfg(fullscr: bool, x_resolution: int, y_resolution: int) -> None:
    """Creates the lava-triangle configuration file"""
    lavatriangle_folder = f"{os.environ['USERPROFILE']}\\AppData\\Roaming\\liblava\\lava triangle"
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
    with open(lavatriangle_config, "a", encoding="UTF-8") as cfg_f:
        for i in lavatriangle_content:
            cfg_f.write(f"{i}\n")


def start_afterburner(path: str, profile: int) -> None:
    """Starts afterburner and loads a profile"""
    subprocess.Popen([path, f"-Profile{profile}"])
    time.sleep(7)
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

    ntdll.NtQueryTimerResolution(ctypes.byref(min_res), ctypes.byref(max_res), ctypes.byref(curr_res))

    if max_res.value <= 10000 and ntdll.NtSetTimerResolution(10000, int(enabled), ctypes.byref(curr_res)) == 0:
        return 0
    return 1


def gpu_instance_paths() -> list:
    """Returns a list of the device instance paths for all present NVIDIA/AMD GPUs"""
    dev_inst_path = []
    # iterate over Enum\PCI\X
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, enum_pci_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as pci_keys:
        for pci_key_index in range(winreg.QueryInfoKey(pci_keys)[0]):
            pci_subkeys = winreg.EnumKey(pci_keys, pci_key_index)

            # iterate over Enum\PCI\X\Y
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, f"{enum_pci_path}\\{pci_subkeys}", 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as pci_subkey:
                for pci_subkey_index in range(winreg.QueryInfoKey(pci_subkey)[0]):
                    sub_keys = f"{pci_subkeys}\\{winreg.EnumKey(pci_subkey, pci_subkey_index)}"

                    # read DeviceDesc inside Enum\PCI\X\Y
                    driver_desc = read_value(f"{enum_pci_path}\\{sub_keys}", "DeviceDesc")
                    if driver_desc is not None:
                        driver_desc = str(driver_desc).upper()
                        if ("NVIDIA_DEV" in driver_desc or "NVIDIA GRAPHICS" in driver_desc) or ("AMD" in driver_desc and "RADEON" in driver_desc):
                            dev_inst_path.append(sub_keys)
    return dev_inst_path


def parse_config(config_path: str) -> dict:
    """Parse a simple configuration file and return a dict of the settings/values"""
    config = {}
    with open(config_path, "r", encoding="UTF-8") as cfg_f:
        for line in cfg_f:
            if "//" not in line:
                line = line.strip("\n")
                setting, _equ, value = line.rpartition("=")
                if setting != "" and value != "":
                    config[setting] = value
    return config


def main() -> int:
    """CLI Entrypoint"""
    version = "0.11.2"

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
    sync_liblava_affinity = bool(int(config["sync_liblava_affinity"]))

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

    has_afterburner = (1 <= afterburner_profile <= 5) and os.path.exists(afterburner_path)

    seconds_per_trial = 10 + (7 * has_afterburner) + (cache_trials + trials) * (duration + 5)
    estimated_time = seconds_per_trial * (total_cpus if custom_cores == [] else len(custom_cores))

    os.makedirs("captures", exist_ok=True)
    output_path = f"captures\\AutoGpuAffinity-{time.strftime('%d%m%y%H%M%S')}"
    runtime_info = f"""
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
    print(runtime_info)
    input("    Press enter to start benchmarking...\n")

    print("info: creating liblava config file")
    create_lava_cfg(fullscreen, x_res, y_res)

    if timer_resolution(True) != 0:
        print("info: unable to set timer-resolution")

    os.mkdir(output_path)
    os.mkdir(f"{output_path}\\CSVs")

    main_table = []
    main_table.append([
        "", "Max", "Avg", "Min", "STDEV",
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

        dec_affinity = convert_affinity(cpu)

        print("info: applying affinity")
        apply_affinity(instance_paths, "write", dec_affinity)
        time.sleep(5)

        if has_afterburner:
            print(f"info: loading afterburner profile {afterburner_profile}")
            start_afterburner(afterburner_path, afterburner_profile)

        affinity_args = []
        if sync_liblava_affinity:
            affinity_args = ["/affinity", str(dec_affinity)]

        subprocess.run(["start", *affinity_args, "bin\\liblava\\lava-triangle.exe"], shell=True, check=False)

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

            subprocess.Popen([
                "bin\\PresentMon\\PresentMon.exe",
                "-stop_existing_session",
                "-no_top",
                "-timed", str(duration),
                "-process_name", "lava-triangle.exe",
                "-output_file", f"{output_path}\\CSVs\\{file_name}.csv",
            ], **subprocess_null)

            time.sleep(duration + 5)
            kill_processes("PresentMon.exe")

            if not os.path.exists(f"{output_path}\\CSVs\\{file_name}.csv"):
                if has_xperf:
                    subprocess.run([
                        xperf_path,
                        "-d", f"{output_path}\\xperf\\raw\\{file_name}.etl"
                    ], **subprocess_null, check=False)
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

        raw_csvs = []
        for trial in range(1, trials + 1):
            raw_csvs.append(f"{output_path}\\CSVs\\CPU-{cpu}-Trial-{trial}.csv")

        aggregated_csv = f"{output_path}\\CSVs\\CPU-{cpu}-Aggregated.csv"
        aggregate(raw_csvs, aggregated_csv)
        if not os.path.exists(f"{output_path}\\CSVs\\CPU-{cpu}-Aggregated.csv"):
            print("error: csv aggregation unsuccessful")
            return 1

        if has_xperf:
            # merge etls
            raw_etls = []
            for trial in range(1, trials + 1):
                raw_etls.append(f"{output_path}\\xperf\\raw\\CPU-{cpu}-Trial-{trial}.etl")

            subprocess.run([
                xperf_path,
                "-merge", *raw_etls,
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
        with open(f"{output_path}\\CSVs\\CPU-{cpu}-Aggregated.csv", "r", encoding="UTF-8") as csv_f:
            for row in csv.DictReader(csv_f):
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
        for metric in ("Max", "Avg", "Min"):
            fps_data.append(f"{compute_frametimes(frametime_data, metric):.2f}")

        fps_data.append(f"-{compute_frametimes(frametime_data, 'STDEV'):.2f}")

        for metric in ("Percentile", "Lows"):
            for value in (1, 0.1, 0.01, 0.005):
                fps_data.append(f"{compute_frametimes(frametime_data, metric, value):.2f}")
        main_table.append(fps_data)

        # begin parsing dpc/isr data
        if has_xperf:
            print(f"info: cpu {cpu} - parsing dpc/isr data")
            with open(f"{output_path}\\xperf\\merged\\CPU-{cpu}-Merged.txt", "r", encoding="UTF-8") as report_f:
                report_lines = [x.strip("\n") for x in report_f]

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
                        if len(line) == 2:
                            for _ in range(line[1] + 1):
                                usec_data.append(line[0])
                        i += 1

                    length = len(usec_data)
                    dpc_isrdata = []
                    dpc_isrdata.append(f"CPU {cpu} {'DPCs' if dpcs else 'ISRs'}")
                    for metric in (95, 96, 97, 98, 99, 99.1, 99.2, 99.3, 99.4, 99.5, 99.6, 99.7, 99.8, 99.9):
                        dpc_isrdata.append(f"<={sorted(usec_data)[int(math.ceil((length * metric) / 100)) - 1]} ??s")

                    if dpcs:
                        dpc_table.append(dpc_isrdata)
                    else:
                        isr_table.append(dpc_isrdata)

    colored_output = colored_output and platform.release() != "" and int(platform.release()) >= 10

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

    if os.path.exists("C:\\kernel.etl"):
        os.remove("C:\\kernel.etl")

    timer_resolution(False)

    for column in range(1, len(main_table[0])):
        best_value = float(main_table[1][column])
        for row in range(1, len(main_table)):
            fps = float(main_table[row][column])
            if fps > best_value:
                best_value = fps

        # iterate over the entire row again and find matches
        # this way we can append a * or green text to all duplicate values
        # as it is only fair to do so
        for row in range(1, len(main_table)):
            fps = abs(float(main_table[row][column]))
            main_table[row][column] = f"{fps:.2f}"
            if fps == abs(best_value):
                new_value = f"{green}*{float(main_table[row][column]):.2f}{default}"
                main_table[row][column] = new_value

    if has_xperf:
        for table in [dpc_table, isr_table]:
            for column in range(1, len(table[0])):
                best_value = float(table[1][column].strip("<= ??s"))
                for row in range(1, len(table)):
                    usecs = float(table[row][column].strip("<= ??s"))
                    if usecs < best_value:
                        best_value = usecs

                for row in range(1, len(table)):
                    usecs = float(table[row][column].strip("<= ??s"))
                    if usecs == best_value:
                        new_value = f"{green}*<={int(table[row][column].strip('<= ??s'))} ??s{default}"
                        table[row][column] = new_value

    table_main = tabulate(main_table, headers="firstrow", tablefmt="fancy_grid", floatfmt=".2f")
    if has_xperf:
        table_dpc = tabulate(dpc_table, headers="firstrow", tablefmt="fancy_grid")
        table_isr = tabulate(isr_table, headers="firstrow", tablefmt="fancy_grid")

    frametime_analysis_url = "https://boringboredom.github.io/Frame-Time-Analysis"
    print_result_info = f"""
        > Drag and drop the aggregated CSVs into {frametime_analysis_url} for a graphical representation of the data.
        > Affinities for all GPUs have been reset to the Windows default (none).
        > Consider running this tool a few more times to see if the same core is consistently performant.
        > If you see absurdly low values for 0.005% lows, you should discard the results and re-run the tool.
        > Report.txt in the working directory contains the output above for later reference.
    """

    print(runtime_info)
    print("   FPS/frametime data:\n")
    print(table_main)

    if has_xperf:
        print("\n    DPC and ISR latency data for dxgkrnl.sys:\n")
        print(table_dpc)
        print(table_isr)

    print(print_result_info)

    # remove color codes from tables
    if colored_output:
        table_arrays = [main_table]
        if has_xperf:
            table_arrays.extend([dpc_table, isr_table])
        for array in table_arrays:
            for outer_index, outer_value in enumerate(array):
                for inner_index, inner_value in enumerate(outer_value):
                    if green in str(inner_value) or default in str(inner_value):
                        new_value = str(inner_value).replace(green, "").replace(default, "")
                        array[outer_index][inner_index] = new_value

        table_main = tabulate(main_table, headers="firstrow", tablefmt="fancy_grid", floatfmt=".2f")
        if has_xperf:
            table_dpc = tabulate(dpc_table, headers="firstrow", tablefmt="fancy_grid")
            table_isr = tabulate(isr_table, headers="firstrow", tablefmt="fancy_grid")

    with open(f"{output_path}\\report.txt", "a", encoding="UTF-8") as report_f:
        report_f.write(runtime_info)
        report_f.write("\n    FPS/frametime data:\n\n")
        report_f.write(table_main)

        if has_xperf:
            report_f.write("\n\n    DPC and ISR latency data for dxgkrnl.sys:\n\n")
            report_f.write(f"{table_dpc}\n")
            report_f.write(table_isr)

    return 0


if __name__ == "__main__":
    sys.exit(main())
