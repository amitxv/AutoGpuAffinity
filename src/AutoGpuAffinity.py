import argparse
import csv
import ctypes
import datetime
import json
import os
import shutil
import subprocess
import sys
import textwrap
import time
import traceback
import winreg
from configparser import ConfigParser
from typing import Any, Dict, List, Tuple, Union

import wmi

from computeframetimes import Fps

stdnull = {"stdout": subprocess.DEVNULL, "stderr": subprocess.DEVNULL}
program_path = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(__file__)


def create_lava_cfg(enable_fullscren: bool, x_resolution: int, y_resolution: int) -> None:
    lava_triangle_folder = f"{os.environ['APPDATA']}\\liblava\\lava triangle"
    os.makedirs(lava_triangle_folder, exist_ok=True)
    lava_triangle_config = f"{lava_triangle_folder}\\window.json"

    config_content = {
        "default": {
            "decorated": True,
            "floating": False,
            "fullscreen": enable_fullscren,
            "height": y_resolution,
            "maximized": False,
            "monitor": 0,
            "resizable": True,
            "width": x_resolution,
            "x": 0,
            "y": 0,
        }
    }

    with open(lava_triangle_config, "w", encoding="utf-8") as file:
        json.dump(config_content, file, indent=4)


def kill_processes(*targets: str) -> None:
    for process in targets:
        subprocess.run(["taskkill", "/F", "/IM", process], **stdnull, check=False)


def read_value(path: str, value_name: str) -> Union[Tuple[Any, int], None]:
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
            return winreg.QueryValueEx(key, value_name)[0]
    except FileNotFoundError:
        return None


def apply_affinity(hwids: List[str], cpu: int = -1, apply: bool = True) -> None:
    for hwid in hwids:
        policy_path = f"SYSTEM\\ControlSet001\\Enum\\{hwid}\\Device Parameters\\Interrupt Management\\Affinity Policy"
        if apply and cpu > -1:
            decimal_affinity = 1 << cpu
            bin_affinity = bin(decimal_affinity).lstrip("0b")
            le_hex = int(bin_affinity, 2).to_bytes(8, "little").rstrip(b"\x00")

            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, policy_path) as key:
                winreg.SetValueEx(key, "DevicePolicy", 0, winreg.REG_DWORD, 4)
                winreg.SetValueEx(key, "AssignmentSetOverride", 0, winreg.REG_BINARY, le_hex)

        else:
            try:
                with winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE, policy_path, 0, winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY
                ) as key:
                    winreg.DeleteValue(key, "DevicePolicy")
                    winreg.DeleteValue(key, "AssignmentSetOverride")
            except FileNotFoundError:
                pass

    subprocess.run([f"{program_path}\\bin\\restart64\\restart64.exe", "/q"], check=False)


def start_afterburner(path: str, profile: int) -> None:
    with subprocess.Popen([path, f"/Profile{profile}", "/Q"]) as process:
        time.sleep(5)
        process.kill()


def display_results(csv_directory: str, enable_color: bool) -> None:
    results: Dict[str, Dict[str, float]] = {}

    if enable_color:
        green = "\x1b[92m"
        default = "\x1b[0m"
        os.system("color")
    else:
        green = ""
        default = ""

    cpus = sorted([int(file.strip("CPU-.csv")) for file in os.listdir(csv_directory)])

    for cpu in cpus:
        csv_file = f"CPU-{cpu}.csv"

        frametimes: List[float] = []

        with open(f"{csv_directory}\\{csv_file}", "r", encoding="utf-8") as file:
            for row in csv.DictReader(file):
                # convert key names to lowercase because column names changed in a newer version of PresentMon
                row = {key.lower(): value for key, value in row.items()}

                if (ms_between_presents := row.get("msbetweenpresents")) is not None:
                    frametimes.append(float(ms_between_presents))

        fps = Fps(frametimes)

        # results of current CPU in results dict
        results[cpu] = {
            "maximum": round(fps.maximum(), 2),
            "average": round(fps.average(), 2),
            "minimum": round(fps.minimum(), 2),
            # negate positive value so that highest negative value will be the lowest absolute value
            "stdev": round(-fps.stdev(), 2),
            **{
                f"{metric}{value}": round(getattr(fps, metric)(value), 2)
                for metric in ("percentile", "lows")
                for value in (1, 0.1, 0.01, 0.005)
            },
        }

    # analyze best values for each metric
    for metric in (
        "maximum",
        "average",
        "minimum",
        "stdev",
        # "percentile1", "percentile0.1" etc
        *(tuple(f"{metric}{value}" for metric in ("percentile", "lows") for value in (1, 0.1, 0.01, 0.005))),
    ):
        first_key = next(iter(results))  # gets first key name, usually "0" for CPU 0
        best_value = results[first_key][metric]  # base value

        for _, _results in results.items():
            metric_value = _results[metric]
            if metric_value > best_value:
                best_value = metric_value

        # iterate over all values again and find matches
        for _cpu, _results in results.items():
            metric_value = _results[metric]
            # abs is for negative stdev
            # :.2f is for .00 numerical formatting
            new_value = f"{abs(metric_value):.2f}"
            # apply color if match is found
            _results[metric] = f"{green}*{new_value}{default}" if metric_value == best_value else new_value

    os.system("<nul set /p=\x1B[8;50;1000t")

    # print values to table
    print(f"{'CPU':<5}", end="")

    for metric in (
        "Max",
        "Avg",
        "Min",
        "STDEV",
        "1 %ile",
        "0.1 %ile",
        "0.01 %ile",
        "0.005 %ile",
        "1% Low",
        "0.1% Low",
        "0.01% Low",
        "0.005% Low",
    ):
        print(f"{metric:<13}", end="")

    print()

    for _cpu, _results in results.items():
        print(f"{_cpu:<5}", end="")
        for metric, metric_value in _results.items():
            ## padding needs to be larger to compensate for color chars
            right_padding = 22 if "[" in metric_value else 13
            print(f"{metric_value:<{right_padding}}", end="")
        print()

    print()


def main() -> int:
    version = "0.15.9"

    print(f"AutoGpuAffinity Version {version} - GPLv3\nGitHub - https://github.com/amitxv\n")

    if not ctypes.windll.shell32.IsUserAnAdmin():
        print("error: administrator privileges required")
        return 1

    gpu_hwids: List[str] = [gpu.PnPDeviceID for gpu in wmi.WMI().Win32_VideoController()]

    if (cpu_count := os.cpu_count()) is not None:
        cpu_count -= 1  # os.cpu_count() returns core count not last CPU index
    else:
        print("error: unable to get CPU count")
        return 1

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--version",
        action="version",
        version=f"AutoGpuAffinity v{version}",
    )
    parser.add_argument(
        "--config",
        metavar="<config>",
        type=str,
        help="path to config file",
    )
    parser.add_argument(
        "--analyze",
        metavar="<csv directory>",
        type=str,
        help="analyze csv files from a previous benchmark",
    )
    parser.add_argument(
        "--apply_affinity",
        metavar="<cpu>",
        type=str,
        help="assign a single core affinity to graphics drivers",
    )
    args = parser.parse_args()

    windows_version_info = sys.getwindowsversion()

    if args.analyze:
        display_results(args.analyze, windows_version_info.major >= 10)
        return 0

    basicdisplay_start = read_value("SYSTEM\\CurrentControlSet\\Services\\BasicDisplay", "Start")

    if basicdisplay_start == 4:
        print("error: please enable the BasicDisplay driver to prevent issues with restarting the GPU driver")
        return 1

    if args.apply_affinity:
        requested_affinity = int(args.apply_affinity)
        if not 0 <= requested_affinity <= cpu_count:
            print("error: invalid affinity")
            return 1

        apply_affinity(gpu_hwids, requested_affinity)
        print(f"info: set gpu driver affinity to: CPU {requested_affinity}")
        return 0
    # use 1.6.0 on Windows Server
    presentmon = f"PresentMon-{'1.9.0' if windows_version_info.major >= 10 and windows_version_info.product_type != 3 else '1.6.0'}-x64.exe"

    config_path = args.config if args.config is not None else f"{program_path}\\config.ini"
    user32 = ctypes.windll.user32

    subject_paths: Dict[int, str] = {
        1: f"{program_path}\\bin\\liblava\\lava-triangle.exe",
        2: f"{program_path}\\bin\\Benchmark.DirectX9.Black.White.exe",
    }

    # delimiters=("=") is required for file path errors with colons
    config = ConfigParser(delimiters="=")
    config.read(config_path)

    if not gpu_hwids:
        print("error: no graphics card found")
        return 1

    if not os.path.exists(config_path):
        print("error: config file not found")
        return 1

    if config.getint("settings", "cache_duration") < 0 or config.getint("settings", "benchmark_duration") <= 0:
        print("error: invalid durations specified")
        return 1

    if config.getboolean("xperf", "enabled") and not os.path.exists(config.get("xperf", "location")):
        print("error: invalid xperf path specified")
        return 1

    if config.getint("MSI Afterburner", "profile") > 0 and not os.path.exists(
        config.get("MSI Afterburner", "location")
    ):
        print("error: invalid MSI Afterburner path specified")
        return 1

    if (subject_path := subject_paths.get(config.getint("settings", "subject"))) is None:
        print("error: invalid subject specified")
        return 1

    subject_fname = os.path.basename(subject_path)

    # can't update config with list, must be a string
    custom_cpus = json.loads(config.get("settings", "custom_cpus"))

    if custom_cpus:
        # remove duplicates and sort
        benchmark_cpus = sorted(list(set(custom_cpus)))

        if not all(0 <= cpu <= cpu_count for cpu in benchmark_cpus):
            print("error: invalid cpus in custom_cpus array")
            return 1
    else:
        benchmark_cpus = list(range(cpu_count + 1))

    session_directory = f"{program_path}\\captures\\AutoGpuAffinity-{time.strftime('%d%m%y%H%M%S')}"
    estimated_time_seconds = (
        10
        + config.getint("settings", "cache_duration")
        + config.getint("settings", "benchmark_duration")
        + (5 if config.getint("MSI Afterburner", "profile") > 0 else 0)
    ) * len(benchmark_cpus)

    estimated_time = datetime.timedelta(seconds=estimated_time_seconds)
    finish_time = datetime.datetime.now() + estimated_time

    print(
        textwrap.dedent(
            f"""        Session Directory        {session_directory}
        Cache Duration           {config.get("settings", "cache_duration")}
        Benchmark Duration       {config.get("settings", "benchmark_duration")}
        Benchmark CPUs           {"All" if not custom_cpus else ','.join([str(cpu) for cpu in benchmark_cpus])}
        Subject                  {os.path.splitext(subject_fname)[0]}
        Estimated Time           {estimated_time}
        Estimated End Time       {finish_time.strftime('%H:%M:%S')}
        Load Afterburner         {config.getint("MSI Afterburner", "profile") > 0}
        DPC/ISR Logging          {config.getboolean("xperf", "enabled")}
        Save ETLs                {config.getboolean("xperf", "save_etls")}
        Window Mode              {f"Fullscreen ({user32.GetSystemMetrics(0)}x{user32.GetSystemMetrics(1)})" if config.getboolean("liblava", "fullscreen") else f"Windowed ({config.get('liblava', 'x_resolution')}x{config.get('liblava', 'y_resolution')})"}
        Sync Affinity            {config.getboolean("settings", "sync_driver_affinity")}
        """
        )
    )

    if not config.getboolean("settings", "skip_confirmation"):
        input("info: press enter to start benchmarking...")

    if config.getint("settings", "subject") == 1:
        create_lava_cfg(
            config.getboolean("liblava", "fullscreen"),
            config.getint("liblava", "x_resolution"),
            config.getint("liblava", "y_resolution"),
        )

    # this will create all of the required folders
    os.makedirs(f"{session_directory}\\CSVs", exist_ok=True)

    # stop any existing trace sessions and processes
    if config.getboolean("xperf", "enabled"):
        os.mkdir(f"{session_directory}\\xperf")
        subprocess.run([config.get("xperf", "location"), "-stop"], **stdnull, check=False)

    kill_processes("xperf.exe", subject_fname, presentmon)

    for cpu in benchmark_cpus:
        print(f"info: benchmarking CPU {cpu}")

        apply_affinity(gpu_hwids, cpu)
        time.sleep(5)

        if (profile := config.getint("MSI Afterburner", "profile")) > 0:
            start_afterburner(config.get("MSI Afterburner", "location"), profile)

        affinity_args: List[str] = []
        if config.getboolean("settings", "sync_driver_affinity"):
            affinity_args.extend(["/affinity", str(1 << cpu)])

        subprocess.run(
            ["start", "", *affinity_args, subject_path],
            shell=True,
            check=False,
        )

        # 5s offset to allow subject to launch
        time.sleep(5 + config.getint("settings", "cache_duration"))

        if config.getboolean("xperf", "enabled"):
            subprocess.run([config.get("xperf", "location"), "-on", "base+interrupt+dpc"], check=False)

        subprocess.run(
            [
                f"{program_path}\\bin\\PresentMon\\{presentmon}",
                "-stop_existing_session",
                "-no_top",
                "-timed",
                config.get("settings", "benchmark_duration"),
                "-process_name",
                subject_fname,
                "-output_file",
                f"{session_directory}\\CSVs\\CPU-{cpu}.csv",
                "-terminate_after_timed",
            ],
            **stdnull,
            check=False,
        )

        if not os.path.exists(f"{session_directory}\\CSVs\\CPU-{cpu}.csv"):
            print("error: csv log unsuccessful, this may be due to a missing dependency or windows component")
            shutil.rmtree(session_directory)
            apply_affinity(gpu_hwids, apply=False)
            return 1

        if config.getboolean("xperf", "enabled"):
            subprocess.run(
                [config.get("xperf", "location"), "-d", f"{session_directory}\\xperf\\CPU-{cpu}.etl"],
                **stdnull,
                check=False,
            )

            with subprocess.Popen(
                [
                    config.get("xperf", "location"),
                    "-quiet",
                    "-i",
                    f"{session_directory}\\xperf\\CPU-{cpu}.etl",
                    "-o",
                    f"{session_directory}\\xperf\\CPU-{cpu}.txt",
                    "-a",
                    "dpcisr",
                ]
            ) as process:
                process.wait()
                if process.returncode != 0:
                    print("error: unable to generate dpcisr report")
                    shutil.rmtree(session_directory)
                    apply_affinity(gpu_hwids, apply=False)
                    return 1

            if not config.getboolean("xperf", "save_etls"):
                os.remove(f"{session_directory}\\xperf\\CPU-{cpu}.etl")

        kill_processes("xperf.exe", subject_fname, presentmon)

    # cleanup
    apply_affinity(gpu_hwids, apply=False)

    if os.path.exists("C:\\kernel.etl"):
        os.remove("C:\\kernel.etl")

    print()
    display_results(f"{session_directory}\\CSVs", windows_version_info.major >= 10)

    return 0


if __name__ == "__main__":
    __exit_code__ = 0
    try:
        __exit_code__ = main()
    except KeyboardInterrupt:
        sys.exit(1)
    except Exception:
        print(traceback.format_exc())
        __exit_code__ = 1
    finally:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        process_array = (ctypes.c_uint * 1)()
        num_processes = kernel32.GetConsoleProcessList(process_array, 1)
        # only pause if script was ran by double-clicking
        if num_processes < 3:
            input("info: press enter to exit")

        sys.exit(__exit_code__)
