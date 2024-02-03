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
from typing import Any

import wmi
from compute_frametimes import Fps

program_path = (
    os.path.dirname(sys.executable)
    if getattr(sys, "frozen", False)
    else os.path.dirname(__file__)
)


def create_lava_cfg(
    enable_fullscren: bool,
    x_resolution: int,
    y_resolution: int,
) -> None:
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
        },
    }

    with open(lava_triangle_config, "w", encoding="utf-8") as file:
        json.dump(config_content, file, indent=4)


def kill_processes(*targets: str) -> None:
    for process in targets:
        subprocess.run(
            ["taskkill", "/F", "/IM", process],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )


def read_value(path: str, value_name: str) -> Any | None:
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            path,
            0,
            winreg.KEY_READ | winreg.KEY_WOW64_64KEY,
        ) as key:
            return winreg.QueryValueEx(key, value_name)[0]
    except FileNotFoundError:
        return None


def apply_affinity(hwids: list[str], cpu: int = -1, apply: bool = True) -> None:
    for hwid in hwids:
        policy_path = f"SYSTEM\\ControlSet001\\Enum\\{hwid}\\Device Parameters\\Interrupt Management\\Affinity Policy"
        if apply and cpu > -1:
            decimal_affinity = 1 << cpu
            bin_affinity = bin(decimal_affinity).lstrip("0b")
            le_hex = int(bin_affinity, 2).to_bytes(8, "little").rstrip(b"\x00")

            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, policy_path) as key:
                winreg.SetValueEx(key, "DevicePolicy", 0, winreg.REG_DWORD, 4)
                winreg.SetValueEx(
                    key,
                    "AssignmentSetOverride",
                    0,
                    winreg.REG_BINARY,
                    le_hex,
                )

        else:
            try:
                with winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE,
                    policy_path,
                    0,
                    winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY,
                ) as key:
                    winreg.DeleteValue(key, "DevicePolicy")
                    winreg.DeleteValue(key, "AssignmentSetOverride")
            except FileNotFoundError:
                pass

    subprocess.run(
        [f"{program_path}\\bin\\restart64\\restart64.exe", "/q"],
        check=False,
    )


def start_afterburner(path: str, profile: int) -> None:
    with subprocess.Popen([path, f"/Profile{profile}", "/Q"]) as process:
        time.sleep(5)
        process.kill()


def print_table(formatted_results: dict[str, dict[str, str]]):
    # print table headings
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
        print(f"{metric:<12}", end="")

    print()  # new line

    # print values for each heading
    for _cpu, _results in formatted_results.items():
        print(f"{_cpu:<5}", end="")
        for metric_value in _results.values():
            # padding needs to be larger to compensate for color chars
            right_padding = 21 if "[" in metric_value else 12
            print(f"{metric_value:<{right_padding}}", end="")
        print()  # new line

    print()  # new line


def display_results(csv_directory: str, enable_color: bool) -> None:
    results: dict[str, dict[str, float]] = {}

    # each index represents the rank (e.g. index 0 is 1st)
    colors: list[str] = [
        "\x1b[92m",  # Green
        "\x1b[93m",  # Yellow
    ]

    if enable_color:
        default = "\x1b[0m"
        os.system("color")
    else:
        default = ""

    cpus = sorted([int(file.strip("CPU-.csv")) for file in os.listdir(csv_directory)])
    num_cpus = len(cpus)
    # 1 CPUs means no ranking will be done
    # 2 CPUs means only one metric will be ranked since it's binary
    # always leave last place unranked

    top_n_values = num_cpus - 1 if num_cpus < 3 else len(colors)

    for cpu in cpus:
        csv_file = f"CPU-{cpu}.csv"

        frametimes: list[float] = []

        with open(f"{csv_directory}\\{csv_file}", encoding="utf-8") as file:
            for row in csv.DictReader(file):
                # convert key names to lowercase because column names changed in a newer version of PresentMon
                row_lower = {key.lower(): value for key, value in row.items()}

                if (
                    ms_between_presents := row_lower.get("msbetweenpresents")
                ) is not None:
                    frametimes.append(float(ms_between_presents))

        fps = Fps(frametimes)

        # results of current CPU in results dict
        results[str(cpu)] = {
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

    formatted_results: dict[str, dict[str, str]] = {cpu: {} for cpu in results}

    # analyze best values for each metric
    for metric in (
        "maximum",
        "average",
        "minimum",
        "stdev",
        # "percentile1", "percentile0.1" etc
        *(
            tuple(
                f"{metric}{value}"
                for metric in ("percentile", "lows")
                for value in (1, 0.1, 0.01, 0.005)
            )
        ),
    ):
        # set of all values within the metric
        values = {_results[metric] for _results in results.values()}

        # create ordered list without duplicates of top n values
        top_values = list(dict.fromkeys(sorted(values, reverse=True)[:top_n_values]))

        for _cpu, _results in results.items():
            metric_value = _results[metric]

            # abs is for negative values such as stdev
            # :.2f is for .00 numerical formatting
            new_value = f"{abs(metric_value):.2f}"

            # determine rank of value
            if enable_color:
                try:
                    nth_best = top_values.index(metric_value)
                    color = colors[nth_best]
                    new_value = f"{color}{new_value}{default}"
                except ValueError:
                    # don't highlight value as top n by leaving it unmodified
                    pass

            formatted_results[_cpu][metric] = new_value

    os.system("<nul set /p=\x1B[8;50;1000t")

    print_table(formatted_results)


def main() -> int:  # noqa: PLR0911, C901, PLR0912, D103, PLR0915
    version = "0.15.11"

    print(
        f"AutoGpuAffinity Version {version} - GPLv3\nGitHub - https://github.com/amitxv\nDonate - https://www.buymeacoffee.com/amitxv\n",
    )

    if not ctypes.windll.shell32.IsUserAnAdmin():
        print("error: administrator privileges required")
        return 1

    gpu_hwids: list[str] = [
        gpu.PnPDeviceID for gpu in wmi.WMI().Win32_VideoController()
    ]

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

    basicdisplay_start = read_value(
        "SYSTEM\\CurrentControlSet\\Services\\BasicDisplay",
        "Start",
    )

    if basicdisplay_start is None:
        print("error: unable to get BasicDisplay start type")
        return 1

    if int(basicdisplay_start) == 4:
        print(
            "error: please enable the BasicDisplay driver to prevent issues with restarting the GPU driver",
        )
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

    config_path = (
        args.config if args.config is not None else f"{program_path}\\config.ini"
    )
    user32 = ctypes.windll.user32

    subject_paths: dict[int, str] = {
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

    if (
        config.getint("settings", "cache_duration") < 0
        or config.getint("settings", "benchmark_duration") <= 0
    ):
        print("error: invalid durations specified")
        return 1

    if config.getboolean("xperf", "enabled") and not os.path.exists(
        config.get("xperf", "location"),
    ):
        print("error: invalid xperf path specified")
        return 1

    if config.getint("MSI Afterburner", "profile") > 0 and not os.path.exists(
        config.get("MSI Afterburner", "location"),
    ):
        print("error: invalid MSI Afterburner path specified")
        return 1

    if (
        subject_path := subject_paths.get(config.getint("settings", "subject"))
    ) is None:
        print("error: invalid subject specified")
        return 1

    subject_fname = os.path.basename(subject_path)

    # can't update config with list, must be a string
    custom_cpus = json.loads(config.get("settings", "custom_cpus"))

    if custom_cpus:
        # remove duplicates and sort
        benchmark_cpus = sorted(set(custom_cpus))

        if not all(0 <= cpu <= cpu_count for cpu in benchmark_cpus):
            print("error: invalid cpus in custom_cpus array")
            return 1
    else:
        benchmark_cpus = list(range(cpu_count + 1))

    session_directory = (
        f"{program_path}\\captures\\AutoGpuAffinity-{time.strftime('%d%m%y%H%M%S')}"
    )
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
        """,
        ),
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
        subprocess.run(
            [config.get("xperf", "location"), "-stop"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )

    kill_processes("xperf.exe", subject_fname, presentmon)

    for cpu in benchmark_cpus:
        print(f"info: benchmarking CPU {cpu}")

        apply_affinity(gpu_hwids, cpu)
        time.sleep(5)

        if (profile := config.getint("MSI Afterburner", "profile")) > 0:
            start_afterburner(config.get("MSI Afterburner", "location"), profile)

        affinity_args: list[str] = []
        if config.getboolean("settings", "sync_driver_affinity"):
            affinity_args.extend(["/affinity", hex(1 << cpu)])

        subprocess.run(
            ["start", "", *affinity_args, subject_path],
            shell=True,
            check=False,
        )

        # 5s offset to allow subject to launch
        time.sleep(5 + config.getint("settings", "cache_duration"))

        if config.getboolean("xperf", "enabled"):
            subprocess.run(
                [config.get("xperf", "location"), "-on", "base+interrupt+dpc"],
                check=False,
            )

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
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )

        if not os.path.exists(f"{session_directory}\\CSVs\\CPU-{cpu}.csv"):
            print(
                "error: csv log unsuccessful, this may be due to a missing dependency or windows component",
            )
            shutil.rmtree(session_directory)
            apply_affinity(gpu_hwids, apply=False)
            return 1

        if config.getboolean("xperf", "enabled"):
            subprocess.run(
                [
                    config.get("xperf", "location"),
                    "-d",
                    f"{session_directory}\\xperf\\CPU-{cpu}.etl",
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
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
                ],
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
        num_processes = kernel32.GetConsoleProcesslist(process_array, 1)
        # only pause if script was ran by double-clicking
        if num_processes < 3:
            input("info: press enter to exit")

        sys.exit(__exit_code__)
