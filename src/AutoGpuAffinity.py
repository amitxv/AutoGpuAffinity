import psutil
import wmi
import winreg
import os
import time
import argparse
import win32com
import subprocess
import pandas
import csv
import math
from termcolor import colored
from tabulate import tabulate
import ctypes
import sys
import platform

version = '2.0.0'

if ctypes.windll.shell32.IsUserAnAdmin() == False:
    print('Administrator privileges required.')
    quit()

if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    os.chdir(sys._MEIPASS)

parser = argparse.ArgumentParser()
parser.add_argument('-v', '--version', action='store_true', help='show version and exit')
parser.add_argument('-t', '--trials', type=int, metavar='', help='specify the number of trials to benchmark per CPU (default 3)', default=3)
parser.add_argument('-d', '--duration', metavar='', type=int, help='specify the duration of each trial in seconds (default 30)', default=30)
parser.add_argument('-x', '--xperf_log', metavar='', type=bool, help='enable or disable DPC/ISR logging with xperf (Windows ADK required if True) (default True)', default=True)
parser.add_argument('-c', '--app_caching', metavar='', type=int, help='specify the timeout in seconds for application caching after liblava is launched, reliability of results may be affected negatively if too low (default 20)', default=20)
args = parser.parse_args()

if bool(args.version):
    print(f'Current version: {version}')
    quit()

def writeKey(path, valueName, dataType, valueData):
    with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
        winreg.SetValueEx(key, valueName, 0, dataType, valueData)

def deleteKey(path, value_name):
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY) as key:
            try:
                winreg.DeleteValue(key, value_name)
            except:
                pass
    except:
        pass

def getAffinity(thread, return_value):
    dec_affinity = 0
    dec_affinity |= 1 << thread
    bin_affinity = bin(dec_affinity).replace('0b', '')
    le_hex = int(bin_affinity, 2).to_bytes(8, 'little').rstrip(b'\x00')
    if return_value == 'dec':
        return dec_affinity
    elif return_value == 'hex':
        return le_hex

def killProcess(name):
    for p in psutil.process_iter():
        if p.name() == name:
            p.kill()

def calc(frametime_data, metric, value=None):
    if metric == 'Max':
        return 1000 / min(frametime_data)
    elif metric == 'Avg':
        Avg = sum(frametime_data) / len(frametime_data)
        return 1000 / Avg
    elif metric == 'Min':
        return 1000 / max(frametime_data)
    elif metric == 'Percentile':
        return 1000 / frametime_data[math.ceil(value / 100 * len(frametime_data)) - 1]
    elif metric == 'Lows':
        currentTotal = 0.0
        for present in frametime_data:
            currentTotal += present
            if currentTotal >= value / 100 * sum(frametime_data):
                return 1000 / present

threads = psutil.cpu_count()
cores = psutil.cpu_count(logical=False)
gpu_info = wmi.WMI().Win32_VideoController()
wsh = win32com.client.Dispatch('WScript.Shell')
subprocess_null = {'stdout': subprocess.DEVNULL, 'stderr': subprocess.DEVNULL}

if threads > cores:
    HT = True
else:
    HT = False

xperf_paths = [
    'C:\\Program Files\\Microsoft Windows Performance Toolkit\\xperf.exe',
    'C:\\Program Files (x86)\\Windows Kits\\10\\Windows Performance Toolkit\\xperf.exe'
]

xperf_location = None
for i in xperf_paths:
    if os.path.exists(i):
        xperf_location = i

if args.xperf_log and bool(xperf_location):
    xperf = True
else:
    xperf = False

estimated = ((15 + args.app_caching + args.duration) * args.trials) * cores

print_info = f'''
    AutoGpuAffinity Command Line

        Trials: {args.trials}
        Trial Duration: {args.duration} sec
        Cores: {cores}
        Threads: {threads}
        Hyperthreading: {HT}
        Log DPCs/ISRs (xperf): {args.xperf_log}
        Xperf path: {xperf_location}
        Time for completion: {estimated/60:.2f} min

    > Nobody is responsible if you damage your PC or operating system. Run at your own risk.
    > Do not touch your mouse/keyboard while this tool runs to avoid collecting invalid data.
    > Close any background apps you have open.

    Press any key to start benchmarking...
'''
input(print_info)

lavatriangle_folder = f'{os.environ["USERPROFILE"]}\\AppData\\Roaming\\liblava\\lava triangle'
try:
    os.makedirs(lavatriangle_folder)
except:
    pass
lavatriangle_config = f'{lavatriangle_folder}\\window.json'
if os.path.exists(lavatriangle_config):
    os.remove(lavatriangle_config)

lavatriangle_content = [
    '{',
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
    '    }',
    '}'
]

with open(lavatriangle_config, 'a') as f:
    for i in lavatriangle_content:
        f.write(f'{i}\n')

working_dir = f'{os.environ["TEMP"]}\\AutoGpuAffinity{time.strftime("%d%m%y%H%M%S")}'
os.mkdir(working_dir)
os.mkdir(f'{working_dir}\\raw')
if xperf:
    os.mkdir(f'{working_dir}\\xperf')
os.mkdir(f'{working_dir}\\aggregated')

main_table = []
main_table.append(['', 'Max', 'Avg', 'Min', '1 %ile', '0.1 %ile', '0.01 %ile', '0.005 %ile' , '1% Low', '0.1% Low', '0.01% Low', '0.005% Low'])

# kill all processes before loop
if xperf:
    subprocess.run([xperf_location, '-stop'], **subprocess_null)
    killProcess('xperf.exe')
killProcess('lava-triangle.exe')
killProcess('PresentMon.exe')
killProcess('PresentMon-1.6.0-x64.exe')

if platform.release() == 10:
    present_mon_exe = "PresentMon.exe"
else:
    present_mon_exe = "PresentMon-1.6.0-x64.exe"

PRESENTMON_DEBUG_ARG = {"PresentMon.exe": "-track_debug", "PresentMon-1.6.0-x64.exe": "-verbose"}

active_thread = 0
while active_thread != threads:
    for item in gpu_info:
        writeKey(f'SYSTEM\\ControlSet001\\Enum\\{item.PnPDeviceID}\\Device Parameters\\Interrupt Management\\Affinity Policy', 'DevicePolicy', 4, 4)
        writeKey(f'SYSTEM\\ControlSet001\\Enum\\{item.PnPDeviceID}\\Device Parameters\\Interrupt Management\\Affinity Policy', 'AssignmentSetOverride', 3, getAffinity(active_thread, 'hex'))
    subprocess.run(['restart64.exe', '/q'])
    time.sleep(5)
    subprocess.run(['cmd.exe', '/c', 'start', '/affinity', f'{getAffinity(active_thread, "dec")}', 'lava-triangle.exe'])
    time.sleep(args.app_caching + 5)

    for active_trial in range(1, args.trials + 1):
        print(f'Currently benchmarking: CPU-{active_thread}-Trial-{active_trial}/{args.trials}...')
        wsh.AppActivate('lava triangle')
        if xperf:
            subprocess.run([xperf_location, '-on', 'base+interrupt+dpc'])
        try:
            subprocess.run([present_mon_exe, '-stop_existing_session', '-no_top', PRESENTMON_DEBUG_ARG[present_mon_exe], '-timed', f'{args.duration}', '-process_name', 'lava-triangle.exe', '-output_file', f'{working_dir}\\raw\\CPU-{active_thread}-Trial-{active_trial}.csv'], timeout=args.duration + 5, **subprocess_null)
        except subprocess.TimeoutExpired:
            pass
        if xperf:
            subprocess.run([xperf_location, '-stop'], **subprocess_null)
            subprocess.run([xperf_location, '-i', 'C:\\kernel.etl', '-o', f'{working_dir}\\xperf\\CPU-{active_thread}-Trial-{active_trial}.txt', '-a', 'dpcisr'])
            killProcess('xperf.exe')
    killProcess('PresentMon.exe')
    killProcess('lava-triangle.exe')

    CSVs = []
    for trial in range(1, args.trials + 1):
        CSV = f'{working_dir}\\raw\\CPU-{active_thread}-Trial-{trial}.csv'
        CSVs.append(pandas.read_csv(CSV))
        aggregated = pandas.concat(CSVs)
        aggregated.to_csv(f'{working_dir}\\aggregated\\CPU-{active_thread}-aggregated.csv', index=False)
    
    frametimes = []
    with open(f'{working_dir}\\aggregated\\CPU-{active_thread}-aggregated.csv', 'r') as f:
        for row in csv.DictReader(f):
            if row['msBetweenPresents'] is not None:
                frametimes.append(float(row['msBetweenPresents']))
    frametimes = sorted(frametimes, reverse=True)

    data = []
    data.append(f'CPU {active_thread}')
    for metric in ('Max', 'Avg', 'Min'):
        data.append(float(f'{calc(frametimes, metric):.2f}'))

    for metric in ('Percentile', 'Lows'):
        for value in (1, 0.1, 0.01, 0.005):
            data.append(float(f'{calc(frametimes, metric, value):.2f}'))
    main_table.append(data)

    if HT:
        active_thread += 2
    else:
        active_thread += 1

if os.path.exists('C:\\kernel.etl'):
    os.remove('C:\\kernel.etl')

os.system('color')
os.system('cls')
os.system('mode 300, 1000')

for item in gpu_info:
    deleteKey(f'SYSTEM\\ControlSet001\\Enum\\{item.PnPDeviceID}\\Device Parameters\\Interrupt Management\\Affinity Policy', 'DevicePolicy')
    deleteKey(f'SYSTEM\\ControlSet001\\Enum\\{item.PnPDeviceID}\\Device Parameters\\Interrupt Management\\Affinity Policy', 'AssignmentSetOverride')

for column in range(1, len(main_table[0])):
    highest_fps = 0
    row_index = ''
    for row in range(1, len(main_table)):
        fps = main_table[row][column]
        if fps > highest_fps:
            highest_fps = fps
            row_index = row
    main_table[row_index][column] = colored(main_table[row_index][column], 'green')

result = f'''

    > Raw and aggregated data is located in the following directory:

            {working_dir}
        
    > Drag and drop the aggregated data into "https://boringboredom.github.io/Frame-Time-Analysis" for for a graphical representation of the data.

    > Affinities for all GPUs have been reset to the Windows default (none).

    > Green values in the table below indicate it is the highest value for a given metric/value.

    > Consider running this tool a few more times to see if the same core is consistently performant.

    > If you see absurdly low values for 0.005% Lows, you should discard the results and re-run the tool.

'''

print(result)
print(tabulate(main_table, headers='firstrow', tablefmt='fancy_grid', floatfmt='.2f'), '\n')
input('Press any key to exit...')
