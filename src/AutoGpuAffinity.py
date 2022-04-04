import psutil
import wmi
import winreg
import os
import time
import subprocess
import pandas
import csv
import math
from termcolor import colored
from tabulate import tabulate
import platform

gpu_info = wmi.WMI().Win32_VideoController()
subprocess_null = {'stdout': subprocess.DEVNULL, 'stderr': subprocess.DEVNULL}

def kill_processes(*targets):
    for p in psutil.process_iter():
        if p.name() in targets:
            p.kill()

def calc(frametime_data, metric, value=None):
    if metric == 'Max':
        return 1000 / frametime_data['min']
    elif metric == 'Avg':
        return 1000 / (frametime_data['sum'] / frametime_data['len'])
    elif metric == 'Min':
        return 1000 / frametime_data['max']
    elif metric == 'Percentile':
        return 1000 / frametime_data['frametimes'][math.ceil(value / 100 * frametime_data['len']) - 1]
    elif metric == 'Lows':
        current_total = 0
        for present in frametime_data['frametimes']:
            current_total += present
            if current_total >= value / 100 * frametime_data['sum']:
                return 1000 / present

def apply_affinity(action, thread=None):
    def write_key(path, value_name, data_type, value_data):
        with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
            winreg.SetValueEx(key, value_name, 0, data_type, value_data)

    def delete_key(path, value_name):
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY) as key:
                try:
                    winreg.DeleteValue(key, value_name)
                except FileNotFoundError:
                    pass
        except FileNotFoundError:
            pass

    if thread is not None:
        dec_affinity = 0
        dec_affinity |= 1 << thread
        bin_affinity = bin(dec_affinity).replace('0b', '')
        le_hex = int(bin_affinity, 2).to_bytes(8, 'little').rstrip(b'\x00')

    for item in gpu_info:
        gpu_id = item.PnPDeviceID
        if action == 'write':
            write_key(f'SYSTEM\\ControlSet001\\Enum\\{gpu_id}\\Device Parameters\\Interrupt Management\\Affinity Policy', 'DevicePolicy', 4, 4)
            write_key(f'SYSTEM\\ControlSet001\\Enum\\{gpu_id}\\Device Parameters\\Interrupt Management\\Affinity Policy', 'AssignmentSetOverride', 3, le_hex)
        elif action == 'delete':
            delete_key(f'SYSTEM\\ControlSet001\\Enum\\{gpu_id}\\Device Parameters\\Interrupt Management\\Affinity Policy', 'DevicePolicy')
            delete_key(f'SYSTEM\\ControlSet001\\Enum\\{gpu_id}\\Device Parameters\\Interrupt Management\\Affinity Policy', 'AssignmentSetOverride')
        subprocess.run(['pnputil', '/disable-device', gpu_id], **subprocess_null)
        subprocess.run(['pnputil', '/enable-device', gpu_id], **subprocess_null)

def create_lava_cfg():
    lavatriangle_folder = f'{os.environ["USERPROFILE"]}\\AppData\\Roaming\\liblava\\lava triangle'
    try:
        os.makedirs(lavatriangle_folder)
    except FileExistsError:
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

def log(msg):
    print(f'[{time.strftime("%H:%M")}] CLI: {msg}')

def main():
    version = 4.0

    config = {}
    with open('config.txt', 'r') as f:
        for line in f:
            if '//' not in line:
                line = line.strip('\n')
                setting, equ, value = line.rpartition('=')
                if setting != '' and value != '':
                    config[setting] = value
          
    trials = int(config['trials'])
    duration = int(config['duration'])
    dpcisr = int(config['dpcisr'])
    xperf_path = str(config['xperf_path'])
    cache_trials = int(config['cache_trials'])

    if trials <= 0 or cache_trials < 0 or duration <= 0:
        raise ValueError('invalid trials, cache_trials or duration in config')

    threads = psutil.cpu_count()
    cores = psutil.cpu_count(logical=False)

    if threads > cores:
        HT = True
    else:
        HT = False

    if dpcisr != 0 and os.path.exists(xperf_path):
        xperf = True
    else:
        xperf = False

    estimated = (5 + (cache_trials * (duration+ 5)) + (trials * (duration + 5))) * cores
    output_path = f'captures\\AutoGpuAffinity-{time.strftime("%d%m%y%H%M%S")}'
    print_info = f'''
    AutoGpuAffinity v{version} Command Line

        Trials: {trials}
        Trial Duration: {duration} sec
        Cores: {cores}
        Threads: {threads}
        Hyperthreading/SMT: {HT}
        Log dpc/isr with xperf: {xperf}
        Cache trials: {cache_trials}
        Time for completion: {estimated/60:.2f} min
        Session Working directory: \\{output_path}\\
    '''
    print(print_info)
    input('    Press enter to start benchmarking...\n')

    create_lava_cfg()

    os.mkdir(output_path)
    os.mkdir(f'{output_path}\\CSVs')
    if xperf: os.mkdir(f'{output_path}\\xperf')

    main_table = []
    main_table.append(['', 'Max', 'Avg', 'Min', '1 %ile', '0.1 %ile', '0.01 %ile', '0.005 %ile' , '1% Low', '0.1% Low', '0.01% Low', '0.005% Low'])

    # kill all processes before loop
    if xperf: subprocess.run([xperf_path, '-stop'], **subprocess_null)
    kill_processes('xperf.exe', 'lava-triangle.exe', 'PresentMon.exe')

    active_thread = 0
    while active_thread < threads:
        apply_affinity('write', active_thread)
        time.sleep(5)
        subprocess.Popen(['bin\\lava-triangle.exe'], **subprocess_null)

        if cache_trials > 0:
            for trial in range(1, cache_trials + 1):
                log(f'CPU {active_thread} - Cache Trial: {trial}/{cache_trials}')
                time.sleep(duration + 5)

        for trial in range(1, trials + 1):
            file_name = f'CPU-{active_thread}-Trial-{trial}'
            log(f'CPU {active_thread} - Recording Trial: {trial}/{trials}')

            if xperf: subprocess.run([xperf_path, '-on', 'base+interrupt+dpc'])

            try:
                subprocess.run(['bin\\PresentMon.exe', '-stop_existing_session', '-no_top', '-verbose', '-timed', f'{duration}', '-process_name', 'lava-triangle.exe', '-output_file', f'{output_path}\\CSVs\\{file_name}.csv'], timeout=duration + 5, **subprocess_null)
            except subprocess.TimeoutExpired:
                pass

            if not os.path.exists(f'{output_path}\\CSVs\\{file_name}.csv'):
                raise FileNotFoundError('CSV log unsuccessful, this is due to a missing dependency or windows component.')
    
            if xperf:
                subprocess.run([xperf_path, '-stop'], **subprocess_null)
                subprocess.run([xperf_path, '-i', 'C:\\kernel.etl', '-o', f'{output_path}\\xperf\\{file_name}.txt', '-a', 'dpcisr'])

        kill_processes('xperf.exe', 'lava-triangle.exe', 'PresentMon.exe')

        CSVs = []
        for trial in range(1, trials + 1):
            CSV = f'{output_path}\\CSVs\\{file_name}.csv'
            CSVs.append(pandas.read_csv(CSV))
            aggregated = pandas.concat(CSVs)
            aggregated.to_csv(f'{output_path}\\CSVs\\CPU-{active_thread}-Aggregated.csv', index=False)
        
        frametimes = []
        with open(f'{output_path}\\CSVs\\CPU-{active_thread}-Aggregated.csv', 'r') as f:
            for row in csv.DictReader(f):
                if row['MsBetweenPresents'] is not None:
                    frametimes.append(float(row['MsBetweenPresents']))
        frametimes = sorted(frametimes, reverse=True)

        frametime_data = {}
        frametime_data['frametimes'] = frametimes
        frametime_data['min'] = min(frametimes)
        frametime_data['max'] = max(frametimes)
        frametime_data['sum'] = sum(frametimes)
        frametime_data['len'] = len(frametimes)

        data = []
        data.append(f'CPU {active_thread}')
        for metric in ('Max', 'Avg', 'Min'):
            data.append(f'{calc(frametime_data, metric):.2f}')

        for metric in ('Percentile', 'Lows'):
            for value in (1, 0.1, 0.01, 0.005):
                data.append(f'{calc(frametime_data, metric, value):.2f}')
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
    apply_affinity('delete')

    try:
        if int(platform.release()) >= 10:
            highest_fps_color = True
        else:
            highest_fps_color = False
    except:
        highest_fps_color = False

    for column in range(1, len(main_table[0])):
        highest_fps = 0
        row_index = ''
        for row in range(1, len(main_table)):
            fps = float(main_table[row][column])
            if fps > highest_fps:
                highest_fps = fps
                row_index = row
        if highest_fps_color:
            new_value = colored(f'*{float(main_table[row_index][column]):.2f}', 'green')
        else:
            new_value  = f'*{float(main_table[row_index][column]):.2f}'
        main_table[row_index][column] = new_value

    print_result_info = f'''
        > Drag and drop the aggregated data (located in the working directory) into https://boringboredom.github.io/Frame-Time-Analysis for a graphical representation of the data.
        > Affinities for all GPUs have been reset to the Windows default (none).
        > Consider running this tool a few more times to see if the same core is consistently performant.
        > If you see absurdly low values for 0.005% Lows, you should discard the results and re-run the tool.

    '''

    print(print_info)
    print(tabulate(main_table, headers='firstrow', tablefmt='fancy_grid', floatfmt='.2f'), '\n')
    print(print_result_info)

    os.system('cmd.exe /k')

if __name__ == '__main__':
    main()