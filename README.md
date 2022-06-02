## AutoGpuAffinity

<img src="./img/exampleoutput.png" width="1000"> 

CLI tool to automatically benchmark the most performant core based on lows/percentile fps in [liblava](https://github.com/liblava/liblava).

Contact: https://twitter.com/amitxv

## Disclaimer
I am not responsible for damage caused to computer. There is a risk of your GPU driver not responding after restarting it during the tests.

## Maintaing a consistent benchmarking environment:

 - Set static overclocks for the GPU/CPU
 - Disable every other p-state except p0
 - Disable c-states/ disable idle
 - Close background applications
 - Do not touch your mouse/keyboard while this tool runs

## Usage

- [Windows ADK](https://docs.microsoft.com/en-us/windows-hardware/get-started/adk-install) is required for DPC/ISR logging with xperf.

- To maintain overclock settings with MSI Afterburner throughout the tests:

    - save the desired settings to a profile (e.g profile 1)
    - configure the path along with the profile to load in config.txt
    
- Download and extract the latest release from the [releases tab](https://github.com/amitxv/AutoGpuAffinity/releases).

- CD to the extracted directory and run ``AutoGpuAffinity.exe`` through the command-line and press enter when ready. If you double-click instead of running in the CLI, then the program will close after benchmarking and you will not be able to see the table of results.

- After the tool has benchmarked each core, a table will be displayed with the results.

## Determine if results are reliable:

Run the tool (not trials) two or three times. If the same core is consistently performant and no 0.005% Lows values are absurdly low compared to other results, then your results are reproducible and your testing environment is consistent. However, if the fps is dropping significantly, consider properly tuning your setup before using this tool.