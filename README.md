# Handbrake Batcher and Merger

This script allows the user to automate the Handbrake conversion and post conversion MkvMerge commands
Just download the Handbrake Batcher and Merger.ps1 and launch it

## First Launch

With the first launch the script generate the basic configuration files
.\configuration\options.json - Here are stored the basic and global configuration
.\configuration\Working Folder List.csv - Here are stored the folder to work on (recursive search) and the per-folder configurations

If specified a configuration at folder level it wins. If not specified the global one is used

After the first launch the script terminate creating only the files specified above. Review it and launch again the script to see it in action.
Please be aware of what you modify. The script is sensible with the data manually inserted. But just try it.

## Allowed instances

Only 1 instance at time of the script is allowed (powershell.exe or pwsh.exe with the script name).

If an instance of HandBrake.Worker.exe or HandBrakeCLI.exe is running, this is counted as an instance.

The script automatically wait for the instance to end to run itself

## Options

### runOptions

waitSecondsOption: this is a numeric indicator to set how to wait between instance check. 

If launching the script to times only 1 instance is allowed. When the first instance 

