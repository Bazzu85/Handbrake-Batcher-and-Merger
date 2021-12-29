# Handbrake Batcher and Merger

This script allows the user to automate the Handbrake conversion and post conversion MkvMerge commands
Just download the Handbrake Batcher and Merger.ps1 and launch it

## First Launch

With the first launch the script generate the basic configuration files
.\configuration\options.json - Here are stored the basic and global configuration
.\configuration\Working Folder List.csv - Here are stored the folder to work on (recursive search) and the per-folder configurations


After the first launch the script terminate creating only the files specified above. Review it and launch again the script to see it in action.
Please be aware of what you modify. The script is sensible with the data manually inserted. But just try it.

## Allowed instances

Only 1 instance at time of the script is allowed (powershell.exe or pwsh.exe with the script name).

If an instance of HandBrake.Worker.exe or HandBrakeCLI.exe is running, this is counted as an instance.

The script automatically wait for the instance to end to run itself

## Log

Logs are stored in .\log folder

## Options

If the same option is in Working Folder List.csv is threated as a per-folder options and is applied only to that folder. 

If not specified the global options.json option is applied

### runOptions

waitSecondsOption: this is a numeric indicator to set how many seconds to wait between instances check.

debugLog: activate debug log (true/false).

runHandbrake: tells the script to run the HandbrakeCLI command for every accepted file

runMkvMerge: tells the script to run the MkvMerge command for every accepted file

