# Handbrake Batcher and Merger

This script allows the user to automate the Handbrake conversion and post conversion MkvMerge commands
Just download the Handbrake Batcher and Merger.ps1 and launch it

Future features and todo list:

- [ ] Allow the user to customize if running a single or multiple instance of the script
- [ ] Allow the user to customize the instance .exe to check. Now defaulting in the code to "Handbrake.Worker.exe", "HandBrakeCLI.exe", "powershell.exe" etc
- [ ] Change the handbrake conversion part to a generic command part to allow the user to use other command line software (ex. FFMPEG)
- [ ] Add a variable for the destination converted name append suffix defaulting now to "_handbrake"
- [ ] Change some mkvmerge variables to more generic ones to let the user free to do more customization over the commands
- [x] Add a new variable to customize how many day keep the logs

## First Launch

With the first launch the script generate the basic configuration files
.\configuration\options.json - Here are stored the basic and global configuration
.\configuration\Working Folder List.csv - Here are stored the folder to work on (recursive search) and the per-folder configurations


After the first launch the script terminate creating only the files specified above. Review it and launch again the script to see it in action.
Please be aware of what you modify. The script is sensible with the data manually inserted. But just try it.

Note that HandBrake, HandbrakeCLI and MkvMerge need to be available on the system

## Allowed instances

Only 1 instance at time of the script is allowed (powershell.exe or pwsh.exe with the script name).

If an instance of HandBrake.Worker.exe or HandBrakeCLI.exe is running, this is counted as an instance.

The script automatically wait for the instance to end to run itself

## Log

Logs are stored in .\log folder

## Automatically discarded files

If the script have to work on a FileName with already an "FileName_handbrake.mkv" generated (maybe from another run), this file is automatically threated as already elaborated.
When running the Handbrake command the actual "FileName_handbrake.mkv" is added to a csv and removed after the correct conversion. If the script is interrupted during this conversion, the next run will retry the file as already in csv exclusion list.

## Options

If the same option is in Working Folder List.csv is threated as a per-folder options and is applied only to that folder. 

If not specified the global options.json option is applied

Please be aware in the options.json that some characters need to be specified escaped (with a preceding '\')

### runOptions

waitSecondsOption: this is a numeric indicator to set how many seconds to wait between instances check.

keepLogForDays: this is a numeric indicator to set how many days the log are kept in the log folder.

debugLog: activate debug log (true/false).

runHandbrake: tells the script to run the HandbrakeCLI command for every accepted file

runMkvMerge: tells the script to run the MkvMerge command for every accepted file

### fileFolderOptions

excludeFileList: specify the mask to exclude files from recursive elaboration. Faster than excludeFileFolderList option.

excludeFileFolderList: specify the mask to exclude files and folders from recursive elaboration. 
The option is primary for folders that cannot specified in excludeFileList

includeList: specify the mask to include files from recursive elaboration

### conversionOptions

conversionCustomField1: Custom field 1. Default: the location of preset json

conversionCustomField2: Custom field 1. Default: the preset to work on

Add to the ps1 main file as many conversionCustomField you want duplicating conversionCustomField2

conversionCommand: command to launch for execute conversion.

#### Wildcards: 

||conversionCustomField1|| -> Custom field 1

||conversionCustomField2|| -> Custom field 2

||inputFile|| -> the input file that's the script is running on

||outputFile|| -> the output file converted by HandbrakeCLI. this is generated with the input file name, appending "_handrake.mkv".

### mkvMergeOptions

mkvMergeLocation: the MkvMerge.exe file location

mkvMergeCommand: The command to launch for execute the MkvMerge.exe conversion.

You can work with my other project https://github.com/Bazzu85/MKVmergeBatcher/releases to extract command to use.

#### Wildcards: 

||mkvMergeLocation|| -> the mkvmerge location specified in mkvMergeLocation options

||outputFileName|| -> the output file name of the conversion. is composed or FileFolder\tmp\FileName.mkv

||handbrakeFileName|| -> the converted handbrake file

||inputFileName|| -> the input file from the execution. The same used by handbrakeCLI

