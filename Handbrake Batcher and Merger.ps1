Function Test-IfAlreadyRunning {
    <#
    .SYNOPSIS
        If CURRENT instance is running wait for N seconds.
    .PARAMETER ScriptName
        Name of this script
    .EXAMPLE
        $ScriptName = $MyInvocation.MyCommand.Name
        Test-IfAlreadyRunning -ScriptName $ScriptName
    .NOTES
        $PID is a Built-in Variable for the current script''s Process ID number
    .LINK
    #>
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullorEmpty()]
            [String]$ScriptName
        )

        [Bool]$loop = $true
        $i = 0
        while ($loop){
            #Get array of all powershell scripts currently running
            $PsScriptsRunning = Get-CimInstance -ClassName Win32_Process | Where-Object{($_.name -eq 'powershell.exe') -or ($_.name -eq 'pwsh.exe') -or ($_.name -eq 'HandBrake.Worker.exe') -or ($_.name -eq 'HandBrakeCLI.exe')} | select-object commandline,ProcessId

            $i++
            [Bool]$found = $false
            LogWrite $DEBUG $("Searching process. Try number $i")
            #enumerate each element of array and compare
            ForEach ($PsCmdLine in $PsScriptsRunning){
                [Int32]$OtherPID = $PsCmdLine.ProcessId
                [String]$OtherCmdLine = $PsCmdLine.commandline
                #Are other instances of this script already running?
                If (($OtherCmdLine.Contains($ScriptName)) -And ($OtherPID -ne $PID) -and (!$found)){
                    LogWrite $INFO $("PID [$OtherPID] is already running (this script). Waiting $waitSeconds seconds")
                    $found = $true
                }
                If (($OtherCmdLine.Contains("HandBrake.Worker.exe")) -And ($OtherPID -ne $PID) -and (!$found) -and ($runHandbrake)){
                    LogWrite $INFO $("PID [$OtherPID] is already running (HandBrake.Worker.exe). Waiting $waitSeconds seconds")
                    $found = $true
                }
                If (($OtherCmdLine.Contains("HandBrakeCLI.exe")) -And ($OtherPID -ne $PID) -and (!$found) -and ($runHandbrake)){
                    LogWrite $INFO $("PID [$OtherPID] is already running (HandBrakeCLI.exe). Waiting $waitSeconds seconds")
                    $found = $true
                }
            }

            if (!$found){
                $loop = $false
            } else {
                Start-Sleep -Second $waitSeconds
            }
        }
    } #Function Test-IfAlreadyRunning

Function LogWrite ($logType , $logString){
    if (!$logType){
        $logType ="INFO "
    }
    $logOk = $true
    if (($logType -eq "DEBUG") -and !$debugLog){
        $logOk = $false
    }
    if ($logOk){
        $logstring = "$(Get-Date -Format "yyyyMMdd_HHmm") $logType $logstring"
        Write-Host $logstring
        Add-content $logfile -value $logstring
    }
}
Function Get-OptionsFromJson {
    #Default options object
    $options = @{
        runOptions= @{
            debugLog1 = $true
            runHandbrake=$false
            runMkvMerge=$false
            waitSecondsOption=60
        }
        handbrakeOptions= @{
            handbrakePresetToUse=0 #Number of preset to use
            handbrakePresetList=@("Convert to h265 Medium 720p (only video)",
                                "Convert to h265 Medium (only video)",
                                "Convert to h265 Medium 480p + audio AAC"
                                )
            handbrakePresetLocation="C:\Users\elbaz\AppData\Roaming\HandBrake\presets.json"
            handbrakeCommand='HandBrakeCLI.exe --preset-import-file "||handbrakePresetLocation||" -Z "||handbrakePreset||" -i "||inputFile||" -o "||outputFile||"'
        }
        mkvMergeOptions= @{
            mkvMergeCommandToLaunch=0 #Number of command to use
            mkvMergeCommandList=@('||mkvMergeLocation|| --ui-language en --output ^"||outputFileName||^" --language 0:und --compression 0:none --no-track-tags  --no-global-tags ^"^(^" ^"||handbrakeFileName||^" ^"^)^" --no-video ^"^(^" ^"||inputFileName||^" ^"^)^"',
                                '||mkvMergeLocation|| --ui-language en --output ^"||outputFileName||^" --language 0:und --compression 0:none --no-track-tags  --no-global-tags ^"^(^" ^"||handbrakeFileName||^" ^"^)^" --no-video ^"^(^" ^"||inputFileName||^" ^"^)^"'
                                )
            mkvMergeLocation='"C:\Program Files\MKVToolNix\mkvmerge.exe"'
        }
        fileFolderOptions= @{
            includeList=@("*.mp4",
                        "*.mkv"
                            )
            excludeFileList=@("*_handbrake.mkv*")
            excludeFileFolderList=@("\1ok","\1tmp","\1wrn")
        }
    }
    # If not found, create the .config file
    if (!(Test-Path -Path $configFile)){
        ConvertTo-Json -InputObject $options | Out-File $configFile
    }
    $options = Get-Content -LiteralPath $configFile | ConvertFrom-Json
    return $options
}

Function Get-WorkingFolderListFile {

    # Check if folderList.txt is present. 
    # If not create it
    # If yes read all folders in array
    if (!(Test-Path -Path $workingFolderListFile)){
        $workingFolderListFile = @($PSScriptRoot) | Out-File -FilePath $workingFolderListFile
    } else {
        $workingFolderListFile = Get-Content -Path $workingFolderListFile
    }
    return $workingFolderListFile
}
Function Get-WorkingHandbrakeList {
    # Check if folderList.txt is present. 
    # If not create it
    # If yes read all folders in array
    if (!(Test-Path -Path $workingHandbrakeFile)){
        #$workingHandbrakeList = "" | Out-File -FilePath $workingHandbrakeFile
    } else {
        $workingHandbrakeList = @(Get-Content -Path $workingHandbrakeFile)
    }
    return $workingHandbrakeList
}

Function Set-WorkingHandbrakeList ($workingHandbrakeList) {
    # Write the workingHandbrakeList into file
    Set-Content $workingHandbrakeFile -value $workingHandbrakeList
}

#Main Code

#Set title
$host.UI.RawUI.WindowTitle = $MyInvocation.MyCommand.Name.Replace(".ps1","")

Write-Host "Handbrake Batcher and Merger by Bazzu v.1.0"
$host
#Set log types
$DEBUG = "DEBUG"
$INFO  = "INFO "

$configFile = "$PSScriptRoot\$($MyInvocation.MyCommand.Name.Replace('.ps1','')).config"
$options = Get-OptionsFromJson

$workingFolderListFile = "$PSScriptRoot\Working Folder List.txt"
$workingHandbrakeFile = "$PSScriptRoot\Working Handbrake File List.txt"
$date = Get-Date -Format "yyyyMMdd"
$logFile = "$PSScriptRoot\log\$($MyInvocation.MyCommand.Name.Replace('.ps1',''))_$date.log"
if (!(Test-Path -Path $logFile)){
    New-Item -Path ("$PSScriptRoot\log") -ItemType Directory -Force
}

#Set the working variables

#runOptions section
[bool]$debugLog = $options.runOptions.debugLog
[bool]$runHandbrake = $options.runOptions.runHandbrake
[bool]$runMkvMerge = $options.runOptions.runMkvMerge
$waitSeconds = $options.runOptions.waitSecondsOption

#handbrakeOptions section
$handbrakePresetToUse = $options.handbrakeOptions.handbrakePresetToUse
$handbrakePresetlist = $options.handbrakeOptions.handbrakePresetList
$handbrakePresetLocation = $options.handbrakeOptions.handbrakePresetLocation
$handbrakeCommand = $options.handbrakeOptions.handbrakeCommand
foreach ($handbrakePreset in $handbrakePresetlist){
    LogWrite $DEBUG $("HandBrake. Preset from json: $handbrakePreset")
}
LogWrite $DEBUG $("HandBrake. Preset to use: $($handbrakePresetlist[$handbrakePresetToUse])")
LogWrite $DEBUG $("HandBrake. Preset location: $handbrakePresetLocation")

#mkvMergeOptions section
$mkvMergeLocation = $options.mkvMergeOptions.mkvMergeLocation
$mkvMergeCommandToLaunch = $options.mkvMergeOptions.mkvMergeCommandToLaunch
$mkvMergeCommandList = $options.mkvMergeOptions.mkvMergeCommandList
foreach ($mkvMergeCommand in $mkvMergeCommandList){
    LogWrite $DEBUG $("MkvMerge. Preset from json: $mkvMergeCommand")
}
LogWrite $DEBUG $("Preset to use: $($mkvMergeCommandList[$mkvMergeCommandToLaunch])")

#fileFolderOptions section
$includeList = $options.FileFolderOptions.includeList
$total = $includeList.count
$counter = 0
$includeString = ''
foreach ($include in $includeList){
    $counter += 1
    if ($counter -eq $total){
        $includeString += "$include"
    } else {
        $includeString += "$include , " #empty
    }
}
$excludeFileList = $options.FileFolderOptions.excludeFileList
$total = $excludeFileList.count
$counter = 0
$excludeFileListString = ''
foreach ($excludeFile in $excludeFileList){
    $counter += 1
    if ($counter -eq $total){
        $excludeFileListString += "$excludeFile"
    } else {
        $excludeFileListString += "$excludeFile , " #empty
    }
}
$excludeFileFolderList = $options.FileFolderOptions.excludeFileFolderList

# Get name of current script and check if already running
$ScriptName = $MyInvocation.MyCommand.Name 
Test-IfAlreadyRunning -ScriptName $ScriptName
LogWrite $INFO $("(PID=[$PID]) This is the 1st and only instance allowed to run") #this only shows in one instance

# Get workingFolderListFile from folderList.txt
$workingFolderListFile = Get-WorkingFolderListFile

# Work on all folder in Working Folder List.txt
foreach ($workingFolder in $workingFolderListFile){
    # if the folder is not there, jump to the next
    if (!(Test-Path -Path $workingFolder)){
        LogWrite $DEBUG $("Folder $workingFolder not found. Skipping")
        continue
    }
    LogWrite $INFO $("Running script in folder $workingFolder")
    #Get the file list
    $files = Get-ChildItem -Path $workingFolder -Include $includeList -Recurse -Exclude $excludeFileList 
    foreach ($file in $files){
        #Check if the file exists
        if (!(Test-Path -LiteralPath $file -PathType leaf)){
            LogWrite $INFO $("File $($file.FullName) not found. Jumping to next")
            continue #jump to next item in foreach
        }

        LogWrite $DEBUG $("Working on $($file.name)")

        
        #Check if handbrake was already launched (only if runHandbrake requested). If found jump to next file
        $handbrakeDestinationFile = "$($file.DirectoryName)\$($file.BaseName)_handbrake.mkv"
        $foundWorkingHandbrake = $false

        ### Check if $handbrakeDestinationFile was pending in previous runs. If not jump to next file
        if ($runHandbrake){
            # Using the @ return always an array. If the file has only 1 row return String so it's necessary
            $workingHandbrakeList = @(Get-WorkingHandbrakeList)
            if ($workingHandbrakeList -Contains $handbrakeDestinationFile){
                LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile found from a previous launch. Retrying")
                $foundWorkingHandbrake = $true
            } else {
                [bool]$foundHandbrakeDestinationFile = $false
                if ((Test-Path -LiteralPath $handbrakeDestinationFile -PathType leaf)){
                    LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile found")
                    $foundHandbrakeDestinationFile = $true
                } else {
                    LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile not found")
                }
                if ($foundHandbrakeDestinationFile){
                    continue
                }
            }
        }
        
        #Check if we are working on a file/folder to exclude. If found jump to next file
        $jumpToNextFile = $false
        foreach ($excludeFileFolder in $excludeFileFolderList){
            if ($file.FullName.Contains($excludeFileFolder)){
                $jumpToNextFile = $true
            }
        }
        if ($jumpToNextFile){
            LogWrite $DEBUG $("Skipping file $($file.FullName) for exclusion list")
            continue #jump to next item in foreach
        }

        ### Exec conversion with HandBrakeCli
        LogWrite $INFO $("Working on $($file.FullName)")
        $handbrakeNewcommand = $handbrakeCommand
        $handbrakeNewcommand = $handbrakeNewcommand.Replace("||handbrakePresetLocation||", $handbrakePresetLocation)
        $handbrakeNewcommand = $handbrakeNewcommand.Replace("||handbrakePreset||", $($handbrakePresetlist[$handbrakePresetToUse]))
        $handbrakeNewcommand = $handbrakeNewcommand.Replace("||inputFile||", $file)
        $handbrakeNewcommand = $handbrakeNewcommand.Replace("||outputFile||", $handbrakeDestinationFile)

        LogWrite $INFO $("Handbrake command: $handbrakeNewcommand")
        if ($runHandbrake){
            if (!$foundWorkingHandbrake){
                # Using the @ return always an array. If the file has only 1 row return String so it's necessary
                $workingHandbrakeList = @(Get-WorkingHandbrakeList)
                $workingHandbrakeList += $handbrakeDestinationFile
                Set-WorkingHandbrakeList $workingHandbrakeList
                LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile added to $workingHandbrakeFile")
            }
            Invoke-Expression $handbrakeNewcommand
            $handbrakeExitCode = $LASTEXITCODE
            LogWrite $INFO $("Handbrake rc: $handbrakeExitCode of $($file.name)")
            if ($handbrakeExitCode -eq 0){
                $workingHandbrakeList = Get-WorkingHandbrakeList
                $workingHandbrakeList = $workingHandbrakeList -ne $handbrakeDestinationFile
                LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile removed from $workingHandbrakeFile")
                Set-WorkingHandbrakeList $workingHandbrakeList
                LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile added to $workingHandbrakeFile")
            }
        } else {
            LogWrite $INFO $("Handbrake disabled")
        }

        ### Merge file with command in options
        $inputfile = "$($file.DirectoryName)\$($file.Name)"
        $outputfile = "$($file.DirectoryName + "\1tmp")\$($file.Name.Replace($file.Extension,".mkv"))"
        $mergeNewcommand = $mkvMergeCommandList[$mkvMergeCommandToLaunch]
        $mergeNewcommand = $mergeNewcommand.Replace("||mkvMergeLocation||", $mkvMergeLocation)
        $mergeNewcommand = $mergeNewcommand.Replace("||outputFileName||", $outputfile)
        $mergeNewcommand = $mergeNewcommand.Replace("||inputFileName||", $inputfile)
        $mergeNewcommand = $mergeNewcommand.Replace("||handbrakeFileName||", $handbrakeDestinationFile)

        $outputfile =  $outputfile.Replace("\\","\")

        LogWrite $INFO $("Merging to $outputfile")
        LogWrite $INFO $("Mkvmerge command: cmd /c $mergeNewcommand")
        if ($runMkvMerge){
            cmd /c $mergeNewcommand
            $mergeExitCode = $LASTEXITCODE
            LogWrite $INFO $("MkvMerge rc: $mergeExitCode of $outputfile")
            $moveFile = $false
            if ($mergeExitCode -eq 0){
                $newOutFile = $outputfile.Replace("\1tmp","\1ok")
                $moveFile = $true
            }
            if ($mergeExitCode -eq 1){
                $newOutFile = $outputfile.Replace("\1tmp","\1wrn")
                $moveFile = $true
            }
            if ($moveFile){
                if ((Test-Path -LiteralPath $newOutFile -PathType leaf)){
                    LogWrite $INFO $("Destination file $newOutFile already exists. Skipping move")
                } else {
                    if(!(Test-Path ($file.DirectoryName + "\1ok")))
                    {
                        New-Item -Path ($file.DirectoryName + "\1ok") -ItemType Directory -Force
                    }
                    Move-Item -LiteralPath $outputfile -Destination ($newOutFile) #-Force
                    # if directory is empty, delete it
                    $directoryInfo = Get-ChildItem ($file.DirectoryName + "\1tmp") | Measure-Object
                    if ($directoryInfo.count -eq 0){
                        Remove-Item -LiteralPath ($file.DirectoryName + "\1tmp")
                    }
                    LogWrite $DEBUG $("Moved $outputfile to $newOutFile")
                }
            }
        } else {
            LogWrite $INFO $("MKVmerge disabled")
        }
    }
}


if ($debugLog){
    LogWrite $DEBUG $("Script ended. Closing")
    #pause
}


