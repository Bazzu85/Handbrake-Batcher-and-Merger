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
            mkvMergeCommandList=@('"||mkvMergeLocation||" --ui-language en --output ^"||outputFileName||^" --language 0:und --compression 0:none --no-track-tags  --no-global-tags ^"^(^" ^"||handbrakeFileName||^" ^"^)^" --no-video ^"^(^" ^"||inputFileName||^" ^"^)^"')
            mkvMergeLocation="C:\Program Files\MKVToolNix\mkvmerge.exe"
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
    if (!(Test-Path -Path $optionsFile)){
        ConvertTo-Json -InputObject $options | Out-File $optionsFile
    }
    $options = Get-Content -LiteralPath $optionsFile | ConvertFrom-Json
    return $options
}
Function Get-WorkingFolderListFileCsv {

    if (!(Test-Path -LiteralPath "$PSScriptRoot\configuration")){
        New-Item -Path ("$PSScriptRoot\configuration") -ItemType Directory -Force
    }
    # Check if $workingFolderListCsvFile is present. 
    # If not create it
    # If yes read all folders in array
    if (!(Test-Path -LiteralPath $workingFolderListCsvFile -PathType leaf)){
        $workingFolderListCsv = [ordered]@{
            path = $PSScriptRoot
            handbrakePresetLocation="C:\Users\elbaz\AppData\Roaming\HandBrake\presets.json"
            handbrakePreset="Convert to h265 Medium 720p (only video)"
            handbrakeCommand='HandBrakeCLI.exe --preset-import-file "||handbrakePresetLocation||" -Z "||handbrakePreset||" -i "||inputFile||" -o "||outputFile||"--auto-anamorphic'
            mkvMergeLocation="C:\Program Files\MKVToolNix\mkvmerge.exe"
            mkvMergeCommand='"||mkvMergeLocation||" --ui-language en --output ^"||outputFileName||^" --language 0:und --compression 0:none --no-track-tags  --no-global-tags ^"^(^" ^"||handbrakeFileName||^" ^"^)^" --no-video ^"^(^" ^"||inputFileName||^" ^"^)^"'
            runHandbrake=$false
            runMkvMerge=$false
        }
        $workingFolderListCsv | Export-Csv -Path $workingFolderListCsvFile
    } else {
        $workingFolderListCsv = Import-Csv -Path $workingFolderListCsvFile
    }
    return $workingFolderListCsv
}
Function Get-WorkingHandbrakeListCsv {
    # Check if folderList.txt is present. 
    # If not create it
    # If yes read all folders in array
    if (!(Test-Path -Path $workingHandbrakeFileCsv)){
        LogWrite $DEBUG $("File $workingHandbrakeFileCsv not found")
        $workingHandbrakeList = $null
    } else {
        LogWrite $DEBUG $("File $workingHandbrakeFileCsv found. Importing CSV")
        $workingHandbrakeList = Import-Csv -LiteralPath $workingHandbrakeFileCsv
    }
    return $workingHandbrakeList
}

Function Add-WorkingHandbrakeListCsv ($workingHandbrakeListCsv, $handbrakeDestinationFile) {
    # add the handbrakeDestinationFile into file
    # if empty
    if (!$workingHandbrakeListCsv){
        $workingHandbrakeListCsv = [ordered]@{
            file = $handbrakeDestinationFile
        }
        $workingHandbrakeListCsv | Export-Csv -Path $workingHandbrakeFileCsv
    } else {
        $newRow = [ordered]@{
            file = $handbrakeDestinationFile
        }
        $newRow | Export-Csv -Path $workingHandbrakeFileCsv -Append -Force
    }
}

Function Remove-FromWorkingHandbrakeListCsv ($workingHandbrakeListCsv, $handbrakeDestinationFile) {
    # remove the handbrakeDestinationFile from file
    $newWorkingHandbrakeListCsv = $workingHandbrakeListCsv | Where-Object {$_.file -ne $handbrakeDestinationFile}
    if (!$newWorkingHandbrakeListCsv){
        Remove-Item -LiteralPath $workingHandbrakeFileCsv -Force
    } else {
        $newWorkingHandbrakeListCsv | Export-Csv -Path $workingHandbrakeFileCsv
    }
}

#Main Code

#Set title
$host.UI.RawUI.WindowTitle = $MyInvocation.MyCommand.Name.Replace(".ps1","")

Write-Host "Handbrake Batcher and Merger by Bazzu v.2.0"
$host
#Set log types
$DEBUG = "DEBUG"
$INFO  = "INFO "

$optionsFile = "$PSScriptRoot\configuration\options.json"
$options = Get-OptionsFromJson

$workingFolderListCsvFile = "$PSScriptRoot\configuration\Working Folder List.csv"
$workingHandbrakeFileTxt = "$PSScriptRoot\Working Handbrake File List.txt"
$workingHandbrakeFileCsv = "$PSScriptRoot\configuration\Working Handbrake File List.csv"
$date = Get-Date -Format "yyyyMMdd"
$logFile = "$PSScriptRoot\log\$($MyInvocation.MyCommand.Name.Replace('.ps1',''))_$date.log"
if (!(Test-Path -Path $logFile)){
    New-Item -Path ("$PSScriptRoot\log") -ItemType Directory -Force
}

#Set the working variables

#runOptions section
[bool]$debugLog = $options.runOptions.debugLog
$waitSeconds = $options.runOptions.waitSecondsOption

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

# Get workingFolderListTxtFile from folderList.txt
$workingFolderListCsv = Get-WorkingFolderListFileCsv

# Work on all folder in Working Folder List.txt
foreach ($workingFolder in $workingFolderListCsv){
    # if the folder is not there, jump to the next
    if (!(Test-Path -Path $($workingFolder.path))){
        LogWrite $DEBUG $("Folder $($workingFolder.path) not found. Skipping")
        continue
    }
    LogWrite $INFO $("Running script in folder $($workingFolder.path)")
    
    # Set the variable from custom folder csv or from generic configuration
    if ($($workingFolder.handbrakePresetLocation)){
        $handbrakePresetLocation = $workingFolder.handbrakePresetLocation
    } else {
        $handbrakePresetLocation = $options.handbrakeOptions.handbrakePresetLocation
    }
    if ($($workingFolder.handbrakePreset)){
        $handbrakePreset = $workingFolder.handbrakePreset
    } else {
        $handbrakePresetToUse = $options.handbrakeOptions.handbrakePresetToUse
        $handbrakePresetlist = $options.handbrakeOptions.handbrakePresetList
        $handbrakePreset = $handbrakePresetlist[$handbrakePresetToUse]
        #handbrakeOptions section
        foreach ($handbrakePreset in $handbrakePresetlist){
            LogWrite $DEBUG $("HandBrake. Preset from json: $handbrakePreset")
        }
        LogWrite $DEBUG $("HandBrake. Preset to use: $($handbrakePresetlist[$handbrakePresetToUse])")
        LogWrite $DEBUG $("HandBrake. Preset location: $handbrakePresetLocation")
    }
    if ($($workingFolder.handbrakeCommand)){
        $handbrakeCommand = $workingFolder.handbrakeCommand
    } else {
        $handbrakeCommand = $options.handbrakeOptions.handbrakeCommand
    }
    if ($($workingFolder.mkvMergeLocation)){
        $mkvMergeLocation = $workingFolder.mkvMergeLocation
    } else {
        $mkvMergeLocation = $options.mkvMergeOptions.mkvMergeLocation
    }
    if ($($workingFolder.mkvMergeCommand)){
        $mkvMergeCommand = $workingFolder.mkvMergeCommand
    } else {
        $mkvMergeCommandToLaunch = $options.mkvMergeOptions.mkvMergeCommandToLaunch
        $mkvMergeCommandList = $options.mkvMergeOptions.mkvMergeCommandList
        $mkvMergeCommand = $mkvMergeCommandList[$mkvMergeCommandToLaunch]
        #mkvMergeOptions section
        foreach ($mkvMergeCommand in $mkvMergeCommandList){
            LogWrite $DEBUG $("MkvMerge. Preset from json: $mkvMergeCommand")
        }
        LogWrite $DEBUG $("Preset to use: $($mkvMergeCommandList[$mkvMergeCommandToLaunch])")
    }
    if ($($workingFolder.runHandbrake).ToUpper() -eq "True"){
        [bool]$runHandbrake = $true
    } else {
        if ($($workingFolder.runHandbrake).ToUpper() -eq "False"){
            [bool]$runHandbrake = $false
        } else {
            [bool]$runHandbrake = $options.runOptions.runHandbrake
        }
    }
    if ($($workingFolder.runMkvMerge).ToUpper() -eq "True"){
        [bool]$runMkvMerge = $true
    } else {
        if ($($workingFolder.runMkvMerge).ToUpper() -eq "False"){
            [bool]$runMkvMerge = $false
        } else {
            [bool]$runMkvMerge = $options.runOptions.runMkvMerge
        }
    }

    #Get the file list
    $files = Get-ChildItem -Path $($workingFolder.path) -Include $includeList -Recurse -Exclude $excludeFileList 
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
            $workingHandbrakeList = Get-WorkingHandbrakeListCsv
            # Search if currently output file is in a previous run list. If yes, we can treat it as unfinished
            # if not and the file is already on disk, jump to next file
            foreach ($workingHandbrake in $workingHandbrakeList){
                if ($workingHandbrake.file -eq $handbrakeDestinationFile){
                    LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile found from a previous launch. Retrying")
                    $foundWorkingHandbrake = $true
                    continue
                } 
            }
            if (!$foundWorkingHandbrake){
                [bool]$foundHandbrakeDestinationFile = $false
                if ((Test-Path -LiteralPath $handbrakeDestinationFile -PathType leaf)){
                    LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile found. Jumping to next file")
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
                $workingHandbrakeList = Get-WorkingHandbrakeListCsv
                Add-WorkingHandbrakeListCsv $workingHandbrakeList $handbrakeDestinationFile
                LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile added to $workingHandbrakeFileCsv")
            }
            Invoke-Expression $handbrakeNewcommand
            $handbrakeExitCode = $LASTEXITCODE
            LogWrite $INFO $("Handbrake rc: $handbrakeExitCode of $($file.name)")
            if ($handbrakeExitCode -eq 0){
                $workingHandbrakeList = Get-WorkingHandbrakeListCsv
                LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile removed from $workingHandbrakeFileTxt")
                Remove-FromWorkingHandbrakeListCsv $workingHandbrakeList $handbrakeDestinationFile
                LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile removed from $workingHandbrakeFileTxt")
            }
        } else {
            LogWrite $INFO $("Handbrake disabled")
        }

        ### Merge file with command in options
        $inputfile = "$($file.DirectoryName)\$($file.Name)"
        $outputfile = "$($file.DirectoryName + "\1tmp")\$($file.Name.Replace($file.Extension,".mkv"))"
        $mergeNewcommand = $mkvMergeCommand
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


