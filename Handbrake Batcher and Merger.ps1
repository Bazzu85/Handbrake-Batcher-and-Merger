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
                If (($OtherCmdLine.Contains("HandBrake.Worker.exe")) -And ($OtherPID -ne $PID) -and (!$found)){
                    LogWrite $INFO $("PID [$OtherPID] is already running (HandBrake.Worker.exe). Waiting $waitSeconds seconds")
                    $found = $true
                }
                If (($OtherCmdLine.Contains("HandBrakeCLI.exe")) -And ($OtherPID -ne $PID) -and (!$found)){
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
    #Default $globalOptions object
    $globalOptions = [ordered]@{
        runOptions= [ordered]@{
            debugLog = $false
            runHandbrake = $true
            runMkvMerge = $true
            waitSecondsOption = 60
        }
        handbrakeOptions= [ordered]@{
            handbrakePreset="Convert to h265 Medium 720p (only video)"
            handbrakePresetLocation="C:\Users\elbaz\AppData\Roaming\HandBrake\presets.json"
            handbrakeCommand='HandBrakeCLI.exe --preset-import-file "||handbrakePresetLocation||" -Z "||handbrakePreset||" -i "||inputFile||" -o "||outputFile||" --auto-anamorphic'
        }
        mkvMergeOptions= [ordered]@{
            mkvMergeCommand = '"||mkvMergeLocation||" --ui-language en --output ^"||outputFileName||^" --language 0:und --compression 0:none --no-track-tags  --no-global-tags ^"^(^" ^"||handbrakeFileName||^" ^"^)^" --no-video ^"^(^" ^"||inputFileName||^" ^"^)^"'
            mkvMergeLocation = "C:\Program Files\MKVToolNix\mkvmerge.exe"
        }
        fileFolderOptions= [ordered]@{
            includeList=@("*.mp4",
                        "*.mkv"
                            )
            excludeFileList=@("*_handbrake.mkv*")
            excludeFileFolderList=@("\1ok","\1tmp","\1wrn")
        }
    }
    # If not found, create the .config file
    if (!(Test-Path -Path $optionsFile)){
        ConvertTo-Json -InputObject $globalOptions | Out-File $optionsFile
    }
    #region Read it and build the $globalOptions object
    $globalOptionsFromJson = Get-Content -LiteralPath $optionsFile | ConvertFrom-Json
    if ($null -ne $globalOptionsFromJson.runOptions.debugLog){
        $globalOptions.runOptions.debugLog = $globalOptionsFromJson.runOptions.debugLog
    }
    if ($null -ne $globalOptionsFromJson.runOptions.runHandbrake){
        $globalOptions.runOptions.runHandbrake = $globalOptionsFromJson.runOptions.runHandbrake
    }
    if ($null -ne $globalOptionsFromJson.runOptions.runMkvMerge){
        $globalOptions.runOptions.runMkvMerge = $globalOptionsFromJson.runOptions.runMkvMerge
    }
    if ($globalOptionsFromJson.runOptions.waitSecondsOption){
        $globalOptions.runOptions.waitSecondsOption = $globalOptionsFromJson.runOptions.waitSecondsOption
    }
    if ($globalOptionsFromJson.handbrakeOptions.handbrakePreset){
        $globalOptions.handbrakeOptions.handbrakePreset = $globalOptionsFromJson.handbrakeOptions.handbrakePreset
    }
    if ($globalOptionsFromJson.handbrakeOptions.handbrakePresetLocation){
        $globalOptions.handbrakeOptions.handbrakePresetLocation = $globalOptionsFromJson.handbrakeOptions.handbrakePresetLocation
    }
    if ($globalOptionsFromJson.handbrakeOptions.handbrakeCommand){
        $globalOptions.handbrakeOptions.handbrakeCommand = $globalOptionsFromJson.handbrakeOptions.handbrakeCommand
    }
    if ($globalOptionsFromJson.mkvMergeOptions.mkvMergeCommand){
        $globalOptions.mkvMergeOptions.mkvMergeCommand = $globalOptionsFromJson.mkvMergeOptions.mkvMergeCommand
    }
    if ($globalOptionsFromJson.mkvMergeOptions.mkvMergeLocation){
        $globalOptions.mkvMergeOptions.mkvMergeLocation = $globalOptionsFromJson.mkvMergeOptions.mkvMergeLocation
    }
    if ($globalOptionsFromJson.fileFolderOptions.includeList){
        $globalOptions.fileFolderOptions.includeList = $globalOptionsFromJson.fileFolderOptions.includeList
    }
    if ($globalOptionsFromJson.fileFolderOptions.excludeFileList){
        $globalOptions.fileFolderOptions.excludeFileList = $globalOptionsFromJson.fileFolderOptions.excludeFileList
    }
    if ($globalOptionsFromJson.fileFolderOptions.excludeFileFolderList){
        $globalOptions.fileFolderOptions.excludeFileFolderList = $globalOptionsFromJson.fileFolderOptions.excludeFileFolderList
    }
    #endregion

    # Write the $globalOptions object to json with missing default data
    ConvertTo-Json -InputObject $globalOptions | Out-File $optionsFile

    return $globalOptions
}
Function Get-WorkingFolderList {

    if (!(Test-Path -LiteralPath "$PSScriptRoot\configuration")){
        New-Item -Path ("$PSScriptRoot\configuration") -ItemType Directory -Force
    }
    $defaultWorkingFolderList = [ordered]@{
        path = $PSScriptRoot
        handbrakePresetLocation="C:\Users\elbaz\AppData\Roaming\HandBrake\presets.json"
        handbrakePreset="Convert to h265 Medium 720p (only video)"
        handbrakeCommand='HandBrakeCLI.exe --preset-import-file "||handbrakePresetLocation||" -Z "||handbrakePreset||" -i "||inputFile||" -o "||outputFile||" --auto-anamorphic'
        mkvMergeLocation="C:\Program Files\MKVToolNix\mkvmerge.exe"
        mkvMergeCommand='"||mkvMergeLocation||" --ui-language en --output ^"||outputFileName||^" --language 0:und --compression 0:none --no-track-tags  --no-global-tags ^"^(^" ^"||handbrakeFileName||^" ^"^)^" --no-video ^"^(^" ^"||inputFileName||^" ^"^)^"'
        runHandbrake=$true
        runMkvMerge=$true
    }
    # Check if $workingFolderListCsvFile is present. 
    # If not create it
    # If yes read all folders in array
    if (!(Test-Path -LiteralPath $workingFolderListCsvFile -PathType leaf)){
        $defaultWorkingFolderList | Export-Csv -Path $workingFolderListCsvFile
    } else {
        $workingFolderListFromCsv = Import-Csv -Path $workingFolderListCsvFile
    }
    
    #region Rebuild the csv adding missing columns not found from Import. After that read the Csv again
    $i=0
    foreach ($workingFolderFromCsv in $workingFolderListFromCsv){
        $newRow = $defaultWorkingFolderList
        if ($null -ne $workingFolderFromCsv.path){
            $newRow.path = $workingFolderFromCsv.path
        }
        if ($null -ne $workingFolderFromCsv.handbrakePresetLocation){
            $newRow.handbrakePresetLocation = $workingFolderFromCsv.handbrakePresetLocation
        }
        if ($null -ne $workingFolderFromCsv.handbrakePreset){
            $newRow.handbrakePreset = $workingFolderFromCsv.handbrakePreset
        }
        if ($null -ne $workingFolderFromCsv.handbrakeCommand){
            $newRow.handbrakeCommand = $workingFolderFromCsv.handbrakeCommand
        }
        if ($null -ne $workingFolderFromCsv.mkvMergeLocation){
            $newRow.mkvMergeLocation = $workingFolderFromCsv.mkvMergeLocation
        }
        if ($null -ne $workingFolderFromCsv.mkvMergeCommand){
            $newRow.mkvMergeCommand = $workingFolderFromCsv.mkvMergeCommand
        }
        if ($null -ne $workingFolderFromCsv.runHandbrake){
            $newRow.runHandbrake = $workingFolderFromCsv.runHandbrake
        }
        if ($null -ne $workingFolderFromCsv.runMkvMerge){
            $newRow.runMkvMerge = $workingFolderFromCsv.runMkvMerge
        }
        if ($i -eq 0){
            $newRow | Export-Csv -Path $workingFolderListCsvFile -Force    
        } else {
            $newRow | Export-Csv -Path $workingFolderListCsvFile -Append -Force
        }
        $i++
    }
    $workingFolderList = Import-Csv -Path $workingFolderListCsvFile
    #endregion
    return $workingFolderList
}
Function Get-WorkingHandbrakeList {
    <#
    .SYNOPSIS
        Check if $workingHandbrakeFileCsv exists.
        If not, create it.
        If yes, read it and return the content
    #>
    if (!(Test-Path -Path $workingHandbrakeFileCsv)){
        LogWrite $DEBUG $("File $workingHandbrakeFileCsv not found")
        $workingHandbrakeList = $null
    } else {
        LogWrite $DEBUG $("File $workingHandbrakeFileCsv found. Importing CSV")
        $workingHandbrakeList = Import-Csv -LiteralPath $workingHandbrakeFileCsv
    }
    return $workingHandbrakeList
}

Function AddToWorkingHandbrakeList ($workingHandbrakeListCsv, $handbrakeDestinationFile) {
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

Function RemoveFromWorkingHandbrakeList ($workingHandbrakeListCsv, $handbrakeDestinationFile) {
    # remove the handbrakeDestinationFile from file
    $newWorkingHandbrakeListCsv = $workingHandbrakeListCsv | Where-Object {$_.file -ne $handbrakeDestinationFile}
    if (!$newWorkingHandbrakeListCsv){
        Remove-Item -LiteralPath $workingHandbrakeFileCsv -Force
    } else {
        $newWorkingHandbrakeListCsv | Export-Csv -Path $workingHandbrakeFileCsv
    }
}

Function Get-NewOptions ($globalOptions , $workingFolder){
    # Set the variable from custom folder csv or from generic configuration
    $newOptions = [PSCustomObject]@{
        handbrakePresetLocation = ''
        handbrakePreset = ''
        handbrakeCommand = ''
        mkvMergeLocation = ''
        mkvMergeCommand = ''
        runHandbrake = ''
        runMkvMerge = ''
    }    
    #region Build the $newOptions object merging the infos from csv and the global options json
    if ($($workingFolder.handbrakePresetLocation)){
        $newOptions.handbrakePresetLocation = $workingFolder.handbrakePresetLocation
    } else {
        $newOptions.handbrakePresetLocation = $globalOptions.handbrakeOptions.handbrakePresetLocation
    }
    if ($($workingFolder.handbrakePreset)){
        $newOptions.handbrakePreset = $workingFolder.handbrakePreset
    } else {
        $newOptions.handbrakePreset = $globalOptions.handbrakeOptions.handbrakePreset
    }
    if ($($workingFolder.handbrakeCommand)){
        $newOptions.handbrakeCommand = $workingFolder.handbrakeCommand
    } else {
        $newOptions.handbrakeCommand = $globalOptions.handbrakeOptions.handbrakeCommand
    }
    if ($($workingFolder.mkvMergeLocation)){
        $newOptions.mkvMergeLocation = $workingFolder.mkvMergeLocation
    } else {
        $newOptions.mkvMergeLocation = $globalOptions.mkvMergeOptions.mkvMergeLocation
    }
    if ($($workingFolder.mkvMergeCommand)){
        $newOptions.mkvMergeCommand = $workingFolder.mkvMergeCommand
    } else {
        $newOptions.mkvMergeCommand = $globalOptions.mkvMergeOptions.mkvMergeCommand
    }
    if ($($workingFolder.runHandbrake).ToUpper() -eq "True"){
        [bool]$newOptions.runHandbrake = $true
    } else {
        if ($($workingFolder.runHandbrake).ToUpper() -eq "False"){
            [bool]$newOptions.runHandbrake = $false
        } else {
            [bool]$newOptions.runHandbrake = $globalOptions.runOptions.runHandbrake
        }
    }
    if ($($workingFolder.runMkvMerge).ToUpper() -eq "True"){
        [bool]$newOptions.runMkvMerge = $true
    } else {
        if ($($workingFolder.runMkvMerge).ToUpper() -eq "False"){
            [bool]$newOptions.runMkvMerge = $false
        } else {
            [bool]$newOptions.runMkvMerge = $globalOptions.runOptions.runMkvMerge
        }
    }
    #endregion
    LogWrite $INFO $("Options for current working folder")
    LogWrite $INFO $("HandBrake. handbrakePreset        : $($newOptions.handbrakePreset)")
    LogWrite $INFO $("HandBrake. handbrakePresetLocation: $($newOptions.handbrakePresetLocation)")
    LogWrite $INFO $("HandBrake. handbrakeCommand       : $($newOptions.handbrakeCommand)")
    LogWrite $INFO $("MkvMerge . mkvMergeLocation       : $($newOptions.mkvMergeLocation)")
    LogWrite $INFO $("MkvMerge . mkvMergeCommand        : $($newOptions.mkvMergeCommand)")
    LogWrite $INFO $("MkvMerge . mkvMergeLocation       : $($newOptions.mkvMergeLocation)")
    LogWrite $INFO $("Global   . runHandbrake           : $($newOptions.runHandbrake)")
    LogWrite $INFO $("Global   . runMkvMerge            : $($newOptions.runMkvMerge)")

    return $newOptions
}

Function CreateFolder ($path){
    if (!(Test-Path -Path $path)){
        New-Item -Path ($path) -ItemType Directory -Force
    }
}

#Main Code

#Set title
$host.UI.RawUI.WindowTitle = $MyInvocation.MyCommand.Name.Replace(".ps1","")

Write-Host "Handbrake Batcher and Merger by Bazzu v.2.4"
#Write-Host "Powershell infos: "
#$host

# Get name of current script
$ScriptName = $MyInvocation.MyCommand.Name 

# Set log types
$DEBUG = "DEBUG"
$INFO  = "INFO "

# Create folders if missing
CreateFolder "$PSScriptRoot\configuration"
CreateFolder "$PSScriptRoot\log"

# Set file paths
$optionsFile = "$PSScriptRoot\configuration\options.json"
$workingFolderListCsvFile = "$PSScriptRoot\configuration\Working Folder List.csv"
$workingHandbrakeFileTxt = "$PSScriptRoot\Working Handbrake File List.txt"
$workingHandbrakeFileCsv = "$PSScriptRoot\configuration\Working Handbrake File List.csv"
$date = Get-Date -Format "yyyyMMdd"
$logFile = "$PSScriptRoot\log\$($MyInvocation.MyCommand.Name.Replace('.ps1',''))_$date.log"

# if the base configuration files are missing, remember to abort execution later
if (!(Test-Path -LiteralPath $optionsFile -PathType leaf) -or !(Test-Path -LiteralPath $workingFolderListCsvFile -PathType leaf)){
    [bool]$abortExecution = $true
} else {
    [bool]$abortExecution = $false
}

# Get options and working folders from files
$globalOptions = Get-OptionsFromJson
$workingFolderList = Get-WorkingFolderList

if ($abortExecution){
    Write-Host "Generated missing default configuration files in $PSScriptRoot\configuration. Please review it and launch again"
    pause
    exit
}

# Set the working variables from options.json
[bool]$debugLog = $globalOptions.runOptions.debugLog
$waitSeconds = $globalOptions.runOptions.waitSecondsOption
$includeList = $globalOptions.FileFolderOptions.includeList
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
$excludeFileList = $globalOptions.FileFolderOptions.excludeFileList
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
$excludeFileFolderList = $globalOptions.FileFolderOptions.excludeFileFolderList

# Check if an instance is already running
Test-IfAlreadyRunning -ScriptName $ScriptName
LogWrite $INFO $("(PID=[$PID]) This is the 1st and only instance allowed to run") #this only shows in one instance

# Work on all folder in $workingFolderList
foreach ($workingFolder in $workingFolderList){
    # if the folder doesn't exists , jump to the next
    if (!(Test-Path -Path $($workingFolder.path))){
        LogWrite $DEBUG $("Folder $($workingFolder.path) not found. Skipping")
        continue
    }
    LogWrite $INFO $("Running script in folder $($workingFolder.path)")
    $newOptions = Get-NewOptions $globalOptions $workingFolder

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
        if ($newOptions.runHandbrake){
            $workingHandbrakeList = Get-WorkingHandbrakeList
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
        $handbrakeNewcommand = $newOptions.handbrakeCommand
        $handbrakeNewcommand = $handbrakeNewcommand.Replace("||handbrakePresetLocation||", $newOptions.handbrakePresetLocation)
        $handbrakeNewcommand = $handbrakeNewcommand.Replace("||handbrakePreset||", $newOptions.handbrakePreset)
        $handbrakeNewcommand = $handbrakeNewcommand.Replace("||inputFile||", $file)
        $handbrakeNewcommand = $handbrakeNewcommand.Replace("||outputFile||", $handbrakeDestinationFile)

        LogWrite $INFO $("Handbrake command: $handbrakeNewcommand")
        if ($newOptions.runHandbrake){
            if (!$foundWorkingHandbrake){
                $workingHandbrakeList = Get-WorkingHandbrakeList
                AddToWorkingHandbrakeList $workingHandbrakeList $handbrakeDestinationFile
                LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile added to $workingHandbrakeFileCsv")
            }
            Invoke-Expression $handbrakeNewcommand
            $handbrakeExitCode = $LASTEXITCODE
            LogWrite $INFO $("Handbrake rc: $handbrakeExitCode of $($file.name)")
            if ($handbrakeExitCode -eq 0){
                $workingHandbrakeList = Get-WorkingHandbrakeList
                LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile removed from $workingHandbrakeFileTxt")
                RemoveFromWorkingHandbrakeList $workingHandbrakeList $handbrakeDestinationFile
                LogWrite $DEBUG $("Handbrake file $handbrakeDestinationFile removed from $workingHandbrakeFileTxt")
            }
        } else {
            LogWrite $INFO $("Handbrake disabled")
        }

        ### Merge file with command in globalOptions
        $inputfile = "$($file.DirectoryName)\$($file.Name)"
        $outputfile = "$($file.DirectoryName + "\1tmp")\$($file.Name.Replace($file.Extension,".mkv"))"
        $mergeNewcommand = $newOptions.mkvMergeCommand
        $mergeNewcommand = $mergeNewcommand.Replace("||mkvMergeLocation||", $newOptions.mkvMergeLocation)
        $mergeNewcommand = $mergeNewcommand.Replace("||outputFileName||", $outputfile)
        $mergeNewcommand = $mergeNewcommand.Replace("||inputFileName||", $inputfile)
        $mergeNewcommand = $mergeNewcommand.Replace("||handbrakeFileName||", $handbrakeDestinationFile)

        $outputfile =  $outputfile.Replace("\\","\")

        LogWrite $INFO $("Merging to $outputfile")
        LogWrite $INFO $("Mkvmerge command: cmd /c $mergeNewcommand")
        if ($newOptions.runMkvMerge){
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


