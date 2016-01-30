# This script is based on a function that watches a folder and then logs file changes: http://stackoverflow.com/questions/29066742/watch-file-for-changes-and-run-command-with-powershell
# I have modified the conditions of the FileSystemWatcher, allowed for relative paths, allowed changing the output folder, and passing data into the ScriptBlock

# To run this script from a shortcut, make the target of the shortcut:
#   C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -File ".\sassWatch.ps1"

$InputFolder = ".\resources\sass\"
$OutputFolder = ".\resources\css\"
$Filter = "*.scss"

write-Host "`n******************** Now watching $InputFolder - Press the 'q' key to quit! ********************"
write-Host "(Note: Duplicate events can be expected.)"
# Known "bug/feature" of FileSystemWatcher: some events will fire twice due to OS's complexity. As far as I could find out, the best way to filter the duplicate event is to examine the timestamps, perhaps. not worth the cost/benefits. see: http://weblogs.asp.net/ashben/31773


# relative paths must be resolved to absolute for Watcher
$InputFolder = Resolve-Path $InputFolder
$OutputFolder = Resolve-Path $OutputFolder

# this is the ScriptBlock that will get executed every time the $Watcher sees a change
# (a scriptblock is sort of an anonymous function or lambda)
# available data is: 
#   $Event.SourceEventArgs.Name - just the name of the file that was modified
#   $Event.SourceEventArgs.FullPath - full path of file
#   $Event.MessageData - the string that I passed in via "Register-ObjectEvent -MessageData"
$ScriptBlock = [scriptblock]::Create('
  Write-Host -NoNewline $(Get-Date) $Event.SourceEventArgs.Name "was changed";
  sass $Event.SourceEventArgs.FullPath ("{0}{1}.css" -f $Event.MessageData,(Get-Item $Event.SourceEventArgs.FullPath).BaseName)
  Write-Host (" - {0}.css was written" -f (Get-Item $Event.SourceEventArgs.FullPath).BaseName);
')

# set up FileSystemWatcher, and conditions for when it will trigger
$Watcher = New-Object IO.FileSystemWatcher $InputFolder, $Filter -Property @{
    IncludeSubdirectories = $false
    EnableRaisingEvents = $true
    NotifyFilter = [System.IO.NotifyFilters]::LastWrite
}

# register the event with this PowerShell session
$onChange = Register-ObjectEvent $Watcher Changed -Action $ScriptBlock -MessageData $OutputFolder #messageData property I read about on: http://www.ravichaganti.com/blog/passing-variables-or-arguments-to-an-event-action-in-powershell/

# wait loop to leave powershell open until user presses key
# http://powershell.com/cs/forums/t/8696.aspx
while ($true) {
  Start-Sleep -Milliseconds 500
  if ($Host.UI.RawUI.KeyAvailable -and ("q" -eq $Host.UI.RawUI.ReadKey("IncludeKeyUp,NoEcho").Character)) {
      break;
  }
}

# remove event
# Note: this is only actually necessary the powershell session is left open afterwards. see: https://technet.microsoft.com/en-us/library/hh849896.aspx#sectionSection7
write-Host "****************** Unregistering event ID# $($onChange.Id) ******************`n"
Unregister-Event -SubscriptionId $onChange.Id
