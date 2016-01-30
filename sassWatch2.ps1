# http://stackoverflow.com/questions/29066742/watch-file-for-changes-and-run-command-with-powershell

# known "bug/feature" of FileSystemWatcher: some events will fire twice due to OS's complexity. As far as I could find out, the best way to filter the duplicate event is to examine the timestamps, perhaps. not worth the cost/benefits. see: http://weblogs.asp.net/ashben/31773

$InputFolder = ".\resources\sass\"
$OutputFolder = ".\resources\css\"

#function Wait-SassChange {
#  param(
#      [string]$InputFolder,
#      [string]$OutputFolder
#  )  

  write-Host "`n******************** Now watching $InputFolder - Press the 'q' key to quit! ********************"
  write-Host "(Note: Duplicate events can be expected.)"

  $Filter = "*.scss"
  
  #need to resolve relative paths to absolute for Watcher
  $InputFolder = Resolve-Path $InputFolder
  $OutputFolder = Resolve-Path $OutputFolder
  
  
  $ScriptBlock = [scriptblock]::Create('
    Write-Host -NoNewline $(Get-Date) $Event.SourceEventArgs.Name "was changed";
    sass $Event.SourceEventArgs.FullPath ("{0}{1}.css" -f $Event.MessageData,(Get-Item $Event.SourceEventArgs.FullPath).BaseName)
    Write-Host (" - {0}.css was written" -f (Get-Item $Event.SourceEventArgs.FullPath).BaseName);
  ')

  $Watcher = New-Object IO.FileSystemWatcher $InputFolder, $Filter -Property @{
      IncludeSubdirectories = $false
      EnableRaisingEvents = $true
      NotifyFilter = [System.IO.NotifyFilters]::LastWrite
  }
  $onChange = Register-ObjectEvent $Watcher Changed -Action $ScriptBlock -MessageData $OutputFolder #messageData from: http://www.ravichaganti.com/blog/passing-variables-or-arguments-to-an-event-action-in-powershell/

  #http://powershell.com/cs/forums/t/8696.aspx
  while ($true) {
    Start-Sleep -Milliseconds 500
    if ($Host.UI.RawUI.KeyAvailable -and ("q" -eq $Host.UI.RawUI.ReadKey("IncludeKeyUp,NoEcho").Character)) {
        break;
    }
  }

  write-Host "****************** Unregistering event ID# $($onChange.Id) ******************`n"
  Unregister-Event -SubscriptionId $onChange.Id
#}

#Wait-SassChange -InputFolder $InputFolder -OutputFolder $OutputFolder