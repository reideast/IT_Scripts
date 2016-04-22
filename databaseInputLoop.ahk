IfWinExist, Database File - Microsoft Excel
  IfWinNotActive, Database File - Microsoft Excel, WinActivate, , Database File - Microsoft Excel
  WinWaitActive, Database File - Microsoft Excel

Loop, Read, SERIALS.txt
{
  Send %A_LoopReadLine%
  Send {Enter}
  Sleep 1000
  Send {atldown}n{altup}y
  Sleep 2000

  IfWinExist, Database File - Microsoft Excel
    IfWinNotActive, Database File - Microsoft Excel, WinActivate, , Database File - Microsoft Excel
    WinWaitActive, Database File - Microsoft Excel
}
