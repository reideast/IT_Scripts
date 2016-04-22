;Basic setup
CoordMode Mouse, Relative ;Coordinates are relative to Top Left of windows, not screen
SendMode Input ;Keyboard and mouse inputs cannot be interrupted by human

;Set up some properies
FileIDList = %A_ScriptDir%\ID.txt ;should usually be equal to "E:\ID.txt"
FileWebDBShortcut = %A_Desktop%\WebDB.url
WindowWebDB := "WebDB - Windows Internet Explorer"
WindowForm := "Crystal Reports ActiveX Viewer - Windows Internet Explorer"
WindowLabelPre := "http://example.com/WebDB/pdf/"
WindowLabelPost := ".pdf - Windows Internet Explorer"
CountForms := 0



;Verify ID file
IfNotExist %FileIDList%
{
  MsgBox, 48, Form Print Script, ID List File "%FileIDList%" does not exist within the script's directory.
  return
}
FileGetTime CurrFileDate, FileIDList, M ;m for modified time
FormatTime CurrFileDate, CurrFileDate, YDay
if (CurrFileDate != A_YDay) ;A_YDay is this computer's clock's day of the year
{
  MsgBox, 308, Form Print Script, ID List File has not been modified today.`nDo you want to continue?
  ;48 for Icon Exclamation, 4 for Yes/No, and 256 to make the second button, No, the default += 308
  IfMsgBox No
    return
}




;Verify or open WebDB page
IfWinExist %WindowWebDB%
{
  WinActivate %WindowWebDB%
}
else
{
  ;MsgBox %FileWebDBShortcut% ;**DEBUG**
  Run %FileWebDBShortcut%
  WinWait %WindowWebDB%
  MsgBox 4, Form Print Script, Please choose Yes when the WebDB page has loaded. Click No to exit.
  IfMsgBox No
    return
}




;Loop through the ID File
Loop, Read, %FileIDList%
{
  ++CountForms

  ;**Seems like this isn't needed, the Loop, Read seems to skip blank lines at the end of the file?
  ;**also, blank lines in the middle are errors, and will be cautght by the regex
  ;  ;Validate that this is not a blank line
  ;  if (A_FileReadLine == "")
  ;  {
  ;    MsgBox Line %CountForms% was blank! ;**debug**
  ;    continue ;skip this line, go on the the next
  ;  }

  ;Validate that A_LoopReadLine conforms to format: 00######
  FoundPos := RegExMatch(A_LoopReadLine, "^0{2}\d{6}$")
  If FoundPos != 1
  {
    MsgBox, 52, Form Print Script, Line %CountForms%: "%A_LoopReadLine%", is not a valid ID number.`nDo you want to skip it?
    ;48 for Icon Exclamation, 4 for Yes/No
    IfMsgBox Yes
      continue
  }
  


  ;Print Form Sheet
  WinActivate %WindowWebDB%
  Sleep 200
  Click 70, 310 ;click on "Enter ID Number"
  Sleep 200
  Send {Home}
  Sleep 200
  Send +{End} ;Shift+End to select all
  Sleep 200
  Send %A_LoopReadLine% ;"paste" current ID #
  Sleep 200
  Send {Enter} ;Will cause web page to open %WindowForm%
  
  ;Note: While the WindowForm is loading, it has a different window title.
  ;It isn't until the crystal reports has fully loaded that the name == %WindowForm%.
  ;That's why I don't have to look for the Form to fully load with any trickery.
  WinWait %WindowForm%
  Sleep 2000 ;wait for crystal reports to full load

  ;**DEBUG** Wiggle the mouse to solve that it's not clicking on Print about 1/5 of the time.
  MouseMove 48, 150
  Sleep 200
  MouseMove 53, 150
  Sleep 200

  Sleep 300
  Click 50, 150 ;click on Print icon
  
  Loop
  {
    Sleep 300
    IfWinExist Print
    {
      Break ;exit win-wait-until-exist loop
    }
    Else
    {
      Sleep 2000 ;wait a full 2 seconds before clicking print again
      MouseMove 48, 150
      Sleep 200
      MouseMove 53, 150
      Sleep 200
      Click 50, 150 ;click on Print icon
    }
  }


  ;Print dialog is now open!
  Sleep 200
  Send {Enter}
  Sleep 500


  ;Wait until last page has printed
  ;  This works because Crystal Reports literally go through each page while printing,
  ;  so by looking when the "next arrow" is grayed out, we can see that it's on the last page.
  ArrowColor = 0xFFFFFF ;set up variable outside of loop for efficiency...well, this is not required because all vars are global in ahk...
  Loop
  {
    Sleep 200
    MouseMove 438, 153 ;necessary trickery: jiggle the mouse
    Sleep 200
    WinActivate %WindowForm% ;necessary trickery: refocus on WindowForm
    Sleep 1000
    PixelGetColor, ArrowColor, 438, 153
    If (ArrowColor == "0x808080") ;therefore, the arrow should be 0x808080, or light grey, and we are on the last page
    {
      break ;exit color loop
    }
    Else ;black, so there is still a next page
    {
      Sleep 3000 ;wait before checking again
    }
  }
  Sleep 3000 ;wait to account for last page finishing

   
  WinClose %WindowForm%
  WinActivate %WindowWebDB%
  
  Sleep 1000 ;**Debug** wait a full second until printing label**

  ;Print Shipping Label
  Sleep 200
  Click 70, 440 ;click  on "ShippingID"
  Sleep 200
  Send {Home}
  Sleep 200
  Send +{End} ;select all
  Sleep 200
  Send %A_LoopReadLine% ;"paste" current ID #
  Sleep 200
  Click 190, 440 ;click on Preview
  
  WinWait %WindowLabelPre%%A_LoopReadLine%%WindowLabelPost%
  Sleep 500
Sleep 10000 ;**DEBUG**
  Send ^p
  Sleep 3000
  Send {PgUp}
  Sleep 3000
  Send {Enter}
  Sleep 10000 ;1 sec
  WinClose %WindowLabelPre%%A_LoopReadLine%%WindowLabelPost%

} ;END file read line loop


MsgBox, 64, Form Print Script, Printed %CountForms% Form Sheets and Labels. ;0 for OK button, 64 for Info icon
