; ************************
; *    Print in Order    *
; *    By Andrew East    *
; ************************
;
; This is a script to execute a printing operation on a list of serial numbers.
;
; The database program alphabetizes the serial numbers entered by the technician, but it is desirable
; to print them in a specific order. It the serial numbers' ordering stored in SERIALS.txt, which must
; match exactly the serial numbers stored in the database program.

#SingleInstance Ignore ;this directive will not let the user run a second instance of this script accidentally

;set to false to "mock print" only
ENABLE_PRINTING := true
;controls printing of label or worksheet.
ENABLE_LABEL := true
ENABLE_WORKSHEET := true

DEBUG_MODE := true ;show extra information in the dialog boxes
EXPORT_SORTED_SN := true ;export to a text file

RESOURCES_DIR = %A_ScriptDir%\Resources\PrintInOrder ;location where image files exist
SERIALS_FILE = %A_ScriptDir%\SERIALS.txt

SendMode Input



;**********************************************************
;*  Calculate Order The Serial Numbers Will be in the DB  *
;**********************************************************
numSN := 0
dbOrder := Object() ;dbOrder["serialNumber"] = numerical order of SERIALS.txt
printingOrderSNs := Object() ;printingOrderSNs[numerical printing order] = "serialNumber" of that numerical number
printingOrderPositions := Object() ;printingOrderPositions[numerical printing order] = numerical index of alphabetical order, ie. distance from top


;Populate arrays from the text file, storing both the order of the text file and the alphabetical order (ie. DB order)
Loop, Read, %SERIALS_FILE%
{
  dbOrder.Insert(A_LoopReadLine, A_Index) ;Because of how AutoHotKey's .Insert() works, array index (ie. SNs) will be in alphabetical order
  printingOrderSNs[numSN] := A_LoopReadLine ;index will be in numerical order of file line
  numSN++
}

;Populate a new array, which is ordered by SERIALS.txt, and uses the order of dbOrder to store how many down from the first item this item is
downArrows := 0 ;starting at first item, zero from the top
For serialNum, orderedNum in dbOrder
{
  printingOrderPositions[orderedNum] := downArrows
  downArrows++
}

;For Debug: export the sorted s/n's to a text file
if EXPORT_SORTED_SN
{
  fileDelete %RESOURCES_DIR%\AHKOrder.txt
  textToAppend := ""
  for serialNum, orderedNum in dbOrder
    textToAppend := textToAppend . serialNum . "`n"
  fileAppend %textToAppend%, %RESOURCES_DIR%\AHKOrder.txt
  ;msgbox **Debug**`nExported alphabetical-sorted array to AHKOrder.txt
}

/* *** Basic Debug Output ***
output := "dbOrder is:`n"
For serialNum, orderedNum in dbOrder
  output := output . "Key: " . serialNum . " = value: " . orderedNum . "`n"
msgbox % output

output := "printingOrderSNs is:`n"
For serialNum, orderedNum in printingOrderSNs 
  output := output . "Key: " . serialNum . " = value: " . orderedNum . "`n"
msgbox % output

output := "printingOrderPositions is:`n"
For serialNum, orderedNum in printingOrderPositions
  output := output . "Key: " . serialNum . " = value: " . orderedNum . "`n"
msgbox % output
*/



;*****************************************************
;*  All Pre-Calculations Done, Switch to DB Program  *
;*****************************************************

;Delay until DB program is open
DB_PROGRAM := "db_program_title_bar"
DB_PROGRAM_FORM := "db_program_SN_form_title_bar"
Loop,
{
  ;DB program is open to its main screen but SN form is not open
  IfWinExist, %DB_PROGRAM%
  {
    IfWinNotExist, %DB_PROGRAM_FORM%
    {
      IfWinNotActive, %DB_PROGRAM%, WinActivate, , %DB_PROGRAM%
      WinWaitActive, %DB_PROGRAM%
      Send {ALTDOWN}q{ALTUP}
      Sleep 1000
      Break ;exit outer loop loop
    }
    Else ;SN Form already open
    {
      Break
    }
  }
  Else ;winNotExist Main DB window
  {
    MsgBox, 5, Print-SN-in-Order Script, DB Program is not open at all. Please open it and retry or choose cancel.
    IfMsgBox Cancel
      ExitApp
  }
}
WinWait, %DB_PROGRAM_FORM%, 
IfWinNotActive, %DB_PROGRAM_FORM%, , WinActivate, %DB_PROGRAM_FORM%, 
WinWaitActive, %DB_PROGRAM_FORM%, 


;Have user verify # of SNs in SERIALS.txt vs. what they see in DB Program
MsgBox 4, Print-QA-in-Order Script, Your SERIALS.txt contains %numSN% serial numbers in scanned order.`n`nAre there %numSN% on the SN Form?
IfMsgBox No
  ExitApp

stopWatchStart := A_Now
FormatTime, formatStopWatchStart, stopWatchStart, hh:mm:ss tt

TESTTIME = %A_Now%

if !DEBUG_MODE
{
  ;blocks all mouse movement (but not clicks! ...can't block clicks)
  BlockInput MouseMove
  ;blocks ALL mouse/keyboard ONLY while a "Send" command is processing
  BlockInput SendAndMouse
}

;Make sure we switch back to SN Form after the dialog box
WinWait, %DB_PROGRAM_FORM%, 
IfWinNotActive, %DB_PROGRAM_FORM%, , WinActivate, %DB_PROGRAM_FORM%, 
WinWaitActive, %DB_PROGRAM_FORM%, 


;***********************************
;*  Find Scrollbar & Scroll Right  *
;***********************************
;Note: Scrolling right for super-long SNs minimizes delays, because the SN Form normally pops up a tool-tips for each item as it scrolls past
ImageSearch, foundX, foundY, 300, 200, 400, 350, %RESOURCES_DIR%\SNFormneedsToScrollRight2.PNG ;search only part of the window
If (ErrorLevel == 0) ;found the scroll bar
{
  Click %foundX%, %foundY% + 7
  Sleep 250
}
Else
{
  ImageSearch, foundX, foundY, 300, 200, 400, 350, %RESOURCES_DIR%\SNFormneedsToScrollRight3.PNG ;search only part of the window
  If (ErrorLevel == 0) ;found the scroll bar
  {
    Click %foundX%, %foundY% + 7
    Sleep 250
  }
  Else
  {
    ImageSearch, foundX, foundY, 300, 200, 400, 350, %RESOURCES_DIR%\SNFormneedsToScrollRight1.PNG ;search only part of the window
    If (ErrorLevel == 0) ;found the scroll bar
    {
      Click %foundX%, %foundY% + 7
      Sleep 250
    }
  }
}
;not checking for Not Found. Don't care if we don't find the un-scrolled scrollbar; this merely mean we don't need to scroll it.


;*****************************************************************************************************
;*  Find image to locate SN Form, Click on first item, then {Home} to ensure first item is selected  *
;*****************************************************************************************************
ImageSearch, foundX, foundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, %RESOURCES_DIR%\SNForm.PNG
If (ErrorLevel == 0)
{
  foundX += 25
  foundY += 27
  Click %foundX%, %foundY%
  Sleep 100
  Send {Home}
}
Else
{
  MsgBox **DEBUG**`nError: Could not find SN Form title bar.
  if !DEBUG_MODE
  {
    BlockInput MouseMoveOff
    BlockInput Default
  }
  ExitApp
}


;***********************************************
;*  Pre-Locate X, Y of Serial Number text box  *
;***********************************************
serialX := 0
serialY := 0
ImageSearch, serialX, serialY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, %RESOURCES_DIR%\SNFormHiddenTextBox.PNG
If (ErrorLevel == 0)
{
  serialX += 100
  serialY += 5
}
Else
{
  MsgBox Error: Could not find Serial Number text box within SN Form.
  if !DEBUG_MODE
  {
    BlockInput MouseMoveOff
    BlockInput Default
  }
  ExitApp
}


;*******************
;*  Printing Loop  *
;*******************
DetectHiddenText, On
prevDown := 0
currDown := 0
currSN := 0
totalSerialsMatched := 0
dynamicUpDown = {Down}
halfMasterItems := numSN // 2 ; // is Interger (floor) division
delta := 0

countHomeEnd := 0
countDelta := 0

For sorted, downArrows in printingOrderPositions
{
  ;focus has already returned to the list of serial numbers

  ;*********************************************************
  ;*  Calculate delta arrow presses to reach current item  *
  ;*********************************************************
  ;notes:
  ; numSN = total s/n in txt file (1-indexed, ie. 24 items means numSN = 24)
  ; halfMasterItems
  ; downArrows = item's position in DB Program: 0 for first item, numSN for 
  ; prevDown = prev item's position in DB Program
  ; currDown := downArrows - prevDown -- sets currDown to Delta arrow presses from prevDown to downArrows

  currDown := Abs(downArrows - prevDown)

  If (downArrows <= halfMasterItems)
  {
    If (downArrows < currDown) ;closer to the top then delta
    {
      Send {Home}
      Sleep 300
      delta := downArrows
      dynamicUpDown = {Down}
      countHomeEnd++
    }
    Else ;delta is closer
    {
      delta := currDown
      If (downArrows - prevDown) < 0
        dynamicUpDown = {Up}
      Else
        dynamicUpDown = {Down}
      countDelta++
    }
  }
  Else ;(downArrows > halfMasterItems)
  {
    If ((numSN - downArrows - 1) < currDown) ;closer to bottom then delta
    {
      Send {End}
      Sleep 300
      delta := numSN - downArrows - 1
      dynamicUpDown = {Up}
      countHomeEnd++
    }
    Else
    {
      delta := currDown
      If (downArrows - prevDown) < 0
        dynamicUpDown = {Up}
      Else
        dynamicUpDown = {Down}
      countDelta++
    }
  }

  ;Send keyboard presses {up} or {down} until at the correct SN position
  Loop, %delta%
  {
    Send %dynamicUpDown% ;will be {Down} or {Up}
    Sleep 150
  }


  ;Click on the Serial Number text box (in order for AHK to "read" this hidden text from the window manager).
  Sleep 500
  Click %serialX%, %serialY%
  Sleep 200

  Loop,
  {
    ImageSearch, foundX, foundY, 630, 50, 750, 200, %RESOURCES_DIR%\clickedSNTextBox.PNG  ;or 660, 90, 700, 150
    If (ErrorLevel == 0)
    {
      Break ;the S/N box is already selected
    }
    Else
    {
      if DEBUG_MODE
      {
        msgbox, , DEBUG, SN Text Box NOT SELECTED yet.`nWaiting 1 second for DB Program to catch up.`n(This message will close automatically.), 1
      }
      else
      {
        Sleep 1000 ;wait a bit more for DB Program to catch up
      }
    }
  }

  ;Test that the correct SN is currently in SN text box
  WinGetText, foundWinText, %DB_PROGRAM_FORM% ;requires the Serial Number text box to be clicked on, which is already done above.
  IfNotInString, foundWinText, % printingOrderSNs[currSN]
  {
    if !DEBUG_MODE
    {
      BlockInput MouseMoveOff
      BlockInput Default
    }
    MsgBox 4, "Print SN in Order Script", % "#" . currSN + 1 . ". Serial number " . printingOrderSNs[currSN] . " was expected from SERIALS.txt, but DB Program reports " . foundWinText . " has been highlighted.`nYou should be able to highlight the correct S/N in DB Program, then hit yes.`n`nChoose No to quit."
    IfMsgBox No
    {
      ExitApp
    }
    Else
    {
      totalSerialsMatched++
      if !DEBUG_MODE
      {
        BlockInput MouseMove
        BlockInput SendAndMouse
      }
      WinWait, %DB_PROGRAM_FORM%, 
      IfWinNotActive, %DB_PROGRAM_FORM%, , WinActivate, %DB_PROGRAM_FORM%, 
      WinWaitActive, %DB_PROGRAM_FORM%, 
      Sleep 500
      Click %serialX%, %serialY%
      Sleep 200
    }
  }
  Else ;Serial Numbers text box DID match SERIALS.txt. Proceed as expected.
  {
    totalSerialsMatched++
  }

  ;**********************************
  ;*  Print Worksheet and Label  *
  ;**********************************
  if ENABLE_PRINTING 
  {
    Sleep 500
    If ENABLE_LABEL
    {
      Send {ALTDOWN}l{ALTUP}
      sleep 500
      If !ENABLE_WORKSHEET ;return to neutral position
      {
        send {shift down}{tab}{tab}{tab}{shift up}
        sleep 1000
      }
    }

    If ENABLE_WORKSHEET
    {
        Send {ALTDOWN}q{ALTUP}
        Sleep 500
        Send n
        sleep 2000
        send {shift down}{tab}{tab}{tab}{tab}{shift up}
        sleep 1000
    }
    Else If !ENABLE_LABEL ; && !ENABLE_WORKSHEET
    {
      MsgBox Printing is enabled, but both labels and SN worksheets are disabled. Script will now exit.
      ExitApp
    }


  }
  Else ;Mock-printing
  {
    Sleep 500
    MsgBox, 64, **Debug**, PLEASE WAIT ONE SECOND.`nMock-print QA/Label., 1 ;one second timeout
    sleep 1000
    Send {ALTDOWN}q{ALTUP}
    Sleep 500
    Send {tab}{Space} ; choose Cancel
    sleep 2000
    send {shift down}{tab}{tab}{tab}{tab}{shift up}
    sleep 1000

    ;Sleep 500
    ;WinWait, %DB_PROGRAM_FORM%, 
    ;IfWinNotActive, %DB_PROGRAM_FORM%, , WinActivate, %DB_PROGRAM_FORM%, 
    ;WinWaitActive, %DB_PROGRAM_FORM%, 
    ;Sleep 500
    ;Click %serialX%, %serialY%
    ;Sleep 500
    ;Send +{tab} ;return to items listbox
  }

  prevDown := downArrows
  currSN++
} ;Loop to next serial number



if !DEBUG_MODE
{
  BlockInput MouseMoveOff
  BlockInput Default
}


If !DEBUG_MODE
{
  MsgBox Script Completed!`n`nPrinted %totalSerialsMatched% items successfully.
}
Else
{
  ;FormatTime, stopWatchStart, stopWatchStart, hh:mm:ss tt
  StopWatchEnd := A_Now
  FormatTime, formatStopWatchEnd, stopWatchEnd, hh:mm:ss tt ;current time
  MsgBox Script Completed!`n`nDebug Info:`ntotalSerialsMatched = %totalSerialsMatched%`n`nStopWatchStart = %formatStopWatchStart%`nStopWatchEnd =  %formatStopWatchEnd%`n`ncountHomeEnd = %countHomeEnd%`ncountDelta = %countDelta%`n`nTESTTIME = %TESTTIME%
}
ExitApp