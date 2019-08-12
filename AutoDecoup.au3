#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

#RequireAdmin
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <Misc.au3>
#include <GUIConstantsEx.au3>
#include <File.au3>
#include <Array.au3>
#include <WinAPIFiles.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <ScrollBarsConstants.au3>
#include <MsgBoxConstants.au3>
#include <GuiEdit.au3>
#include <Clipboard.au3>
#include <StaticConstants.au3>
#include <ScreenCapture.au3>
#include <WinAPIGdi.au3>
#include <WinAPIGdiDC.au3>
#include <WinAPIHObj.au3>
#include <WinAPISysWin.au3>

Global $reason = "Thanks for using this program, see you next time."
Global $folder = @MyDocumentsDir

HotKeySet ("{ESC}","ExitScript")

#Region ### START Main GUI ###
Global $Form1 = GUICreate("AutoDecoup", 197, 268, 192, 124)
WinSetOnTop($Form1, "", 1)
Global $journal = GUICtrlCreateEdit("", 0, 0, 196, 209, BitOR($ES_AUTOVSCROLL,$ES_READONLY,$WS_VSCROLL))
GUICtrlSetData(-1, "")
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlSetCursor (-1, 2)
$GUI_EVENT_FOLDER = GUICtrlCreateButton("//", 170, 215, 25, 25)
$scrollTimes = GUICtrlCreateInput(10, 170, 240, 25, 25)
GUICtrlSetTip (-1, "Auto-scroll power (0=disable)")
$nameInputBox = GUICtrlCreateInput("", 57, 215, 80, 25)
GUICtrlSetState(-1, $GUI_FOCUS)
GUICtrlSetTip (-1, "Target name")
$leftRightOffset = GUICtrlCreateInput(100, 5, 215, 25, 25)
GUICtrlSetTip (-1, "Left and Right offset")
$topBottomOffset = GUICtrlCreateInput(120, 5, 240, 25, 25)
GUICtrlSetTip (-1, "Top and Bottom offset")
$GUI_EVENT_START = GUICtrlCreateButton("Start", 60, 240, 75, 25)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0x008080)
Opt("GUICoordMode", 2)
GUISetCoord(1153, 231)
GUISetState(@SW_SHOW)
#EndRegion ### END Main GUI ###

Requierments()

While 1
   $nMsg = GUIGetMsg()
   Switch $nMsg
	  Case $GUI_EVENT_CLOSE
		 ExitScript()
	  Case $GUI_EVENT_START
		 CaptureRegion()
	  Case $GUI_EVENT_FOLDER
		 ChangeFolder()
   EndSwitch
WEnd

Func Requierments()
   info("Program ready, choose an image name and press Start")
   info("This program is designed to be used with Tab & Enter mostly for speed purpoise")
   info("Current folder -> " & $folder)
EndFunc

Func Scroll()
   $times = GUICtrlRead($scrollTimes)
   MouseWheel($MOUSE_WHEEL_DOWN, $times)
EndFunc

Func ChangeFolder()
   $newFolder = InputBox("Pictures target folder", $folder)
   If Not @error Then
	  $folder = $newFolder
	  info("============================")
	  info("New Folder -> " & $newFolder)
	  info("============================")
   EndIf
EndFunc

Func CaptureRegion()
   Local $hDLL = DllOpen("user32.dll")

   $name = GUICtrlRead($nameInputBox)
   $xOffset = GUICtrlRead($leftRightOffset)
   $yOffset = GUICtrlRead($topBottomOffset)

   GUICtrlSetState($GUI_EVENT_START, $GUI_DISABLE)
   info("============================")
   info("Capturing : " & $name & '.jpg, align the mouse cursor in the center of the picture & press "Print screen" or "Right click" to Cancel.')
   info("============================")
   While 1
	  If _IsPressed("2C") Then
		 Local $mouse = MouseGetPos()
		 $hHBITMAP = _ScreenCapture_Capture($folder & "\" & $name & ".jpg", $mouse[0] - $xOffset, $mouse[1] - $yOffset, $mouse[0] + $xOffset, $mouse[1] + $yOffset, False)
		 If Not @error Then
;~ 			ShellExecute($folder & "\" & $name & ".jpg")
			info("Capture " & $name & " -> Success")
			Scroll()
			GUICtrlSetState($nameInputBox, $GUI_FOCUS)
			ExitLoop
		 Else
			info("Capture -> Error, code : " & @error)
		 EndIf
	  ElseIf _IsPressed("02") Then
		 info("============================")
		 info("Capture " & $name & " -> Manual cancel")
		 info("============================")
		 ExitLoop
	  EndIf
	  Sleep(250)
   WEnd

   DllClose($hDLL)

   GUICtrlSetState($GUI_EVENT_START, $GUI_ENABLE)
EndFunc

Func info($messageJournal, $autresInfos = "")
   GUICtrlSetData($journal, GUICtrlRead($journal) & @CRLF & $messageJournal)
   $end = StringLen(GUICtrlRead($journal))
   _GUICtrlEdit_SetSel($journal, $end, $end)
   _GUICtrlEdit_Scroll($journal, $SB_SCROLLCARET)
EndFunc

Func ExitScript()
   info($reason)
   Sleep(1000)

   Exit
EndFunc