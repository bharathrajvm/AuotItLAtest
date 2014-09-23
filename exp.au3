

;*********************************************************************

;********************************************
;Include Standard Library
;*******************************************
#include <Constants.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <IE.au3>
#include <Clipboard.au3>
#include <Date.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>
#include <GuiTreeView.au3>
#include <GuiImageList.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <GuiTreeView.au3>
#include <File.au3>
#include <GuiToolbar.au3>
#include <Testinc.au3>


Global $cnt = 0
;WinActivate("Microsoft Azure Compute Emulator (Express)")
Local $wnd = WinActive("Microsoft Azure Compute Emulator (Express)")
;MsgBox("","",@error)
;MsgBox("","",$wnd)
  While $wnd = 0
	 ;MsgBox("","",$cnt)
        WinActivate("Microsoft Azure Compute Emulator (Express)")
	  Local $wnd = WinGetHandle("Microsoft Azure Compute Emulator (Express)")
        Sleep(2000)
		$cnt = $cnt + 1
	  if $cnt > 3 Then
		 ;SetError(0)
		 MsgBox("","","Exceeded")
		 Exit
	  EndIf
 WEnd










