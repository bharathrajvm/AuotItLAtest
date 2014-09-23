;*******************************************************************
;Description: Session Affinity
;
;Purpose: Creates a Java Project and publish in cloud with staging target
;Environment and Overwrite previous deployment
;
;Date: 10 Jun 2014 , Modified on 12 June 2014
;Author: Ganesh
;Company: Brillio
;*********************************************************************

;********************************************
;Include Standard Library
;*******************************************
#include <Constants.au3>
#include <GuiToolbar.au3>
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
#include <Testinc.au3>
;******************************************

;***************************************************************
;Initialize AutoIT Key delay
;****************************************************************
AutoItSetOption ( "SendKeyDelay", 400)

;******************************************************************
;Reading test data from xls
;To do - move helper function
;******************************************************************
;Open xls
Local $sFilePath1 = @ScriptDir & "\" & "TestData.xlsx" ;This file should already exist in the mentioned path
Local $oExcel = _ExcelBookOpen($sFilePath1,0,True)
;Dim $oExcel1 = _ExcelBookNew(0)

;Local $sFilePath2 = @ScriptDir & "\" & "Result.xlsx"  ;This file should already exist in the mentioned path
;Local $oExcel1 = _ExcelBookOpen($sFilePath2,0,False)

;Check for error
If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist")
    Exit
 EndIf

 ; Reading xls data into variables
;to do - looping to get the data from desired row of xls
Local $testCaseIteration = _ExcelReadCell($oExcel, 10, 1)
Local $testCaseExecute = _ExcelReadCell($oExcel, 10, 2)
Local $testCaseName = _ExcelReadCell($oExcel, 10, 3)
Local $testCaseDescription = _ExcelReadCell($oExcel, 10, 4)
Local $JunoOrKep  = _ExcelReadCell($oExcel, 10, 5)
Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 10, 6)
;if $JunoOrKep = "Juno" Then
  ; Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 14, 6)
;Else
   ;Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 14, 7)
   ;EndIf
Local $testCaseWorkSpacePath = _ExcelReadCell($oExcel, 10, 8)
Local $testCaseProjectName = _ExcelReadCell($oExcel, 10, 9)
Local $testCaseJspName = _ExcelReadCell($oExcel, 10, 10)
Local $testCaseJspText = _ExcelReadCell($oExcel, 10, 11)
Local $testCaseAzureProjectName = _ExcelReadCell($oExcel, 10, 12)
Local $testCaseCheckJdk = _ExcelReadCell($oExcel, 10, 13)
Local $testCaseJdkPath = _ExcelReadCell($oExcel, 10, 14)
Local $testCaseCheckLocalServer = _ExcelReadCell($oExcel, 10, 15)
Local $testCaseServerPath = _ExcelReadCell($oExcel, 10, 16)
Local $testCaseServerNo = _ExcelReadCell($oExcel, 10, 17)
Local $emulatorURL = _ExcelReadCell($oExcel, 10, 18)
Local $testCaseUrl = _ExcelReadCell($oExcel, 10, 19)
Local $testCaseValidationText = _ExcelReadCell($oExcel, 10, 19)
Local $testCaseSubscription = _ExcelReadCell($oExcel, 10, 20)
Local $testCaseStorageAccount = _ExcelReadCell($oExcel, 10, 21)
Local $testCaseServiceName = _ExcelReadCell($oExcel, 10, 22)
Local $testCaseTargetOS = _ExcelReadCell($oExcel, 10, 23)
Local $testCaseTargetEnvironment = _ExcelReadCell($oExcel, 10, 24)
Local $testCaseCheckOverwrite = _ExcelReadCell($oExcel, 10, 25)
Local $testCaseJDKOnCloud = _ExcelReadCell($oExcel, 10, 28)
Local $testCaseUserName = _ExcelReadCell($oExcel, 10, 29)
Local $testCasePassword = _ExcelReadCell($oExcel, 10, 30)
Local $testcaseNewSessionJSPText = _ExcelReadCell($oExcel, 10, 31)
Local $testcaseCloud = _ExcelReadCell($oExcel, 10, 37)
Local $lcl = _ExcelReadCell($oExcel, 10, 38)
Local $tJDK = _ExcelReadCell($oExcel, 10, 39)
Local $PFXpath = _ExcelReadCell($oExcel, 10, 40)
Local $PFXpassword = _ExcelReadCell($oExcel, 10, 41)
Local $PSFile = _ExcelReadCell($oExcel, 10, 42)
_ExcelBookClose($oExcel,0)
Local $exlid = ProcessExists("excel.exe")
ProcessClose($exlid)
;*******************************************************************************
;*******************************************************************************

Start($testCaseName,$testCaseDescription);Calling Start function from Testinc.au3 script(Custom Script file, Contains common functions that are used in other scripts)

Local $pro = ProcessExists("eclipse.exe")
If $pro > 0 Then
  Delete()
Else
 OpenEclipse($testCaseEclipseExePath,$testCaseWorkSpacePath)
 Delete()
EndIf

;Create javaProject
CreateJavaProject($testCaseProjectName)

;create JSP file
CreateSessionJSPFile()

WinActivate("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
  Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
   Local $htoolBar =  ControlGetHandle($wnd, '','[CLASS:SWT_Window0; INSTANCE:14]')
   Sleep(3000)
   _GUICtrlToolbar_ClickIndex($htoolBar, 0,"Left")

;CreateAzurePackage
CreateAzurePackage($testCaseAzureProjectName, $testCaseCheckJdk, $testCaseJdkPath,$testCaseCheckLocalServer, $testCaseServerPath, $testCaseServerNo,$lcl,$tJDK)

Sleep(3000)
Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")
 Send("^+{NUMPADDIV}")
 Send("{UP}")
 Send("{Enter}")
 Send("{Down 3}")
 Send("{APPSKEY}")
 Sleep(1500)
 Send("e")
 Send("{Left}")
 Send("{UP}")
 Send("{Right}")
 Send("L")
 Send("{Enter}")
 Send("{Tab}")
Send("{Space}")
 Send("{Tab 4} {Enter}")
Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")
 Send("^+{NUMPADDIV}")


If  $testcaseCloud = 1  Then
;Publish to Cloud
;PublishToCloud($testCaseSubscription, $testCaseStorageAccount, $testCaseServiceName, $testCaseTargetOS, $testCaseTargetEnvironment, $testCaseCheckOverwrite, $testCaseUserName, $testCasePassword)
PublishToCloud($testCaseSubscription, $testCaseStorageAccount, $testCaseServiceName, $testCaseTargetOS, $testCaseTargetEnvironment, $testCaseCheckOverwrite, $testCaseUserName, $testCasePassword,$PFXpath,$PFXpassword,$PSFile)
;Wait for 30 seconds
Sleep(30000)
Publish($testCaseProjectName,$testCaseValidationText)
Else
Emulator($emulatorURL)
EndIf


;***************************************************************
;Function to create JSP file and insert code
;***************************************************************
Func CreateSessionJSPFile()
sleep(3000)
Send("{APPSKEY}")
AutoItSetOption ( "SendKeyDelay", 100)
Send("{down}")
Send("{RIGHT}")
Send("{down 14}")
Send("{enter}")
#cs
;create newsession.jsp
MouseClick("primary",105, 395, 1)
Send("{APPSKEY}")
Sleep(1000)
Send("n")
Send("{down 14}")
Send("{enter}")
#ce
Send("newsession.jsp")
Send("!f")
Local $temp = "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
Sleep(2000)
Send("^a")
Send("{Backspace}")
ClipPut($testcaseNewSessionJSPText)
Send("^v")
AutoItSetOption ( "SendKeyDelay", 400)
Send("^+s")
;MouseClick("primary", 74, 114, 1)

 WinWaitActive("Java EE - MyHelloWorld/WebContent/newsession.jsp - Eclipse")
Sleep(3000)

Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/newsession.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")

 Send("{APPSKEY}")
Sleep(1000)
Send("n")
Send("{down 14}")
Send("{enter}")
Send($testCaseJspName)
Send("!f")
Local $temp = "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
Sleep(3000)
WinWaitActive($temp)

; Calling the Winchek Function
Local $funame, $cntrlname
$cntrlname =  "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
$funame = "CreateJSPFile"
wincheck($funame,$cntrlname)

Send("^a")
Send("{Backspace}")
ClipPut($testCaseJspText)
Send("^v")
Send("^+s")
 WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)

EndFunc
;******************************************************************


