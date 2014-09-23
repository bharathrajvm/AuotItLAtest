;*******************************************************************
;Description: Caching
;
;Purpose: Creates a Java Project and publish in cloud with staging target
;Environment and Overwrite previous deployment
;
;Date: 12 Jun 2014 , Modified on 13 June 2014
;Author: Ganesh
;Company: Brillio
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
Local $sFilePath1 =  @ScriptDir & "\" & "TestData.xlsx" ;This file should already exist in the mentioned path
Local $oExcel = _ExcelBookOpen($sFilePath1,0,True)
Dim $oExcel1 = _ExcelBookNew(0)

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
;to do - looping to get the data from desired row of xls; Reading xls data into variables

Local $testCaseIteration = _ExcelReadCell($oExcel, 11, 1)
Local $testCaseExecute = _ExcelReadCell($oExcel, 11, 2)
Local $testCaseName = _ExcelReadCell($oExcel, 11, 3)
Local $testCaseDescription = _ExcelReadCell($oExcel, 11, 4)
Local $JunoOrKep  = _ExcelReadCell($oExcel, 11, 5)
Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 11, 6)
;if $JunoOrKep = "Juno" Then
  ; Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 16, 6)
;Else
   ;Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 16, 7)
   ;EndIf
Local $testCaseWorkSpacePath = _ExcelReadCell($oExcel, 11, 8)
Local $testCaseProjectName = _ExcelReadCell($oExcel, 11, 9)
Local $testCaseJspName = _ExcelReadCell($oExcel, 11, 10)
Local $testCaseJspText = _ExcelReadCell($oExcel, 11, 11)
Local $testCaseAzureProjectName = _ExcelReadCell($oExcel, 11, 12)
Local $testCaseCheckJdk = _ExcelReadCell($oExcel, 11, 13)
Local $testCaseJdkPath = _ExcelReadCell($oExcel, 11, 14)
Local $testCaseCheckLocalServer = _ExcelReadCell($oExcel, 11, 15)
Local $testCaseServerPath = _ExcelReadCell($oExcel, 11, 16)
Local $testCaseServerNo = _ExcelReadCell($oExcel, 11, 17)
Local $testCaseUrl = _ExcelReadCell($oExcel, 11, 19)
Local $testCaseValidationText = _ExcelReadCell($oExcel, 11, 19)
Local $emulatorURL = _ExcelReadCell($oExcel, 11, 18)
Local $testCaseSubscription = _ExcelReadCell($oExcel, 11, 20)
Local $testCaseStorageAccount = _ExcelReadCell($oExcel, 11, 21)
Local $testCaseServiceName = _ExcelReadCell($oExcel, 11, 22)
Local $testCaseTargetOS = _ExcelReadCell($oExcel, 11, 23)
Local $testCaseTargetEnvironment = _ExcelReadCell($oExcel, 11, 24)
Local $testCaseCheckOverwrite = _ExcelReadCell($oExcel, 11, 25)
Local $testCaseJDKOnCloud = _ExcelReadCell($oExcel, 11, 28)
Local $testCaseUserName = _ExcelReadCell($oExcel, 11, 29)
Local $testCasePassword = _ExcelReadCell($oExcel, 11, 30)
Local $testcaseNewSessionJSPText = _ExcelReadCell($oExcel, 11, 31)
Local $testcaseExternalJarPath = _ExcelReadCell($oExcel, 11, 32)
Local $testcaseCertificatePath = _ExcelReadCell($oExcel, 11, 33)
Local $testcaseACSLoginUrlPath = _ExcelReadCell($oExcel, 11, 34)
Local $testcaseACSserverUrlPath = _ExcelReadCell($oExcel, 11, 35)
Local $testcaseACSCertiPath = _ExcelReadCell($oExcel, 11, 36)
Local $testcaseCloud = _ExcelReadCell($oExcel, 11, 37)
Local $lcl = _ExcelReadCell($oExcel, 11, 38)
Local $tJDK = _ExcelReadCell($oExcel, 11, 39)
Local $PFXpath = _ExcelReadCell($oExcel, 11, 40)
Local $PFXpassword = _ExcelReadCell($oExcel, 11, 41)
Local $PSFile = _ExcelReadCell($oExcel, 11, 42)
_ExcelBookClose($oExcel,0)
Local $exlid = ProcessExists("excel.exe")
ProcessClose($exlid)
;*******************************************************************************




Start($testCaseName,$testCaseDescription);Calling Start function from Testinc.au3 script(Custom Script file, Contains common functions that are used in other scripts)

Local $pro = ProcessExists("eclipse.exe")
If $pro > 0 Then
  Delete()
Else
 OpenEclipse($testCaseEclipseExePath,$testCaseWorkSpacePath)
 Delete()
EndIf

;Creating Java Project
CreateJavaProject($testCaseProjectName)

;Creating JSP file and insert code
CreateJSPFile1()

;Adding External JAR FileChangeDir
AddExternalJarFile()

;Create Azure Package
CreateAzurePackage($testCaseAzureProjectName, $testCaseCheckJdk, $testCaseJdkPath,$testCaseCheckLocalServer, $testCaseServerPath, $testCaseServerNo,$lcl,$tJDK)

Sleep(8000)
;Enable co-located caching
EnableCoLocatedCaching()

 WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)

Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")
 Send("^+{NUMPADDIV}")

If $testcaseCloud = 1 Then
;Publish to Cloud
 ;PublishToCloud($testCaseSubscription, $testCaseStorageAccount, $testCaseServiceName, $testCaseTargetOS, $testCaseTargetEnvironment, $testCaseCheckOverwrite, $testCaseUserName, $testCasePassword)
 PublishToCloud($testCaseSubscription, $testCaseStorageAccount, $testCaseServiceName, $testCaseTargetOS, $testCaseTargetEnvironment, $testCaseCheckOverwrite, $testCaseUserName, $testCasePassword,$PFXpath,$PFXpassword,$PSFile)
Sleep(30000)

Publish($testCaseProjectName,$testCaseValidationText)
Else
   Emulator($emulatorURL)
EndIf


#cs
;*****************************************************************
;Function to publish to cloud
;****************************************************************
Func PublishToCloud1()
Sleep(2000)
WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)
 Local $win6 = WinActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 If $win6 = 0 Then
	$cls = "------Error in Publishing to Cloud (Cannot Open: Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse)--------"
	Close($cls)
	Exit
 EndIf
 Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")

Send("^+{NUMPADDIV}")
;Send("{Up}")
Send("{APPSKEY}")
Sleep(1000)

Send("e")
Send("{Left}")
Send("{UP}")
;Send("{Down 21}")
Send("{Right}")
Send("{Enter}")

WinWaitActive("Publish Wizard")
Sleep(3000)
 Local $win7 = WinActive("Publish Wizard")
 If $win7 = 0 Then
	$cls = "------(Cannot Open: Publish Wizard )--------"
	Close($cls)
	Exit
 EndIf
while 1
Dim $hnd =  WinGetText("Publish Wizard","")
StringRegExp($hnd,"Loading Account Settings...",1)
Local $reg = @error
if $reg > 0 Then ExitLoop
WEnd

WinActive("Publish Wizard")
Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:1]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseSubscription)

 Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:2]")
 ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseStorageAccount)


Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:3]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseServiceName)


Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:4]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseTargetOS)

Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:5]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseTargetEnvironment)


Local $cmp = StringCompare($testCaseCheckOverwrite,"UnCheck")
   if $cmp = 0 Then
	   ControlCommand("Publish Wizard","","[CLASSNN:Button4]","Check", "")
	   sleep(3000)
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","UnCheck", "")
   Else
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","UnCheck", "")
	   sleep(3000)
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","Check", "")
   EndIf

Send("{TAB}")
AutoItSetOption ( "SendKeyDelay", 100)
Send($testCaseUserName)
Send("{TAB}")
Send($testCasePassword)
Send("{TAB 2}")
Send($testCasePassword)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB}")
ControlCommand("Publish Wizard","","[CLASSNN:Button5]","Check", "")
Send("{TAB}")
Send("{Enter}")

EndFunc
;*******************************************************************************
#ce

;***************************************************************
;Helper Functions
;***************************************************************

Func CreateJSPFile1()
sleep(3000)
Send("{APPSKEY}")
AutoItSetOption ( "SendKeyDelay", 100)
Send("{down}")
Send("{RIGHT}")
Send("{down 14}")
Send("{enter}")
Send($testCaseJspName)
;Send("{TAB 3}")
;Send("{Enter}")
Send("!f")
Local $temp = "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
Sleep(3000)
WinWaitActive($temp)
Sleep(2000)
		 Local $win4 = WinActive($temp)
		 If $win4 = 0 Then
			$cls = "---Error in Opening: "& $temp &"--------"
			Send("{Esc}")
			Close($cls)
			Exit
		 EndIf

; Calling the Winchek Function
Local $funame, $cntrlname
$cntrlname =  "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
$funame = "CreateJSPFile"
wincheck($funame,$cntrlname)
AutoItSetOption ( "SendKeyDelay", 100)
Send("^a")
Send("{Backspace}")
ClipPut($testCaseJspText)
Send("^v")
AutoItSetOption ( "SendKeyDelay", 400)
Send("^+s")
EndFunc
;******************************************************************

;***************************************************************
;Function to add external JAR file
;***************************************************************
Func AddExternalJarFile()
AutoItSetOption ( "SendKeyDelay", 200)
WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)
;MouseClick("primary",105, 395, 1)
Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")
Send("{APPSKEY}")
Sleep(1000)
Send("b")
Send("{Right}")
Send("c")
WinWaitActive("Properties for MyHelloWorld")

Send("{tab 3}{l}")

Send("!x")
WinWaitActive("JAR Selection")
ClipPut($testcaseExternalJarPath)
Send("^v")
Send("!o")

for $shiftTab = 5 To 1 step -1
Send("+")
Send("{TAB}")
Next

Send("{UP}")
WinWaitActive("Setting Java Build Path")
Send("{Enter}")
WinWaitActive("Properties for MyHelloWorld")
Send("!d")
WinWaitActive("New Assembly Directive")
Send("{down 4}")
Send("!n")
WinWaitActive("New Assembly Directive")
Send("{TAB 3}")
Send("{Space}")
Send("!f")
Send("!a")
Send("{TAB}{ENTER}")
EndFunc



;***************************************************************
;Function to} enable co-located caching
;***************************************************************
Func EnableCoLocatedCaching()
Sleep(2000)
Send("{UP}{ENTER}")
Sleep(2000)
Send("{DOWN 3}")
Sleep(2000)
Send("{APPSKEY}")
Sleep(1000)

#comments-start
if $JunoOrKep = "Juno" Then
Send("g")
Else
Send("e")
EndIf
#comments-end
Send("e")
if $JunoOrKep = "Juno" Then
   Send("{Left}{UP}{Enter}")
   Send("{Down}{Enter}")
   WinWaitActive("[Title:Properties for WorkerRole1]")
ControlCommand("Properties for WorkerRole1","","[CLASSNN:Button1]","Check", "")
;Send("{TAB}{a}{Right}{Down}{tab 10}")
Send("{TAB 8}")
Send("{Enter}")

Else
   Send("{Left}{UP}{right}{down}{Enter}")
   WinWaitActive("[Title:Properties for WorkerRole1]")
ControlCommand("Properties for WorkerRole1","","[CLASSNN:Button1]","Check", "")
Send("{TAB 8}")
Send("{Enter}")

   EndIf

;WinWaitActive("[Title:Properties for WorkerRole1]")
;ControlCommand("Properties for WorkerRole1","","[CLASSNN:Button1]","Check", "")
;Send("{TAB 8}")
;Send("{Enter}")
EndFunc





