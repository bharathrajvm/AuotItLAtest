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
;to do - looping to get the data from desired row of xls
Local $testCaseIteration = _ExcelReadCell($oExcel, 18, 1)
Local $testCaseExecute = _ExcelReadCell($oExcel, 18, 2)
Local $testCaseName = _ExcelReadCell($oExcel, 18, 3)
Local $testCaseDescription = _ExcelReadCell($oExcel, 18, 4)
Local $JunoOrKep  = _ExcelReadCell($oExcel, 18, 5)
Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 18, 6)
;if $JunoOrKep = "Juno" Then
  ; Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 16, 6)
;Else
   ;Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 16, 7)
   ;EndIf
Local $testCaseWorkSpacePath = _ExcelReadCell($oExcel, 18, 8)
Local $testCaseProjectName = _ExcelReadCell($oExcel, 18, 9)
Local $testCaseJspName = _ExcelReadCell($oExcel, 18, 10)
Local $testCaseJspText = _ExcelReadCell($oExcel, 18, 11)
Local $testCaseAzureProjectName = _ExcelReadCell($oExcel, 18, 12)
Local $testCaseCheckJdk = _ExcelReadCell($oExcel, 18, 13)
Local $testCaseJdkPath = _ExcelReadCell($oExcel, 18, 14)
Local $testCaseCheckLocalServer = _ExcelReadCell($oExcel, 18, 15)
Local $testCaseServerPath = _ExcelReadCell($oExcel, 18, 16)
Local $testCaseServerNo = _ExcelReadCell($oExcel, 18, 17)
Local $testCaseUrl = _ExcelReadCell($oExcel, 18, 19)
Local $testCaseValidationText = _ExcelReadCell($oExcel, 18, 19)
Local $testCaseSubscription = _ExcelReadCell($oExcel, 18, 20)
Local $testCaseStorageAccount = _ExcelReadCell($oExcel, 18, 21)
Local $testCaseServiceName = _ExcelReadCell($oExcel, 18, 22)
Local $testCaseTargetOS = _ExcelReadCell($oExcel, 18, 23)
Local $testCaseTargetEnvironment = _ExcelReadCell($oExcel, 18, 24)
Local $testCaseCheckOverwrite = _ExcelReadCell($oExcel, 18, 25)
Local $testCaseJDKOnCloud = _ExcelReadCell($oExcel, 18, 28)
Local $testCaseUserName = _ExcelReadCell($oExcel, 18, 29)
Local $testCasePassword = _ExcelReadCell($oExcel, 18, 30)
Local $testcaseNewSessionJSPText = _ExcelReadCell($oExcel, 18, 31)
Local $testcaseExternalJarPath = _ExcelReadCell($oExcel, 18, 32)
Local $testcaseCertificatePath = _ExcelReadCell($oExcel, 18, 33)
Local $testcaseACSLoginUrlPath = _ExcelReadCell($oExcel, 18, 34)
Local $testcaseACSserverUrlPath = _ExcelReadCell($oExcel, 18, 35)
Local $testcaseACSCertiPath = _ExcelReadCell($oExcel, 18, 36)
Local $testcaseCloud = _ExcelReadCell($oExcel, 18, 37)
Local $lcl = _ExcelReadCell($oExcel, 18, 38)
Local $tJDK = _ExcelReadCell($oExcel, 18, 39)
Local $PFXpath = _ExcelReadCell($oExcel, 18, 40)
Local $PFXpassword = _ExcelReadCell($oExcel, 18, 41)
Local $PSFile = _ExcelReadCell($oExcel, 18, 42)
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
CreateJSPFile($testCaseJspName, $testCaseProjectName, $testCaseJspText)

;Adding External JAR FileChangeDir
AddExternalJarFile()

;Create Azure Package
CreateAzurePackage($testCaseAzureProjectName, $testCaseCheckJdk, $testCaseJdkPath,$testCaseCheckLocalServer, $testCaseServerPath, $testCaseServerNo,$lcl,$tJDK)
Sleep(8000)

;Enable co-located caching
EnableCoLocatedCaching()


;Enable SSL Offloading
EnableSSLOffloading()

;Publish to Cloud
;PublishToCloud($testCaseSubscription, $testCaseStorageAccount, $testCaseServiceName, $testCaseTargetOS, $testCaseTargetEnvironment, $testCaseCheckOverwrite, $testCaseUserName, $testCasePassword)
PublishToCloud1()
Sleep(20000)


For $i = 8 to 1 Step - 1
   Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
   Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:msctls_progress32]")
   Local $syslk = ControlCommand($wnd, "", $wnd1,"IsVisible", "")
If $i = 1 and $syslk = 0 Then
   $cls = "-----Time Out!-----"
   Close($cls)
   Exit
Else
      ;Send("{Enter}")
		 If $syslk = 0 Then
			;Check RDP and Open excel
			CheckRDPConnection()
			Sleep(10000)
			;Check for published key word in Azure activity log and update excel
			ValidateTextAndUpdateExcel($testCaseProjectName, $testCaseValidationText)
			sleep(7000)
			$cls = 1
			Close($cls)
   ;Exit
		 Else
			Sleep(120000)
		 EndIf
EndIf
Next

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

Sleep(18000)
Local $act = WinActive("Upload certificate")
Local $wnd = WinGetHandle("Upload certificate")
If $wnd > 0 Then    ;Checking the Upload certificate window to upload the PFX file (this fuction is for SSL offloading scenarios)
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:Edit; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")
 AutoItSetOption ( "SendKeyDelay", 100)
Send($PFXpath)
Send("{Tab 2}")
Send($PFXpassword)
Send("{tab}")
Send("{Enter}")
EndIf


EndFunc
;*******************************************************************************


;***************************************************************
;Function to add external JAR file
;***************************************************************
Func AddExternalJarFile()
AutoItSetOption ( "SendKeyDelay", 200)
WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)
Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")
;MouseClick("primary",105, 395, 1)
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
Sleep(1500)
Send("{UP}{ENTER}")
Sleep(1500)
Send("{DOWN 3}")
Sleep(1500)
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
   Sleep(2000)
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



;****************************************************************
;Function to enable SSL Offloading
;***************************************************************
Func EnableSSLOffloading()
Sleep(2000)
AutoItSetOption ( "SendKeyDelay", 100)
Send("{APPSKEY}")
;Send("{Up}{Enter}{down}{down}{down}{APPSKEY}")
Sleep(1000)
;if $JunoOrKep = "Juno" Then
;Send("g")
;Else
;Send("e")
;EndIf
Send("e")
if $JunoOrKep = "Juno" Then
Send("{Left}{UP}{Right}{s 2}{Enter}")
Else
Send("{s}{Left}{UP}{UP}{Right}{s 2}{Enter}")
EndIf
WinWaitActive("[Title:Properties for WorkerRole1]")
ControlCommand("Properties for WorkerRole1","","[CLASSNN:Button1]","Check", "")
WinWaitActive("[Title:SSL Offloading]")
Send("{Enter}")
Send("{Tab 4}")
;Dim $hWnd = WinGetHandle("[Title:Properties for WorkerRole1]")
;MsgBox("","",$hWnd)
;Local $hToolBar = ControlGetHandle($hWnd, "", "[Text:<a>Certificates...</a>]")
;MsgBox("","",$hToolBar)
;ControlClick($hWnd,"",6424586,"left",1)
 ;MouseClick("primary",1039, 133, 1)

Local $wnd = WinGetHandle("[Title:Properties for WorkerRole1]")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysLink; INSTANCE:2]")
 ControlClick($wnd,"",$wnd1,"left")
WinWaitActive("[Title:Properties for MyAzureProject]")
Send("{Tab 2}")
Send("{Enter}")
WinWaitActive("[Title:Certificate]")
Send("{Tab 2}")
Send("{Enter}")
Sleep(2000)
Send($testcaseCertificatePath)
;ClipPut($testcaseCertificatePath)
Send("!O")
Send("{Tab 2}")
Send("{Enter}")
Send("{Tab}{Enter}")
Send("{Tab 3}")
Send("{Enter}")
EndFunc
;*********************************************************************

