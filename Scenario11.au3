;*******************************************************************
;Description: ACS
;
;Purpose:
;

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
Local $sFilePath1 =  @ScriptDir & "\" & "TestData.xlsx";This file should already exist in the mentioned path
Local $oExcel = _ExcelBookOpen($sFilePath1,0,True)

;Dim $oExcel1 = _ExcelBookNew(0)


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
Local $testCaseIteration = _ExcelReadCell($oExcel, 13, 1)
Local $testCaseExecute = _ExcelReadCell($oExcel, 13, 2)
Local $testCaseName = _ExcelReadCell($oExcel, 13, 3)
Local $testCaseDescription = _ExcelReadCell($oExcel, 13, 4)
Local $JunoOrKep  = _ExcelReadCell($oExcel, 13, 5)
Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 13, 6)
;if $JunoOrKep = "Juno" Then
  ; Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 16, 6)
;Else
   ;Local $testCaseEclipseExePath = _ExcelReadCell($oExcel, 16, 7)
   ;EndIf
Local $testCaseWorkSpacePath = _ExcelReadCell($oExcel, 13, 8)
Local $testCaseProjectName = _ExcelReadCell($oExcel, 13, 9)
Local $testCaseJspName = _ExcelReadCell($oExcel, 13, 10)
Local $testCaseJspText = _ExcelReadCell($oExcel, 13, 11)
Local $testCaseAzureProjectName = _ExcelReadCell($oExcel, 13, 12)
Local $testCaseCheckJdk = _ExcelReadCell($oExcel, 13, 13)
Local $testCaseJdkPath = _ExcelReadCell($oExcel, 13, 14)
Local $testCaseCheckLocalServer = _ExcelReadCell($oExcel, 13, 15)
Local $testCaseServerPath = _ExcelReadCell($oExcel, 13, 16)
Local $testCaseServerNo = _ExcelReadCell($oExcel, 13, 17)
Local $testCaseUrl = _ExcelReadCell($oExcel, 13, 19)
Local $emulatorURL = _ExcelReadCell($oExcel, 13, 18)
Local $testCaseValidationText = _ExcelReadCell($oExcel, 13, 19)
Local $testCaseSubscription = _ExcelReadCell($oExcel, 13, 20)
Local $testCaseStorageAccount = _ExcelReadCell($oExcel, 13, 21)
Local $testCaseServiceName = _ExcelReadCell($oExcel, 13, 22)
Local $testCaseTargetOS = _ExcelReadCell($oExcel, 13, 23)
Local $testCaseTargetEnvironment = _ExcelReadCell($oExcel, 13, 24)
Local $testCaseCheckOverwrite = _ExcelReadCell($oExcel, 13, 25)
Local $testCaseJDKOnCloud = _ExcelReadCell($oExcel, 13, 28)
Local $testCaseUserName = _ExcelReadCell($oExcel, 13, 29)
Local $testCasePassword = _ExcelReadCell($oExcel, 13, 30)
Local $testcaseNewSessionJSPText = _ExcelReadCell($oExcel, 13, 31)
Local $testcaseExternalJarPath = _ExcelReadCell($oExcel, 13, 32)
Local $testcaseCertificatePath = _ExcelReadCell($oExcel, 13, 33)
Local $testcaseACSLoginUrlPath = _ExcelReadCell($oExcel, 13, 34)
Local $testcaseACSserverUrlPath = _ExcelReadCell($oExcel, 13, 35)
Local $testcaseACSCertiPath = _ExcelReadCell($oExcel, 13, 36)
Local $testcaseCloud = _ExcelReadCell($oExcel, 13, 37)
Local $lcl = _ExcelReadCell($oExcel, 13, 38)
Local $tJDK = _ExcelReadCell($oExcel, 13, 39)
Local $PFXpath = _ExcelReadCell($oExcel, 13, 40)
Local $PFXpassword = _ExcelReadCell($oExcel, 13, 41)
Local $PSFile = _ExcelReadCell($oExcel, 13, 42)
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

;Configure
Configure()

Sleep(3000)
;Creating JSP File
CreateJSPFile($testCaseJspName, $testCaseProjectName, $testCaseJspText)

;Creating Azure Package
CreateAzurePackage($testCaseAzureProjectName, $testCaseCheckJdk, $testCaseJdkPath,$testCaseCheckLocalServer, $testCaseServerPath, $testCaseServerNo,$lcl,$tJDK)


If $testcaseCloud = 1 Then
;Publishing to Cloud
;PublishToCloud($testCaseSubscription, $testCaseStorageAccount, $testCaseServiceName, $testCaseTargetOS, $testCaseTargetEnvironment, $testCaseCheckOverwrite, $testCaseUserName, $testCasePassword)
PublishToCloud($testCaseSubscription, $testCaseStorageAccount, $testCaseServiceName, $testCaseTargetOS, $testCaseTargetEnvironment, $testCaseCheckOverwrite, $testCaseUserName, $testCasePassword,$PFXpath,$PFXpassword,$PSFile)
Sleep(20000)
Publish($testCaseProjectName,$testCaseValidationText)
Else
Emulator($emulatorURL)
EndIf





;**************************************Configuring***************************************************
Func Configure()
   ;Configuring
Send("{APPSKEY}")
Send("B")
Send("{Right}")
Send("c!a!n")
Send("{Tab 6}")
Send("^a {Delete}")
AutoItSetOption ( "SendKeyDelay", 50)
Send($testcaseACSLoginUrlPath)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{Tab}")
AutoItSetOption ( "SendKeyDelay", 50)
Send($testcaseACSserverUrlPath)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{tab 2} {Enter}")
AutoItSetOption ( "SendKeyDelay", 50)
Send($testcaseACSCertiPath)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{Enter}")
Send("{Tab 3}{Space}{Tab}{Space}{Enter}{Tab 2}{Enter}")
Sleep(2000)
Send("{Enter}")
Send("!{F4}")
;Send("{Enter}")
EndFunc

