#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <TreeViewConstants.au3>
#include <GuiComboBoxEx.au3>
#include <Excel.au3>
#include <Array.au3>
#include <GuiButton.au3>
#include <MsgBoxConstants.au3>
 #include <GuiListBox.au3>
 #include <Testinc.au3>

;**********************************************************
;Open xls
;**********************************************************
Dim $sFilePath1 = @ScriptDir & "\TestData.xlsx" ;This file should already exist in the mentioned path
Dim $oExcel = _ExcelBookOpen($sFilePath1,0,False)

;Check for error
If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist")
    Exit
EndIf

Dim $a = _ExcelReadSheetToArray($oExcel,1,1)
Dim $arrayRowCount = UBound($a, 1)

Dim $arrayColumnCount = UBound($a, 2)


;*************************************************************************
;GUI layout
;**************************************************************************
Dim $hMainGUI = GUICreate("Example",1000,900,"","",$WS_CAPTION,$WS_EX_APPWINDOW)
Dim $hCancelButton = GUICtrlCreateButton("Cancel", 610, 650, 85, 25)
Dim $hSaveButton = GUICtrlCreateButton("Save", 510, 650, 85, 25)
Dim $hUpdateButton = GUICtrlCreateButton("Update",410,650, 85, 25)

GUICtrlCreateLabel("Select Testcase",350,3)
GUICtrlCreateLabel("TargetEnvironmentUnpublish", 350, 40)
GUICtrlCreateLabel("JDKOnCloud", 350, 80)
GUICtrlCreateLabel("UserName", 350, 120)
GUICtrlCreateLabel("Password", 350, 160)
GUICtrlCreateLabel("NewsessionJSPText", 350, 200)
GUICtrlCreateLabel("ExternalJARPath", 350, 240)
GUICtrlCreateLabel("CertificatePath", 350, 280)
GUICtrlCreateLabel("EclipsePath", 350, 320)
GUICtrlCreateLabel("Scenarios", 350, 360)


GUICtrlCreateLabel("Eclipse", 3, 3)
GUICtrlCreateLabel("EclipseWorkspace", 3, 40)
GUICtrlCreateLabel("WebProjectName", 3, 80)
GUICtrlCreateLabel("JSP FileName", 3, 120)
GUICtrlCreateLabel("JSP Text", 3, 160)
GUICtrlCreateLabel("Azure Project Name", 3, 200)
GUICtrlCreateLabel("JDK Path", 3, 240)
GUICtrlCreateLabel("Server Check", 3, 280)
GUICtrlCreateLabel("Server Path", 3, 320)
GUICtrlCreateLabel("Server Type", 3, 360)
GUICtrlCreateLabel("Validation Text", 3, 400)
GUICtrlCreateLabel("Subscription", 3, 440)
GUICtrlCreateLabel("Storage Account", 3, 480)
GUICtrlCreateLabel("Service Name", 3, 520)
GUICtrlCreateLabel("Target OS", 3, 560)
GUICtrlCreateLabel("Target Envi", 3, 600)
GUICtrlCreateLabel("Overwrite", 3, 640)
GUICtrlCreateLabel("ServiceNameUnpublish", 3, 680)

Dim $hTestcase =  GUICtrlCreateCombo("",500,3,200,20)
Dim $TargetEnvironmentUnpublish = GUICtrlCreateInput("",500,40,200,20)
Dim $JDKOnCloud = GUICtrlCreateInput("",500,80,200,20)
Dim $UserName = GUICtrlCreateInput("",500,120,200,20)
Dim $Password = GUICtrlCreateInput("",500,160,200,20)
Dim $newsession = GUICtrlCreateInput("",500,200,200,20)
Dim $ExternalJARPath = GUICtrlCreateInput("",500,240,200,20)
Dim $CertificatePath = GUICtrlCreateInput("",500,280,200,20)
Dim $EclipsePath = GUICtrlCreateInput("",500,320,200,20)
Dim $list = _GUICtrlListBox_Create($hMainGUI,"",550,360,230,250, $LBS_EXTENDEDSEL)


Dim $hEclipseType =  GUICtrlCreateInput("",130,3,200,20)
Dim $hWorkspace =	GUICtrlCreateInput("", 130,40,200,20)
Dim $hWebProjName =	GUICtrlCreateInput("", 130,80,200,20)
Dim $hJSPFileName =	GUICtrlCreateInput("", 130,120,200,20)
Dim $hJSPText =	GUICtrlCreateInput("", 130,160,200,20)
Dim $hAzureProjName =	GUICtrlCreateInput("", 130,200,200,20)
Dim $hJDKPath = GUICtrlCreateInput("", 130,240,200,20)
Dim $hServerCheck = GUICtrlCreateCheckbox("",130,280,20,20)
Dim $hServerPath = GUICtrlCreateInput("", 130,320,200,20)
Dim $hServerType = GUICtrlCreateInput("", 130,360,200,20)
Dim $hValidationText = GUICtrlCreateInput("", 130,400,200,20)
Dim $hSubscription = GUICtrlCreateInput("", 130,440,200,20)
Dim $hStorageAccount = GUICtrlCreateInput("", 130,480,200,20)
Dim $hServiceName= GUICtrlCreateInput("", 130,520,200,20)
Dim $hTargetOS = GUICtrlCreateInput("", 130,560,200,20)
Dim $hTargetEnvi = GUICtrlCreateInput("", 130,600,200,20)
Dim $hOverWriteCheck = GUICtrlCreateCheckbox("",130,635,200,20)
Dim $hServiceNameUnpublish = GUICtrlCreateInput("",130,680,200,20)
Dim $hExecutionButton = GUICtrlCreateButton("Execute",410,550,125,45)

;******************************************************************************************
Local $aItems,$sItems
Local $scenarioIndex = 3
Local $loopVariable = 0
#comments-start
For $loopVariable = 2 to $arrayRowCount-1 step 1
Local $temp = $a[$loopVariable][3]
GUICtrlSetData($hTestcase, $temp)
Next
#comments-end
For $loopVariable = 2 to $arrayRowCount-1 step 1
Local $temp = $a[$loopVariable][3] & " - " & $a[$loopVariable][4]
_GUICtrlListBox_AddString($list, $temp)
Next



GUISetState(@SW_SHOW)
while 1
$msg = GUIGetMsg()

;***********************************************************************
;For update button action
;***********************************************************************
if $msg = $hUpdateButton Then
Dim $autoUpdateVariable = GUICtrlRead($hTestcase)



Dim $autoUpdateCount = 0
For $loopVariable = 1 to $arrayRowCount-1 step 1
   if $autoUpdateVariable = $a[$loopVariable][3] then
	  $autoUpdateCount = $loopVariable
	  ExitLoop
   EndIf
Next
GUICtrlSetData($hEclipseType, $a[$autoUpdateCount][5])
GUICtrlSetData($EclipsePath, $a[$autoUpdateCount][6])
GUICtrlSetData($hWorkspace, $a[$autoUpdateCount][8])
GUICtrlSetData($hWebProjName, $a[$autoUpdateCount][9])
GUICtrlSetData($hJSPFileName, $a[$autoUpdateCount][10])
GUICtrlSetData($hJSPText, $a[$autoUpdateCount][11])
GUICtrlSetData($hAzureProjName, $a[$autoUpdateCount][12])
if $a[$autoUpdateCount][13] = "Check" Then
   _GUICtrlButton_SetCheck($hServerCheck, $BST_CHECKED)
   Else
_GUICtrlButton_SetCheck($hServerCheck, $BST_UNCHECKED )
   EndIf
GUICtrlSetData($hJDKPath, $a[$autoUpdateCount][14])
GUICtrlSetData($hServerPath, $a[$autoUpdateCount][16])
GUICtrlSetData($hServerType, $a[$autoUpdateCount][17])
GUICtrlSetData($hValidationText, $a[$autoUpdateCount][19])
GUICtrlSetData($hSubscription, $a[$autoUpdateCount][20])
GUICtrlSetData($hStorageAccount, $a[$autoUpdateCount][21])
GUICtrlSetData($hServiceName, $a[$autoUpdateCount][22])
GUICtrlSetData($hTargetOS, $a[$autoUpdateCount][23])
GUICtrlSetData($hTargetEnvi, $a[$autoUpdateCount][24])
GUICtrlSetData($hServiceNameUnpublish, $a[$autoUpdateCount] [26])
GUICtrlSetData($TargetEnvironmentUnpublish, $a[$autoUpdateCount] [27])
GUICtrlSetData($JDKOnCloud, $a[$autoUpdateCount] [28])
GUICtrlSetData($UserName, $a[$autoUpdateCount] [29])
GUICtrlSetData($Password, $a[$autoUpdateCount] [30])
GUICtrlSetData($newsession, $a[$autoUpdateCount] [31])
GUICtrlSetData($ExternalJARPath, $a[$autoUpdateCount] [32])
GUICtrlSetData($CertificatePath, $a[$autoUpdateCount] [33])


if $a[$autoUpdateCount][25] = "Check" Then
   _GUICtrlButton_SetCheck($hOverWriteCheck, $BST_CHECKED)
   Else
_GUICtrlButton_SetCheck($hOverWriteCheck, $BST_UNCHECKED )
   EndIf
;******************************************************************************

;***********************************************************************
;For Save button action
;***********************************************************************
ElseIf $msg = $hSaveButton then

$autoUpdateVariable = GUICtrlRead($hTestcase)

For $loopVariable = 1 to $arrayRowCount-1 step 1
   if $autoUpdateVariable = $a[$loopVariable][3] then
	  $autoUpdateCount = $loopVariable
	  ExitLoop
   EndIf
Next

$a[$autoUpdateCount][5] = GUICtrlRead($hEclipseType)
$a[$autoUpdateCount][6] = GUICtrlRead($EclipsePath)
$a[$autoUpdateCount][8] = GUICtrlRead($hWorkspace)
$a[$autoUpdateCount][9] = GUICtrlRead($hWebProjName)
$a[$autoUpdateCount][10] = GUICtrlRead($hJSPFileName)
$a[$autoUpdateCount][11] = GUICtrlRead($hJSPText)
$a[$autoUpdateCount][12] = GUICtrlRead($hAzureProjName)

if GUICtrlRead($hServerCheck) = $GUI_CHECKED Then
   $a[$autoUpdateCount][13] = "Check"
   Else
$a[$autoUpdateCount][13] = "UnCheck"
   EndIf
$a[$autoUpdateCount][14] = GUICtrlRead($hJDKPath)
$a[$autoUpdateCount][16] = GUICtrlRead($hServerPath)
$a[$autoUpdateCount][17] = GUICtrlRead($hServerType)
$a[$autoUpdateCount][19] = GUICtrlRead($hValidationText)
$a[$autoUpdateCount][20] = GUICtrlRead($hSubscription)
$a[$autoUpdateCount][21] = GUICtrlRead($hStorageAccount)
$a[$autoUpdateCount][22] = GUICtrlRead($hServiceName)
$a[$autoUpdateCount][23] = GUICtrlRead($hTargetOS)
$a[$autoUpdateCount][24] = GUICtrlRead($hTargetEnvi)
$a[$autoUpdateCount][26] = GUICtrlRead($hServiceNameUnpublish)
$a[$autoUpdateCount][27] = GUICtrlRead($TargetEnvironmentUnpublish)
$a[$autoUpdateCount][28] = GUICtrlRead($JDKOnCloud)
$a[$autoUpdateCount][29] = GUICtrlRead($UserName)
$a[$autoUpdateCount][30] = GUICtrlRead($Password)
$a[$autoUpdateCount][31] = GUICtrlRead($newsession)
$a[$autoUpdateCount][32] = GUICtrlRead($ExternalJARPath)
$a[$autoUpdateCount][33] = GUICtrlRead($CertificatePath)


if GUICtrlRead($hOverWriteCheck) = $GUI_CHECKED Then
   $a[$autoUpdateCount][25] = "Check"
   Else
$a[$autoUpdateCount][25] = "UnCheck"
   EndIf


;_ExcelWriteArray($oExcel,0,0,$a)

 _ExcelWriteSheetFromArray($oExcel,$a)


If @error Then MsgBox($MB_SYSTEMMODAL, "Not Saved", "Problem while writing into array", 3)
$flag = _ExcelBookSave($oExcel,0)
If @error Then MsgBox($MB_SYSTEMMODAL, "Not Saved", "Problem while saving", 3)
;****************************************************************************

;***********************************************************************
;For Cancel button action
;***********************************************************************
ElseIf $msg = $hCancelButton then
   $flag = _ExcelBookSave($oExcel,0)
   If @error Then MsgBox($MB_SYSTEMMODAL, "Not Saved", "Problem while saving", 3)
   _ExcelBookClose($oExcel, 1, 0)
    GUIDelete($hMainGUI)
   ExitLoop
ElseIf $msg = $hExecutionButton Then

   ;Dim $name = GUICtrlRead($hTestcase)D:\KWS
  ;Dim $AutoItExe =  @ScriptDir & '\'& $name &'.exe'
  ;Run($AutoItExe)

   $aItems = _GUICtrlListBox_GetSelItemsText($list)
   For $iI = 1 To UBound($aItems) - 1
   $sItems = $aItems[$iI]
   Local $spl = StringSplit($sItems," - ")
   Dim $AutoItExe =  @ScriptDir & '\'& $spl[1] &'.exe'
   Dim $result = Run($AutoItExe)
   ;Dim $result = RunAs("bharathraj.vm","BRILLIO.COM","meaning@123$",0,$AutoItExe)
   ProcessWaitClose($result)
   Sleep(3000)
    Local $pid1 = ProcessExists("eclipse.exe")
   ProcessClose($epid)
   ProcessClose("javaw.exe")
   Next

EndIf
;******************************************************************************
wend