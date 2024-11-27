#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include "functions.au3" ; Import functions from the external file
#include "strings.au3"   ; Import string constants

Global $sInputPath = "", $sOutputPath = ""

; Create GUI window
$hGUI = GUICreate("File Merger", 400, 250)

; Input Path
GUICtrlCreateLabel($sLabelSelectInputPath, 10, 10)
$hInputPathButton = GUICtrlCreateButton($sButtonBrowse, 130, 10, 80, 30)
$hInputPathDisplay = GUICtrlCreateLabel("", 220, 10, 160, 30) ; Label to display input path

; Output Path
GUICtrlCreateLabel($sLabelSelectOutputPath, 10, 60)
$hOutputPathButton = GUICtrlCreateButton($sButtonBrowse, 130, 60, 80, 30)
$hOutputPathDisplay = GUICtrlCreateLabel("", 220, 60, 160, 30) ; Label to display output path

; Merge Button
$hMergeButton = GUICtrlCreateButton($sButtonMerge, 150, 120, 100, 40)

; Event loop
GUISetState()

While 1
    $nMsg = GUIGetMsg()

    Switch $nMsg
        Case $GUI_EVENT_CLOSE
            Exit

        Case $hInputPathButton
            ; Input Path Selection
            $sInputPath = FileSelectFolder("Select Folder Containing Text Files", $sInputPath)
            If @error Then
                ShowError($sErrorNoInputFolderSelected)
            Else
                ; Update input path label
                GUICtrlSetData($hInputPathDisplay, $sInputPath)
            EndIf

        Case $hOutputPathButton
            ; Output Path Selection
            $sOutputPath = FileSelectFolder("Select Output Folder", $sOutputPath)
            If @error Then
                ShowError($sErrorNoOutputFolderSelected)
            Else
                ; Update output path label
                GUICtrlSetData($hOutputPathDisplay, $sOutputPath)
            EndIf

        Case $hMergeButton
            ; Merge Button Logic
            If StringLen($sInputPath) == 0 Or StringLen($sOutputPath) == 0 Then
                ShowError($sErrorMissingPaths)
            Else
                ; Call the function to merge files
                $sResult = MergeFiles($sInputPath, $sOutputPath)
                If StringInStr($sResult, $sErrorNoTextFiles) > 0 Or StringInStr($sResult, $sErrorFailedToOpenFile) > 0 Then
                    ShowError($sResult)
                Else
                    ShowSuccess($sResult)
                EndIf
            EndIf
    EndSwitch
WEnd
