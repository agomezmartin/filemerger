#include "strings.au3"  ; Import string constants

; Check if the function is already defined
#IfNotDef _FileListToArray
    ; Function to list files in a directory
    Func _FileListToArray($sDir, $sPattern = "*.*", $iFlag = 0)
        Local $aFiles[1]
        Local $hSearch, $sFile
        $hSearch = FileFindFirstFile($sDir & "\" & $sPattern)
        If $hSearch = -1 Then Return SetError(1, 0, "")

        While 1
            $sFile = FileFindNextFile($hSearch)
            If @error Then ExitLoop
            If BitAND(FileGetAttrib($sFile), 16) = 0 Then ; Ignore directories
                $aFiles[0] += 1
                ReDim $aFiles[$aFiles[0] + 1]
                $aFiles[$aFiles[0]] = $sFile
            EndIf
        Wend
        FileClose($hSearch)
        If $aFiles[0] = 0 Then Return SetError(1, 0, "")
        Return $aFiles
    EndFunc
#EndIf

; Function to merge files
Func MergeFiles($sInputPath, $sOutputPath)
    ; Validate that there are text files in the input path
    $aFiles = _FileListToArray($sInputPath, "*.txt", 1)
    If @error Then
        Return $sErrorNoTextFiles
    Else
        ; Merge files
        $sMergedContent = ""
        For $i = 1 To $aFiles[0]
            $sFilePath = $sInputPath & "\" & $aFiles[$i]
            $sMergedContent &= FileRead($sFilePath) & @CRLF
        Next

        ; Save the merged content to the output folder
        $sOutputFile = $sOutputPath & "\merged_output.txt"
        $hFile = FileOpen($sOutputFile, 2) ; 2 = overwrite mode
        If $hFile = -1 Then
            Return $sErrorFailedToOpenFile
        Else
            FileWrite($hFile, $sMergedContent)
            FileClose($hFile)
            Return $sSuccessFilesMerged
        EndIf
    EndIf
EndFunc

; Function to display error message
Func ShowError($sMessage)
    MsgBox($MB_ICONERROR, "Error", $sMessage)
EndFunc

; Function to display success message
Func ShowSuccess($sMessage)
    MsgBox($MB_ICONINFORMATION, "Success", $sMessage)
EndFunc
