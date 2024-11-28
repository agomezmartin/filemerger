#include <File.au3>
#include <Array.au3>
#include <Strings.au3>
#include <ErrorMessages.au3>

Func BrowseForFolder($prompt)
    Local $folder = FileSelectFolder($prompt, "", 1)
    If @error Then Return ""
    Return $folder
EndFunc

Func MergeFiles($sourcePath, $destPath)
    If Not FileExists($sourcePath) Or Not FileExists($destPath) Then
        MsgBox(16, $ERRORS_Title, $ERRORS_InvalidPath)
        Return
    EndIf

    Local $fileTypes = [".txt", ".doc", ".docx", ".xlsx", ".xls", ".xlsm", ".ppt", ".pptx", ".xml", ".html", ".po", ".vb", ".dita"]

    For $i = 0 To UBound($fileTypes) - 1
        Local $files = _FileListToArray($sourcePath, "*" & $fileTypes[$i])
        If IsArray($files) Then
            MergeByType($files, $sourcePath, $destPath, $fileTypes[$i])
        EndIf
    Next
EndFunc

Func MergeByType($files, $sourcePath, $destPath, $fileType)
    Local $outputFile = $destPath & "\Merged" & $fileType
    For $i = 1 To $files[0]
        Local $fullPath = $sourcePath & "\" & $files[$i]
        _FileWriteToLine($outputFile, $i, FileRead($fullPath), 1)
    Next
    MsgBox(64, $STRINGS_MergeCompleteTitle, $STRINGS_MergeCompleteMsg & $fileType)
EndFunc
