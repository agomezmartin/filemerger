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
        Else
            MsgBox(48, $STRINGS_NoFilesFoundTitle, $STRINGS_NoFilesFoundMsg & $fileTypes[$i])
        EndIf
    Next
EndFunc

Func MergeByType($files, $sourcePath, $destPath, $fileType)
    Local $outputFile = $destPath & "\Merged" & $fileType
    Local $hFile = FileOpen($outputFile, $FO_OVERWRITE)

    If $hFile = -1 Then
        MsgBox(16, $ERRORS_Title, $ERRORS_CannotWrite & $outputFile)
        Return
    EndIf

    For $i = 1 To $files[0]
        Local $fullPath = $sourcePath & "\" & $files[$i]
        Local $fileContent = FileRead($fullPath)

        If @error Then
            MsgBox(16, $ERRORS_Title, $ERRORS_FileReadError & $files[$i])
            ContinueLoop
        EndIf

        FileWrite($hFile, $fileContent & @CRLF) ; Add content and newline to separate files
    Next

    FileClose($hFile)
    MsgBox(64, $STRINGS_MergeCompleteTitle, $STRINGS_MergeCompleteMsg & $fileType)
EndFunc
