#include-once
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
            If $fileTypes[$i] = ".doc" Or $fileTypes[$i] = ".docx" Then
                MergeWordFiles($files, $sourcePath, $destPath, $fileTypes[$i])
            Else
                MergeByType($files, $sourcePath, $destPath, $fileTypes[$i])
            EndIf
        EndIf
        ; No message box is displayed if files are not found
    Next
EndFunc

Func MergeByType($files, $sourcePath, $destPath, $fileType)
    Local $outputFile = $destPath & "\Merged" & $fileType
    Local $hFile = FileOpen($outputFile, $FO_OVERWRITE + $FO_UTF8)

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

        FileWrite($hFile, $fileContent & @CRLF)
    Next

    FileClose($hFile)
    MsgBox(64, $STRINGS_MergeCompleteTitle, $STRINGS_MergeCompleteMsg & $fileType)
EndFunc

Func MergeWordFiles($files, $sourcePath, $destPath, $fileType)
    Local $oWord = ObjCreate("Word.Application")
    If Not IsObj($oWord) Then
        MsgBox(16, $ERRORS_Title, $ERRORS_NoWordCOM)
        Return
    EndIf

    $oWord.Visible = False
    Local $oDoc = $oWord.Documents.Add
    Local $outputFile = $destPath & "\Merged" & $fileType

    For $i = 1 To $files[0]
        Local $fullPath = $sourcePath & "\" & $files[$i]
        $oWord.Selection.InsertFile($fullPath)

        If @error Then
            MsgBox(16, $ERRORS_Title, $ERRORS_FileReadError & $files[$i])
            ContinueLoop
        EndIf

        $oWord.Selection.TypeParagraph() ; Add spacing between merged files
    Next

    $oDoc.SaveAs($outputFile)
    If @error Then
        MsgBox(16, $ERRORS_Title, $ERRORS_CannotWrite & $outputFile)
    Else
        MsgBox(64, $STRINGS_MergeCompleteTitle, $STRINGS_MergeCompleteMsg & $fileType)
    EndIf

    $oDoc.Close(False)
    $oWord.Quit()
EndFunc
