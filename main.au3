#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <Functions.au3>
#include <Strings.au3>
#include <ErrorMessages.au3>

; Initialize the GUI
Opt("GUICoordMode", 1)
Opt("GUIResizeMode", 1)

Global $hGUI = GUICreate($STRINGS_Title, 400, 300, -1, -1, $WS_SIZEBOX + $WS_SYSMENU)
Global $txtSource = GUICtrlCreateInput("", 10, 10, 280, 20)
Global $btnBrowseSource = GUICtrlCreateButton($STRINGS_Browse, 300, 10, 80, 20)
Global $txtDestination = GUICtrlCreateInput("", 10, 50, 280, 20)
Global $btnBrowseDest = GUICtrlCreateButton($STRINGS_Browse, 300, 50, 80, 20)
Global $btnMerge = GUICtrlCreateButton($STRINGS_Merge, 150, 100, 100, 30)

GUISetState(@SW_SHOW)

; Main loop
While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            Exit
        Case $btnBrowseSource
            $sourcePath = BrowseForFolder($STRINGS_SelectSource)
            If $sourcePath <> "" Then GUICtrlSetData($txtSource, $sourcePath)
        Case $btnBrowseDest
            $destPath = BrowseForFolder($STRINGS_SelectDestination)
            If $destPath <> "" Then GUICtrlSetData($txtDestination, $destPath)
        Case $btnMerge
            MergeFiles(GUICtrlRead($txtSource), GUICtrlRead($txtDestination))
    EndSwitch
WEnd
