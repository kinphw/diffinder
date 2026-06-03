Attribute VB_Name = "ModRibbon"
Option Explicit

' Ribbon callbacks. The custom ribbon (ribbon/customUI14.xml) wires its buttons
' here. Each callback must take an IRibbonControl argument, so these thin
' wrappers simply forward to the public entry points in ModDiff.

' SSOT: these literals are overwritten at build time from version.json by
' build.ps1. The values here are only fallbacks for a manual/unbuilt import.
Public Const APP_NAME As String = "Diffinder"
Public Const APP_VERSION As String = "0.0.0-dev"
Public Const APP_AUTHOR As String = "unknown"

Public Sub Ribbon_PopupPicker(ByVal control As IRibbonControl)
    Diff_ByPopupPicker
End Sub

Public Sub Ribbon_SelectedCells(ByVal control As IRibbonControl)
    Diff_SelectedThreeCells
End Sub

Public Sub Ribbon_ShowAbout(ByVal control As IRibbonControl)
    MsgBox APP_NAME & " " & APP_VERSION & vbCrLf & _
           "author : " & APP_AUTHOR, vbInformation, APP_NAME
End Sub
