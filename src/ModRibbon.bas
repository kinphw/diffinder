Attribute VB_Name = "ModRibbon"
Option Explicit

' Ribbon callbacks. The custom ribbon (ribbon/customUI14.xml) wires its buttons
' here. Each callback must take an IRibbonControl argument, so these thin
' wrappers simply forward to the public entry points in ModDiff.

Public Const APP_VERSION As String = "Diffinder 0.1.0"

Public Sub Ribbon_PopupPicker(ByVal control As IRibbonControl)
    Diff_ByPopupPicker
End Sub

Public Sub Ribbon_SelectedCells(ByVal control As IRibbonControl)
    Diff_SelectedThreeCells
End Sub

Public Sub Ribbon_ShowAbout(ByVal control As IRibbonControl)
    MsgBox APP_VERSION, vbInformation, "Diffinder"
End Sub
