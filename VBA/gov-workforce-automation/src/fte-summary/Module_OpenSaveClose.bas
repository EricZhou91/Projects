Option Explicit
' Module_OpenSaveClose.bas (refactored)

Public Sub OpenSaveClose()
    On Error GoTo CleanFail
    MacroBegin

    Dim fd As FileDialog
    Dim wb As Workbook

    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo CleanExit
        Set wb = Workbooks.Open(.SelectedItems(1))
    End With

    ' Activate next sheet if possible
    Dim idx As Long
    idx = ActiveSheet.Index
    If idx < wb.Sheets.Count Then wb.Sheets(idx + 1).Activate

    wb.Save
    wb.Close SaveChanges:=False
    MsgBox "Workbook processed (activated next sheet, saved, closed).", vbInformation

CleanExit:
    MacroEnd
    Exit Sub
CleanFail:
    MsgBox "OpenSaveClose failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub