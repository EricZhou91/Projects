Option Explicit
' mod_Helpers.bas
' Common helper routines for safe macro execution

Public Sub MacroBegin()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

Public Sub MacroEnd()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Public Function SheetExists(ByVal sheetName As String, Optional ByVal wb As Workbook) As Boolean
    Dim s As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set s = wb.Worksheets(sheetName)
    SheetExists = Not s Is Nothing
    On Error GoTo 0
End Function