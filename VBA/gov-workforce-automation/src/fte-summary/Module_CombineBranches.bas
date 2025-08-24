Option Explicit
' Module_CombineBranches.bas (refactored)
' Adds small post-processing helpers

Public Sub HighlightMissingReportsTo()
    On Error GoTo CleanFail
    MacroBegin

    Dim ws As Worksheet
    Dim lastRowD As Long, lastRowF As Long
    Dim rngD As Range, rngF As Range, cellF As Range, found As Range

    Set ws = ActiveSheet
    lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    lastRowF = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    Set rngD = ws.Range("D2:D" & lastRowD)
    Set rngF = ws.Range("F2:F" & lastRowF)

    For Each cellF In rngF
        Set found = rngD.Find(What:=cellF.Value, LookIn:=xlValues, LookAt:=xlWhole)
        If found Is Nothing Then
            cellF.Interior.Color = vbYellow
        End If
    Next cellF

    MsgBox "Highlight complete.", vbInformation

CleanExit:
    MacroEnd
    Exit Sub
CleanFail:
    MsgBox "HighlightMissingReportsTo failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub MergeDuplicatePositionNumbers()
    On Error GoTo CleanFail
    MacroBegin

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim dict As Object, key As Variant

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    Set dict = CreateObject("Scripting.Dictionary")

    For i = lastRow To 2 Step -1
        key = ws.Cells(i, "D").Value
        If Not dict.Exists(key) Then
            dict.Add key, i
        Else
            ws.Cells(dict(key), "A").Value = ws.Cells(dict(key), "A").Value & ", " & ws.Cells(i, "A").Value
            ws.Cells(dict(key), "A").Interior.Color = vbYellow
            ws.Rows(i).Delete
        End If
    Next i

    MsgBox "Duplicate merge complete.", vbInformation

CleanExit:
    MacroEnd
    Exit Sub
CleanFail:
    MsgBox "MergeDuplicatePositionNumbers failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub