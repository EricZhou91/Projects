Option Explicit
' Module_FTECount.bas (refactored)
' Computes FTE per branch and builds an "FTE Summary" sheet

Private Function IsExcludedToken(ByVal fullNameLower As String) As Boolean
    Dim toks As Variant, i As Long
    toks = StatusTokensExcluded()
    For i = LBound(toks) To UBound(toks)
        If InStr(1, fullNameLower, LCase$(toks(i)), vbTextCompare) > 0 Then
            IsExcludedToken = True
            Exit Function
        End If
    Next i
End Function

Public Sub FTECountOnAllSheets()
    On Error GoTo CleanFail
    MacroBegin

    Dim wsSummary As Worksheet
    Dim b() As Variant, i As Long, countFTE As Long

    b = BranchNames()

    ' Create or clear summary
    If SheetExists("FTE Summary") Then
        Set wsSummary = ThisWorkbook.Worksheets("FTE Summary")
        wsSummary.Cells.Clear
    Else
        Set wsSummary = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count))
        wsSummary.Name = "FTE Summary"
    End If

    wsSummary.Range("A1").Value = "Branch"
    wsSummary.Range("B1").Value = "FTE"

    For i = LBound(b) To UBound(b)
        countFTE = FTECountOnSheet(CStr(b(i)))
        wsSummary.Cells(i + 2, 1).Value = CStr(b(i))
        wsSummary.Cells(i + 2, 2).Value = countFTE
    Next i

    wsSummary.Cells(UBound(b) + 3, 1).Value = "Total"
    wsSummary.Cells(UBound(b) + 3, 2).Formula = "=SUM(B2:B" & (UBound(b) + 2) & ")"
    wsSummary.Columns("A:B").AutoFit

    MsgBox "FTE Summary generated.", vbInformation

CleanExit:
    MacroEnd
    Exit Sub
CleanFail:
    MsgBox "FTECountOnAllSheets failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Function FTECountOnSheet(ByVal sheetName As String) As Long
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim nm As String, cnt As Long

    If Not SheetExists(sheetName) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(sheetName)

    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row  ' assumes data present in column D
    For i = 2 To lastRow
        nm = LCase$(Trim$(CStr(ws.Cells(i, "A").Value)))
        If nm <> vbNullString Then
            If Not IsExcludedToken(nm) Then
                cnt = cnt + 1
            End If
        End If
    Next i
    ' write label/value on sheet (optional)
    ws.Range("I2").Value = "Number of FTE:"
    ws.Range("J2").Value = cnt

    FTECountOnSheet = cnt
End Function

Public Sub CombineBranchesToList()
    On Error GoTo CleanFail
    MacroBegin

    Dim wsList As Worksheet
    Dim b() As Variant, i As Long
    Dim ws As Worksheet, lastRow As Long, nextRow As Long

    b = BranchNames()

    If SheetExists("Combined list") Then
        Set wsList = ThisWorkbook.Worksheets("Combined list")
        wsList.Cells.Clear
    Else
        Set wsList = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count))
        wsList.Name = "Combined list"
    End If

    wsList.Range("A1:F1").Value = Array("Name","Position Title","Department ID","Position Number","Job Code","Reports to")

    nextRow = 2
    For i = LBound(b) To UBound(b)
        If SheetExists(CStr(b(i))) Then
            Set ws = ThisWorkbook.Worksheets(CStr(b(i)))
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            If lastRow >= 2 Then
                wsList.Range("A" & nextRow).Resize(lastRow - 1, 6).Value = ws.Range("A2:F" & lastRow).Value
                nextRow = nextRow + (lastRow - 1)
            End If
        End If
    Next i

    wsList.Columns("A:F").AutoFit
    MsgBox "Combined list built.", vbInformation

CleanExit:
    MacroEnd
    Exit Sub
CleanFail:
    MsgBox "CombineBranchesToList failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub