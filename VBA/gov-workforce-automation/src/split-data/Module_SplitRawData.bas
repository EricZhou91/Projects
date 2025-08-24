Option Explicit
' Module_SplitRawData.bas (refactored)

Private Function BranchFromDept(ByVal deptId As Long, ByVal wb As Workbook) As String
    Dim ws As Worksheet, f As Range
    Set ws = wb.Worksheets(SHEET_BRANCH_LOOKUP)
    Set f = ws.Columns(1).Find(What:=deptId, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        BranchFromDept = CStr(ws.Cells(f.Row, 2).Value)
    End If
End Function

Private Function AddStatusTokens(ByVal fullName As String, ByVal incumbency As String, ByVal expReturn As String, ByVal classDesc As String) As String
    Dim nm As String
    nm = fullName
    If incumbency = "OUT" Then
        nm = nm & " (A/O)"
    ElseIf incumbency = "IN " Then
        nm = nm & " (A/I)"
    ElseIf Len(Trim$(expReturn)) > 0 Then
        nm = nm & " (LoA)"
    End If
    If classDesc = "Fixed Term" Then nm = nm & " (FxT)"
    AddStatusTokens = nm
End Function

Public Sub SplitRawData()
    On Error GoTo CleanFail
    MacroBegin

    Dim p As Variant, wb As Workbook
    p = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*")
    If VarType(p) = vbBoolean And p = False Then GoTo CleanExit

    Set wb = Workbooks.Open(CStr(p))
    CreateContactLists wb
    wb.Close SaveChanges:=False

    MsgBox "Split pipeline complete (new workbook created).", vbInformation

CleanExit:
    MacroEnd
    Exit Sub
CleanFail:
    MsgBox "SplitRawData failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Sub CreateContactLists(ByVal srcWb As Workbook)
    Dim wsRaw As Worksheet
    Set wsRaw = srcWb.Worksheets(SHEET_RAW_DATA)

    Dim outWb As Workbook
    Set outWb = Workbooks.Add

    Dim branches As Variant, i As Long
    branches = BranchNames()

    ' Create Director & Manager first
    AddContactSheet outWb, "Director"
    AddContactSheet outWb, "Manager"

    For i = LBound(branches) To UBound(branches)
        AddContactSheet outWb, CStr(branches(i))
    Next i

    Dim lastRow As Long, r As Long
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, COL_DEPT_ID).End(xlUp).Row

    Dim deptId As Long, branch As String
    Dim firstName As String, lastName As String, pos As String
    Dim jobCode As Variant, posNum As Variant, rptTo As Variant
    Dim incumb As String, expReturn As String, classDesc As String
    Dim fullName As String

    For r = 2 To lastRow
        deptId = CLng(wsRaw.Range(COL_DEPT_ID & r).Value)
        branch = BranchFromDept(deptId, srcWb)

        firstName = CStr(wsRaw.Range(COL_FIRST_NAME & r).Value)
        lastName  = CStr(wsRaw.Range(COL_LAST_NAME & r).Value)
        pos       = CStr(wsRaw.Range(COL_POSITION & r).Value)
        jobCode   = wsRaw.Range(COL_JOB_CODE & r).Value
        posNum    = wsRaw.Range(COL_POSITIONNO & r).Value
        rptTo     = wsRaw.Range(COL_REPORTSTO & r).Value

        incumb    = CStr(wsRaw.Range(COL_INCUMBENCY & r).Value)
        expReturn = CStr(wsRaw.Range(COL_EXP_RETURN & r).Value)
        classDesc = CStr(wsRaw.Range(COL_CLASS_DESC & r).Value)

        fullName = AddStatusTokens(firstName & " " & lastName, incumb, expReturn, classDesc)

        ' Director / Manager buckets
        If InStr(1, pos, "Director", vbTextCompare) > 0 Or _
           pos = "Chief Internal Auditor" Or pos = "Chief Internal Auditor/ADM" Then
            AppendRow outWb.Worksheets("Director"), fullName, pos, deptId, posNum, jobCode, rptTo
        ElseIf (InStr(1, pos, "Manager", vbTextCompare) > 0 Or _
                InStr(1, pos, "Mgr", vbTextCompare) > 0 Or _
                pos = "Exec. Lead & Strategic Advisor") And _
                pos <> "Audit Project Manager" And InStr(1, pos, "LTIP", vbTextCompare) = 0 Then
            AppendRow outWb.Worksheets("Manager"), fullName, pos, deptId, posNum, jobCode, rptTo
        End If

        ' Branch buckets (exclude LTIP/Student/Assistant/Intern)
        If Len(branch) > 0 Then
            If InStr(1, pos, "LTIP", vbTextCompare) = 0 And _
               InStr(1, pos, "Student", vbTextCompare) = 0 And _
               InStr(1, pos, "Assistant", vbTextCompare) = 0 And _
               InStr(1, " " & pos & " ", " Intern ", vbTextCompare) = 0 Then
                AppendRow outWb.Worksheets(branch), fullName, pos, deptId, posNum, jobCode, rptTo
            End If
        End If
    Next r

    ' Post-processing on Manager sheet
    CombineDirectorsToManagers outWb
    MergeDuplicatePositionNumbersSheet outWb.Worksheets("Manager")
    HighlightMissingReportsToSheet outWb.Worksheets("Manager")

    On Error Resume Next
    outWb.Worksheets("Sheet1").Delete
    On Error GoTo 0
End Sub

Private Sub AddContactSheet(ByVal wb As Workbook, ByVal name As String)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    ws.Range("A1:F1").Value = Array("Name","Position Title","Department ID","Position Number","Job Code","Reports to")
End Sub

Private Sub AppendRow(ByVal ws As Worksheet, ByVal nm As String, ByVal pos As String, _
                      ByVal deptId As Variant, ByVal posNum As Variant, _
                      ByVal jobCode As Variant, ByVal rptTo As Variant)
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(r, "A").Value = nm
    ws.Cells(r, "B").Value = pos
    ws.Cells(r, "C").Value = deptId
    ws.Cells(r, "D").Value = posNum
    ws.Cells(r, "E").Value = jobCode
    ws.Cells(r, "F").Value = rptTo
End Sub

Private Sub CombineDirectorsToManagers(ByVal wb As Workbook)
    Dim wsDir As Worksheet, wsMgr As Worksheet, lr As Long, nextRow As Long
    Set wsDir = wb.Worksheets("Director")
    Set wsMgr = wb.Worksheets("Manager")

    lr = wsDir.Cells(wsDir.Rows.Count, "A").End(xlUp).Row
    If lr < 2 Then Exit Sub

    nextRow = wsMgr.Cells(wsMgr.Rows.Count, "A").End(xlUp).Row + 1
    wsMgr.Range("A2:F" & lr).Copy
    wsMgr.Range("A" & nextRow).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
End Sub

Private Sub MergeDuplicatePositionNumbersSheet(ByVal ws As Worksheet)
    Dim lastRow As Long, i As Long
    Dim dict As Object, key As Variant
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
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
End Sub

Private Sub HighlightMissingReportsToSheet(ByVal ws As Worksheet)
    Dim lastRowD As Long, lastRowF As Long
    Dim rngD As Range, rngF As Range, cellF As Range, found As Range

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
End Sub