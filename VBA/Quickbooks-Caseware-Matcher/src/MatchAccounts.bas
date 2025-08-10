Option Explicit

'==================== CONFIG ====================
Const S1_NAME As String = "Sheet1" 'B = excel account name, A/C to fill
Const S2_NAME As String = "Sheet2" 'A = acct #, B = acct name

'Similarity thresholds (0..1)
Const GOOD_MATCH As Double = 0.84   'auto-assign & yellow highlight
Const MIN_MATCH  As Double = 0.74   'weaker suggestion & peach
'================================================

Public Sub MatchAccounts_All()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = FindSheet(S1_NAME)
    Set ws2 = FindSheet(S2_NAME)
    If ws1 Is Nothing Or ws2 Is Nothing Then
        MsgBox "Couldn't find """ & S1_NAME & """ and/or """ & S2_NAME & """ tabs.", vbCritical
        Exit Sub
    End If
    
    ' Fix format first
    fix_format
    
    ' 1) Exact pass (formulas)
    ExactMatch_XLOOKUP ws1, ws2
    
    ' 2) Fuzzy pass (values) for unresolved only
    FuzzyMatch_Levenshtein ws1, ws2
    
    MsgBox "All done: exact + fuzzy passes complete.", vbInformation
End Sub

'==================== FIX FORMAT ====================
Sub fix_format()
'
' fix_format Macro
' Start of fixing format of trial balance
'

'
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Account #"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Excel account name"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Fixed account name"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Debit"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Credit"
    Columns("A:F").Select
    Columns("A:F").EntireColumn.AutoFit
End Sub

'==================== EXACT PASS ====================
Private Sub ExactMatch_XLOOKUP(ws1 As Worksheet, ws2 As Worksheet)
    Dim lastRow As Long
    lastRow = ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ' Make sure formulas aren't treated as text
    ws1.Range("A2:A" & lastRow).NumberFormat = "General"
    ws1.Range("C2:C" & lastRow).NumberFormat = "General"
    
    ' Fill with XLOOKUP (same as your manual approach)
    ws1.Range("A2:A" & lastRow).Formula = _
        "=XLOOKUP($B2," & S2_NAME & "!$B:$B," & S2_NAME & "!$A:$A,"""")"
    ws1.Range("C2:C" & lastRow).Formula = _
        "=XLOOKUP($B2," & S2_NAME & "!$B:$B," & S2_NAME & "!$B:$B,"""")"
End Sub

'==================== FUZZY PASS ====================
Private Sub FuzzyMatch_Levenshtein(ws1 As Worksheet, ws2 As Worksheet)
    Dim last1 As Long, last2 As Long
    last1 = ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row
    last2 = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
    If last1 < 2 Or last2 < 2 Then Exit Sub
    
    ' Build candidate arrays from Sheet2
    Dim n As Long: n = last2 - 1
    If n <= 0 Then Exit Sub
    
    Dim candNorm() As String, candOrig() As String, candAcct() As Variant
    ReDim candNorm(1 To n)
    ReDim candOrig(1 To n)
    ReDim candAcct(1 To n)
    
    Dim i As Long
    For i = 2 To last2
        candOrig(i - 1) = CStr(ws2.Cells(i, "B").Value)
        candNorm(i - 1) = NormalizeText(candOrig(i - 1))
        candAcct(i - 1) = ws2.Cells(i, "A").Value
    Next
    
    ' Find/create a score column on Sheet1
    Dim scoreCol As Long
    scoreCol = Application.Max(6, ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column + 1)
    ws1.Cells(1, scoreCol).Value = "Match Type/Score"
    
    Application.ScreenUpdating = False
    
    Dim r As Long, nm As String, key As String
    Dim bestIdx As Long, bestScore As Double
    
    For r = 2 To last1
        ' Only rows where exact pass left A & C blank (XLOOKUP returned "")
        If Len(ws1.Cells(r, "A").Value) = 0 And Len(ws1.Cells(r, "C").Value) = 0 Then
            nm = CStr(ws1.Cells(r, "B").Value)
            key = NormalizeText(nm)
            
            BestMatchIndex key, candNorm, bestIdx, bestScore
            
            If bestIdx > 0 And bestScore >= GOOD_MATCH Then
                ws1.Cells(r, "A").Value = candAcct(bestIdx)
                ws1.Cells(r, "C").Value = candOrig(bestIdx)
                ws1.Cells(r, scoreCol).Value = "Fuzzy (" & Format(bestScore, "0.00") & ")"
                ws1.Rows(r).Interior.Color = RGB(255, 255, 153) 'light yellow
            ElseIf bestIdx > 0 And bestScore >= MIN_MATCH Then
                ws1.Cells(r, "A").Value = candAcct(bestIdx)
                ws1.Cells(r, "C").Value = candOrig(bestIdx)
                ws1.Cells(r, scoreCol).Value = "Possible (" & Format(bestScore, "0.00") & ")"
                ws1.Rows(r).Interior.Color = RGB(255, 230, 153)  'peach
            Else
                ws1.Cells(r, scoreCol).Value = "No good match"
                ws1.Rows(r).Interior.Color = RGB(255, 199, 206) 'light red
            End If
        End If
    Next r
    
    Application.ScreenUpdating = True
End Sub

'==================== FIND SHEET ====================
Private Function FindSheet(ByVal wanted As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(Trim$(ws.Name), Trim$(wanted), vbTextCompare) = 0 Then
            Set FindSheet = ws
            Exit Function
        End If
    Next ws
    Set FindSheet = Nothing
End Function

'==================== SIMILARITY HELPERS ====================
Private Sub BestMatchIndex(ByVal target As String, ByRef candNorm() As String, _
                           ByRef bestIdx As Long, ByRef bestScore As Double)
    Dim i As Long, score As Double
    bestIdx = 0: bestScore = -1
    For i = LBound(candNorm) To UBound(candNorm)
        score = BetterSimilarity(target, candNorm(i))
        If score > bestScore Then
            bestScore = score
            bestIdx = i
        End If
    Next
End Sub

' Improved similarity: token set + tails-after-colon + Levenshtein + containment
Private Function BetterSimilarity(ByVal a As String, ByVal b As String) As Double
    Dim s1 As String, s2 As String, t1 As String, t2 As String
    s1 = NormalizeText(a): s2 = NormalizeText(b)
    t1 = NormalizeText(TailAfterColon(a))
    t2 = NormalizeText(TailAfterColon(b))
    
    Dim scoreLev As Double, scoreTok As Double, scoreTail1 As Double, scoreTail2 As Double, scoreContain As Double
    scoreLev = SimilarityLev(s1, s2)
    scoreTok = TokenSetScore(s1, s2)
    scoreTail1 = SimilarityLev(t1, s2)
    scoreTail2 = SimilarityLev(s1, t2)
    scoreContain = ContainmentScore(s1, s2)
    
    BetterSimilarity = Application.Max(scoreLev, scoreTok, scoreTail1, scoreTail2, scoreContain)
End Function

Private Function TailAfterColon(ByVal s As String) As String
    Dim p As Long: p = InStrRev(s, ":")
    If p > 0 Then TailAfterColon = Mid$(s, p + 1) Else TailAfterColon = s
End Function

' Order-insensitive token overlap (like token_set_ratio)
Private Function TokenSetScore(ByVal s1 As String, ByVal s2 As String) As Double
    Dim set1 As Object, set2 As Object, tok As Variant
    Set set1 = CreateObject("Scripting.Dictionary")
    Set set2 = CreateObject("Scripting.Dictionary")
    set1.CompareMode = 1: set2.CompareMode = 1
    
    Dim arr1() As String, arr2() As String, w As Variant
    arr1 = Split(Trim$(s1)): arr2 = Split(Trim$(s2))
    
    For Each w In arr1
        If Len(w) > 0 Then set1(w) = True
    Next
    For Each w In arr2
        If Len(w) > 0 Then set2(w) = True
    Next
    
    Dim inter As Long, uni As Long
    inter = 0
    For Each tok In set1.Keys
        If set2.Exists(tok) Then inter = inter + 1
    Next
    uni = set1.Count + set2.Count - inter
    If uni <= 0 Then TokenSetScore = 1 Else TokenSetScore = inter / uni
End Function

' Containment bonus when one string includes the other
Private Function ContainmentScore(ByVal s1 As String, ByVal s2 As String) As Double
    Dim L1 As Long, L2 As Long
    L1 = Len(s1): L2 = Len(s2)
    If L1 = 0 Or L2 = 0 Then ContainmentScore = 0: Exit Function
    
    Dim score As Double: score = 0
    If InStr(1, s1, s2, vbTextCompare) > 0 Then score = Application.Max(score, L2 / L1)
    If InStr(1, s2, s1, vbTextCompare) > 0 Then score = Application.Max(score, L1 / L2)
    ContainmentScore = score
End Function

' Levenshtein wrapped to 0..1 similarity
Private Function SimilarityLev(a As String, b As String) As Double
    Dim d As Long, m As Long
    d = Levenshtein(a, b)
    m = Len(a): If Len(b) > m Then m = Len(b)
    If m = 0 Then SimilarityLev = 1 Else SimilarityLev = 1 - (d / m)
End Function

Private Function Levenshtein(ByVal s As String, ByVal t As String) As Long
    Dim m As Long, n As Long, i As Long, j As Long, cost As Long
    Dim d() As Long
    m = Len(s): n = Len(t)
    ReDim d(0 To m, 0 To n)
    For i = 0 To m: d(i, 0) = i: Next
    For j = 0 To n: d(0, j) = j: Next
    For i = 1 To m
        For j = 1 To n
            cost = IIf(Mid$(s, i, 1) = Mid$(t, j, 1), 0, 1)
            d(i, j) = Application.Min( _
                        d(i - 1, j) + 1, _
                        d(i, j - 1) + 1, _
                        d(i - 1, j - 1) + cost)
        Next j
    Next i
    Levenshtein = d(m, n)
End Function

'==================== NORMALIZE ====================
Private Function NormalizeText(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, "-", " ")
    s = Replace(s, ":", " ")
    s = Replace(s, "/", " ")
    s = Replace(s, ".", " ")
    s = Replace(s, ",", " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeText = s
End Function




