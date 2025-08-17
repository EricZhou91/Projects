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
    
    ' 1) Exact pass (formulas) - unchanged for reliability
    ExactMatch_XLOOKUP ws1, ws2
    
    ' 2) Simple fast fuzzy pass for unresolved only
    SimpleFuzzyMatch ws1, ws2
    
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

'==================== SIMPLE FAST FUZZY PASS ====================
Private Sub SimpleFuzzyMatch(ws1 As Worksheet, ws2 As Worksheet)
    Dim last1 As Long, last2 As Long
    last1 = ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row
    last2 = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
    If last1 < 2 Or last2 < 2 Then Exit Sub
    
    ' Find/create a score column on Sheet1
    Dim scoreCol As Long
    scoreCol = Application.Max(6, ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column + 1)
    ws1.Cells(1, scoreCol).Value = "Match Type/Score"
    
    Application.ScreenUpdating = False
    
    Dim r As Long, i As Long
    Dim targetName As String, targetClean As String, targetTail As String
    Dim candidateName As String, candidateClean As String, candidateTail As String
    Dim bestMatch As Long, bestScore As Double, currentScore As Double
    
    For r = 2 To last1
        ' Only process rows where exact pass left A & C blank (XLOOKUP returned "")
        If Len(ws1.Cells(r, "A").Value) = 0 And Len(ws1.Cells(r, "C").Value) = 0 Then
            targetName = CStr(ws1.Cells(r, "B").Value)
            targetClean = CleanString(targetName)
            targetTail = CleanString(TailAfterColon(targetName))
            
            bestMatch = 0
            bestScore = 0
            
            ' Simple loop through all candidates
            For i = 2 To last2
                candidateName = CStr(ws2.Cells(i, "B").Value)
                candidateClean = CleanString(candidateName)
                candidateTail = CleanString(TailAfterColon(candidateName))
                
                ' Fast similarity check
                currentScore = FastSimpleSimilarity(targetClean, targetTail, candidateClean, candidateTail)
                
                If currentScore > bestScore Then
                    bestScore = currentScore
                    bestMatch = i
                End If
            Next i
            
            ' Apply results
            If bestMatch > 0 And bestScore >= GOOD_MATCH Then
                ws1.Cells(r, "A").Value = ws2.Cells(bestMatch, "A").Value
                ws1.Cells(r, "C").Value = ws2.Cells(bestMatch, "B").Value
                ws1.Cells(r, scoreCol).Value = "Fuzzy (" & Format(bestScore, "0.00") & ")"
                ws1.Rows(r).Interior.Color = RGB(255, 255, 153) 'light yellow
            ElseIf bestMatch > 0 And bestScore >= MIN_MATCH Then
                ws1.Cells(r, "A").Value = ws2.Cells(bestMatch, "A").Value
                ws1.Cells(r, "C").Value = ws2.Cells(bestMatch, "B").Value
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

'==================== FAST SIMPLE SIMILARITY ====================
Private Function FastSimpleSimilarity(ByVal targetClean As String, ByVal targetTail As String, _
                                     ByVal candidateClean As String, ByVal candidateTail As String) As Double
    Dim bestScore As Double: bestScore = 0
    
    ' 1. Exact match (highest priority)
    If targetClean = candidateClean Then
        FastSimpleSimilarity = 1.0
        Exit Function
    End If
    
    ' 2. Tail exact match (for prefix handling)
    If targetTail = candidateClean Or targetClean = candidateTail Or targetTail = candidateTail Then
        FastSimpleSimilarity = 0.95
        Exit Function
    End If
    
    ' 3. Contains check (one string contains the other)
    If Len(targetClean) > 5 And Len(candidateClean) > 5 Then
        If InStr(1, targetClean, candidateClean, vbTextCompare) > 0 Or _
           InStr(1, candidateClean, targetClean, vbTextCompare) > 0 Then
            bestScore = 0.85
        End If
    End If
    
    ' 4. Tail contains check
    If Len(targetTail) > 3 And Len(candidateClean) > 3 Then
        If InStr(1, targetTail, candidateClean, vbTextCompare) > 0 Or _
           InStr(1, candidateClean, targetTail, vbTextCompare) > 0 Then
            If 0.90 > bestScore Then bestScore = 0.90
        End If
    End If
    
    ' 5. Simple word overlap (count common words)
    Dim targetWords() As String, candidateWords() As String
    targetWords = Split(targetClean)
    candidateWords = Split(candidateClean)
    
    Dim commonWords As Long, totalWords As Long
    commonWords = CountCommonWords(targetWords, candidateWords)
    totalWords = UBound(targetWords) + UBound(candidateWords) + 2 - commonWords
    
    If totalWords > 0 Then
        Dim wordScore As Double: wordScore = commonWords / totalWords
        If wordScore > bestScore Then bestScore = wordScore
    End If
    
    FastSimpleSimilarity = bestScore
End Function

'==================== HELPER FUNCTIONS ====================
Private Function TailAfterColon(ByVal s As String) As String
    Dim p As Long: p = InStrRev(s, ":")
    If p > 0 Then 
        TailAfterColon = Trim(Mid$(s, p + 1))
    Else 
        TailAfterColon = s
    End If
End Function

Private Function CleanString(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, "-", " ")
    s = Replace(s, ":", " ")
    s = Replace(s, "/", " ")
    s = Replace(s, ".", " ")
    s = Replace(s, ",", " ")
    s = Replace(s, "&", " and ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanString = s
End Function

Private Function CountCommonWords(ByRef words1() As String, ByRef words2() As String) As Long
    Dim i As Long, j As Long, count As Long
    count = 0
    
    For i = LBound(words1) To UBound(words1)
        For j = LBound(words2) To UBound(words2)
            If words1(i) = words2(j) And Len(words1(i)) > 2 Then
                count = count + 1
                Exit For
            End If
        Next j
    Next i
    
    CountCommonWords = count
End Function

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




