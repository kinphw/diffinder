Attribute VB_Name = "ModDiff"
Option Explicit

Private Const DIFF_EQUAL As Long = 0
Private Const DIFF_DELETE As Long = 1
Private Const DIFF_ADD As Long = 2

Private Const ALIGN_PAIR As Long = 1
Private Const ALIGN_DELETE As Long = 2
Private Const ALIGN_ADD As Long = 3

Private Const LINE_SIM_THRESHOLD As Double = 0.5
Private Const FLOAT_EPSILON As Double = 0.000001

Private Const MAX_LINE_LCS_LINES As Long = 300
Private Const MAX_LINE_LCS_CELLS As Long = 90000
Private Const MAX_CHAR_DIFF_CHARS As Long = 2000
Private Const MAX_CHAR_LCS_CELLS As Long = 1000000
Private Const MAX_WORD_DIFF_TOKENS As Long = 400
Private Const MAX_WORD_LCS_CELLS As Long = 90000

' =========================================================
' Runner
' =========================================================
Public Sub Diff_ByPopupPicker()
    Dim oldCell As Range
    Dim newCell As Range
    Dim outCell As Range

    If Not Runner_TryPickCell("원문 셀을 선택하세요.", oldCell) Then Exit Sub
    If Not Runner_TryPickCell("수정본 셀을 선택하세요.", newCell) Then Exit Sub
    If Not Runner_TryPickCell("결과 셀을 선택하세요.", outCell) Then Exit Sub

    Runner_Execute oldCell, newCell, outCell
End Sub

Public Sub Diff_SelectedThreeCells()
    Dim selectedRange As Range
    Dim targetSheet As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Dim oldCol As Long
    Dim newCol As Long
    Dim outCol As Long
    Dim rowIndex As Long

    If TypeName(Selection) <> "Range" Then
        MsgBox "전/후/비교 3개 셀 또는 전/후 2개 셀을 선택해 주세요.", vbExclamation
        Exit Sub
    End If

    Set selectedRange = Selection

    If selectedRange.Rows.Count = 1 And selectedRange.Columns.Count = 1 Then
        Runner_ExecuteActiveCell selectedRange.Cells(1, 1)
        Exit Sub
    End If

    If Not Runner_TryResolveSelectedRange( _
        selectedRange, targetSheet, firstRow, lastRow, oldCol, newCol, outCol _
    ) Then
        Exit Sub
    End If

    On Error GoTo CleanFail

    Application.ScreenUpdating = False

    For rowIndex = firstRow To lastRow
        Runner_RenderDiff _
            targetSheet.Cells(rowIndex, oldCol), _
            targetSheet.Cells(rowIndex, newCol), _
            targetSheet.Cells(rowIndex, outCol)
    Next rowIndex

    Application.ScreenUpdating = True
    MsgBox "비교가 완료되었습니다.", vbInformation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "비교 중 오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Public Sub Diff_ActiveCellThreeCells()
    If TypeName(ActiveCell) <> "Range" Then
        MsgBox "수정본 셀을 선택해 주세요. 왼쪽은 전, 오른쪽은 비교 결과로 사용합니다.", vbExclamation
        Exit Sub
    End If

    Runner_ExecuteActiveCell ActiveCell
End Sub

Private Sub Runner_Execute( _
    ByVal oldCell As Range, _
    ByVal newCell As Range, _
    ByVal outCell As Range _
)
    On Error GoTo CleanFail

    Application.ScreenUpdating = False

    Runner_RenderDiff oldCell, newCell, outCell

    Application.ScreenUpdating = True
    MsgBox "비교가 완료되었습니다.", vbInformation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "비교 중 오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Private Sub Runner_ExecuteActiveCell(ByVal centerCell As Range)
    If centerCell.Column <= 1 Or centerCell.Column >= centerCell.Worksheet.Columns.Count Then
        MsgBox "현재 셀 기준으로 왼쪽=전, 현재=후, 오른쪽=비교가 되도록 가운데 셀을 선택해 주세요.", vbExclamation
        Exit Sub
    End If

    On Error GoTo CleanFail

    Application.ScreenUpdating = False

    Runner_RenderDiff centerCell.Offset(0, -1), centerCell, centerCell.Offset(0, 1)

    Application.ScreenUpdating = True
    MsgBox "비교가 완료되었습니다.", vbInformation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "비교 중 오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Private Function Runner_TryResolveSelectedRange( _
    ByVal selectedRange As Range, _
    ByRef targetSheet As Worksheet, _
    ByRef firstRow As Long, _
    ByRef lastRow As Long, _
    ByRef oldCol As Long, _
    ByRef newCol As Long, _
    ByRef outCol As Long _
) As Boolean
    If selectedRange.Areas.Count > 1 Then
        MsgBox "하나의 연속된 범위만 선택해 주세요.", vbExclamation
        Exit Function
    End If

    Set targetSheet = selectedRange.Worksheet

    Select Case selectedRange.Columns.Count
        Case 2, 3
            oldCol = selectedRange.Column
            newCol = oldCol + 1
            outCol = oldCol + 2

        Case Else
            MsgBox "전/후 2열 또는 전/후/비교 3열을 선택해 주세요. 예: A:B 또는 A:C", vbExclamation
            Exit Function
    End Select

    If outCol > targetSheet.Columns.Count Then
        MsgBox "결과를 쓸 오른쪽 열이 없습니다.", vbExclamation
        Exit Function
    End If

    Runner_GetSelectedRowBounds selectedRange, firstRow, lastRow
    Runner_TryResolveSelectedRange = True
End Function

Private Sub Runner_GetSelectedRowBounds( _
    ByVal selectedRange As Range, _
    ByRef firstRow As Long, _
    ByRef lastRow As Long _
)
    Dim usedRange As Range

    If selectedRange.Rows.Count = selectedRange.Worksheet.Rows.Count Then
        Set usedRange = selectedRange.Worksheet.usedRange
        firstRow = usedRange.Row
        lastRow = usedRange.Row + usedRange.Rows.Count - 1
    Else
        firstRow = selectedRange.Row
        lastRow = selectedRange.Row + selectedRange.Rows.Count - 1
    End If
End Sub

Private Sub Runner_RenderDiff( _
    ByVal oldCell As Range, _
    ByVal newCell As Range, _
    ByVal outCell As Range _
)
    Dim resultText As String
    Dim starts() As Long
    Dim lens() As Long
    Dim styles() As Long
    Dim segCount As Long

    Engine_BuildDiff _
        NormalizeText(CStr(oldCell.Value2)), _
        NormalizeText(CStr(newCell.Value2)), _
        resultText, starts, lens, styles, segCount

    Renderer_Render outCell, resultText, starts, lens, styles, segCount
End Sub

Private Function Runner_TryPickCell( _
    ByVal promptText As String, _
    ByRef pickedCell As Range _
) As Boolean
    Dim picked As Range

    Set pickedCell = Nothing

    On Error Resume Next
    Set picked = Application.InputBox( _
        Prompt:=promptText, _
        Title:="Excel Diff", _
        Type:=8 _
    )
    On Error GoTo 0

    If picked Is Nothing Then Exit Function

    Set pickedCell = picked.Cells(1, 1)
    Runner_TryPickCell = True
End Function

' =========================================================
' Engine
' =========================================================
Private Sub Engine_BuildDiff( _
    ByVal oldText As String, _
    ByVal newText As String, _
    ByRef resultText As String, _
    ByRef starts() As Long, _
    ByRef lens() As Long, _
    ByRef styles() As Long, _
    ByRef segCount As Long _
)
    Dim oldLines() As String
    Dim newLines() As String
    Dim oldCount As Long
    Dim newCount As Long
    Dim hasOutputLine As Boolean

    Engine_SplitLines oldText, oldLines, oldCount
    Engine_SplitLines newText, newLines, newCount

    resultText = vbNullString
    segCount = 0
    hasOutputLine = False

    If oldText = newText Then
        AddSegment resultText, starts, lens, styles, segCount, oldText, DIFF_EQUAL
        Exit Sub
    End If

    If Not Engine_CanUseLineLcs(oldCount, newCount) Then
        Engine_RenderAllDeleteAdd _
            oldLines, oldCount, _
            newLines, newCount, _
            resultText, starts, lens, styles, segCount, hasOutputLine
        Exit Sub
    End If

    Engine_RenderLineDiff _
        oldLines, oldCount, _
        newLines, newCount, _
        resultText, starts, lens, styles, segCount, hasOutputLine
End Sub

Private Sub Engine_SplitLines( _
    ByVal textValue As String, _
    ByRef lines() As String, _
    ByRef lineCount As Long _
)
    Dim raw() As String
    Dim i As Long

    If Len(textValue) = 0 Then
        lineCount = 0
        ReDim lines(1 To 1)
        Exit Sub
    End If

    raw = Split(textValue, vbLf)
    lineCount = UBound(raw) - LBound(raw) + 1
    ReDim lines(1 To lineCount)

    For i = 0 To lineCount - 1
        lines(i + 1) = raw(i)
    Next i
End Sub

Private Function Engine_CanUseLineLcs( _
    ByVal oldCount As Long, _
    ByVal newCount As Long _
) As Boolean
    If oldCount = 0 Or newCount = 0 Then
        Engine_CanUseLineLcs = True
        Exit Function
    End If

    If oldCount >= MAX_LINE_LCS_LINES Or newCount >= MAX_LINE_LCS_LINES Then Exit Function
    If CDbl(oldCount) * CDbl(newCount) > MAX_LINE_LCS_CELLS Then Exit Function

    Engine_CanUseLineLcs = True
End Function

Private Sub Engine_RenderAllDeleteAdd( _
    ByRef oldLines() As String, _
    ByVal oldCount As Long, _
    ByRef newLines() As String, _
    ByVal newCount As Long, _
    ByRef resultText As String, _
    ByRef starts() As Long, _
    ByRef lens() As Long, _
    ByRef styles() As Long, _
    ByRef segCount As Long, _
    ByRef hasOutputLine As Boolean _
)
    Dim i As Long

    For i = 1 To oldCount
        Engine_AppendWholeLine _
            oldLines(i), DIFF_DELETE, _
            resultText, starts, lens, styles, segCount, hasOutputLine
    Next i

    For i = 1 To newCount
        Engine_AppendWholeLine _
            newLines(i), DIFF_ADD, _
            resultText, starts, lens, styles, segCount, hasOutputLine
    Next i
End Sub

Private Sub Engine_RenderLineDiff( _
    ByRef oldLines() As String, _
    ByVal oldCount As Long, _
    ByRef newLines() As String, _
    ByVal newCount As Long, _
    ByRef resultText As String, _
    ByRef starts() As Long, _
    ByRef lens() As Long, _
    ByRef styles() As Long, _
    ByRef segCount As Long, _
    ByRef hasOutputLine As Boolean _
)
    Dim dp() As Long
    Dim i As Long
    Dim j As Long
    Dim oldGapStart As Long
    Dim newGapStart As Long

    ReDim dp(0 To oldCount, 0 To newCount)

    For i = oldCount - 1 To 0 Step -1
        For j = newCount - 1 To 0 Step -1
            If oldLines(i + 1) = newLines(j + 1) Then
                dp(i, j) = dp(i + 1, j + 1) + 1
            ElseIf dp(i + 1, j) >= dp(i, j + 1) Then
                dp(i, j) = dp(i + 1, j)
            Else
                dp(i, j) = dp(i, j + 1)
            End If
        Next j
    Next i

    i = 0
    j = 0
    oldGapStart = 1
    newGapStart = 1

    Do While i < oldCount Or j < newCount
        If i < oldCount And j < newCount Then
            If oldLines(i + 1) = newLines(j + 1) Then
                Engine_RenderGap _
                    oldLines, oldGapStart, i, _
                    newLines, newGapStart, j, _
                    resultText, starts, lens, styles, segCount, hasOutputLine

                Engine_AppendWholeLine _
                    oldLines(i + 1), DIFF_EQUAL, _
                    resultText, starts, lens, styles, segCount, hasOutputLine

                i = i + 1
                j = j + 1
                oldGapStart = i + 1
                newGapStart = j + 1
            ElseIf dp(i + 1, j) >= dp(i, j + 1) Then
                i = i + 1
            Else
                j = j + 1
            End If
        ElseIf i < oldCount Then
            i = i + 1
        Else
            j = j + 1
        End If
    Loop

    Engine_RenderGap _
        oldLines, oldGapStart, oldCount, _
        newLines, newGapStart, newCount, _
        resultText, starts, lens, styles, segCount, hasOutputLine
End Sub

Private Sub Engine_RenderGap( _
    ByRef oldLines() As String, _
    ByVal oldFrom As Long, _
    ByVal oldTo As Long, _
    ByRef newLines() As String, _
    ByVal newFrom As Long, _
    ByVal newTo As Long, _
    ByRef resultText As String, _
    ByRef starts() As Long, _
    ByRef lens() As Long, _
    ByRef styles() As Long, _
    ByRef segCount As Long, _
    ByRef hasOutputLine As Boolean _
)
    Dim oldBlockCount As Long
    Dim newBlockCount As Long
    Dim i As Long

    oldBlockCount = Engine_BlockCount(oldFrom, oldTo)
    newBlockCount = Engine_BlockCount(newFrom, newTo)

    If oldBlockCount = 0 And newBlockCount = 0 Then Exit Sub

    If oldBlockCount = 0 Then
        For i = newFrom To newTo
            Engine_AppendWholeLine _
                newLines(i), DIFF_ADD, _
                resultText, starts, lens, styles, segCount, hasOutputLine
        Next i
        Exit Sub
    End If

    If newBlockCount = 0 Then
        For i = oldFrom To oldTo
            Engine_AppendWholeLine _
                oldLines(i), DIFF_DELETE, _
                resultText, starts, lens, styles, segCount, hasOutputLine
        Next i
        Exit Sub
    End If

    ' Exact line anchors already split the document. Inside each changed block,
    ' align only reasonably similar lines so unrelated lines stay as full delete/add.
    Engine_AlignChangedBlock _
        oldLines, oldFrom, oldTo, _
        newLines, newFrom, newTo, _
        resultText, starts, lens, styles, segCount, hasOutputLine
End Sub

Private Sub Engine_AlignChangedBlock( _
    ByRef oldLines() As String, _
    ByVal oldFrom As Long, _
    ByVal oldTo As Long, _
    ByRef newLines() As String, _
    ByVal newFrom As Long, _
    ByVal newTo As Long, _
    ByRef resultText As String, _
    ByRef starts() As Long, _
    ByRef lens() As Long, _
    ByRef styles() As Long, _
    ByRef segCount As Long, _
    ByRef hasOutputLine As Boolean _
)
    Dim oldBlockCount As Long
    Dim newBlockCount As Long
    Dim score() As Double
    Dim sim() As Double
    Dim path() As Long
    Dim i As Long
    Dim j As Long

    oldBlockCount = oldTo - oldFrom + 1
    newBlockCount = newTo - newFrom + 1

    ReDim score(0 To oldBlockCount, 0 To newBlockCount)
    ReDim sim(1 To oldBlockCount, 1 To newBlockCount)
    ReDim path(0 To oldBlockCount - 1, 0 To newBlockCount - 1)

    For i = 1 To oldBlockCount
        For j = 1 To newBlockCount
            sim(i, j) = Engine_LineSimilarity( _
                oldLines(oldFrom + i - 1), _
                newLines(newFrom + j - 1) _
            )
        Next j
    Next i

    For i = oldBlockCount - 1 To 0 Step -1
        For j = newBlockCount - 1 To 0 Step -1
            Engine_SetAlignmentStep score, sim, path, i, j
        Next j
    Next i

    i = 0
    j = 0

    Do While i < oldBlockCount Or j < newBlockCount
        If i < oldBlockCount And j < newBlockCount Then
            Select Case path(i, j)
                Case ALIGN_PAIR
                    Engine_RenderWordDiffLine _
                        oldLines(oldFrom + i), _
                        newLines(newFrom + j), _
                        resultText, starts, lens, styles, segCount, hasOutputLine
                    i = i + 1
                    j = j + 1

                Case ALIGN_DELETE
                    Engine_AppendWholeLine _
                        oldLines(oldFrom + i), DIFF_DELETE, _
                        resultText, starts, lens, styles, segCount, hasOutputLine
                    i = i + 1

                Case Else
                    Engine_AppendWholeLine _
                        newLines(newFrom + j), DIFF_ADD, _
                        resultText, starts, lens, styles, segCount, hasOutputLine
                    j = j + 1
            End Select
        ElseIf i < oldBlockCount Then
            Engine_AppendWholeLine _
                oldLines(oldFrom + i), DIFF_DELETE, _
                resultText, starts, lens, styles, segCount, hasOutputLine
            i = i + 1
        Else
            Engine_AppendWholeLine _
                newLines(newFrom + j), DIFF_ADD, _
                resultText, starts, lens, styles, segCount, hasOutputLine
            j = j + 1
        End If
    Loop
End Sub

Private Sub Engine_SetAlignmentStep( _
    ByRef score() As Double, _
    ByRef sim() As Double, _
    ByRef path() As Long, _
    ByVal i As Long, _
    ByVal j As Long _
)
    Dim bestScore As Double
    Dim deleteScore As Double
    Dim addScore As Double
    Dim pairScore As Double

    deleteScore = score(i + 1, j)
    addScore = score(i, j + 1)

    bestScore = deleteScore
    path(i, j) = ALIGN_DELETE

    If addScore > bestScore + FLOAT_EPSILON Then
        bestScore = addScore
        path(i, j) = ALIGN_ADD
    End If

    If sim(i + 1, j + 1) >= LINE_SIM_THRESHOLD Then
        pairScore = sim(i + 1, j + 1) + score(i + 1, j + 1)

        If pairScore >= bestScore - FLOAT_EPSILON Then
            bestScore = pairScore
            path(i, j) = ALIGN_PAIR
        End If
    End If

    score(i, j) = bestScore
End Sub

Private Function Engine_LineSimilarity( _
    ByVal oldLine As String, _
    ByVal newLine As String _
) As Double
    Dim normalizedOld As String
    Dim normalizedNew As String
    Dim commonLength As Long

    normalizedOld = Engine_NormalizeLineForSimilarity(oldLine)
    normalizedNew = Engine_NormalizeLineForSimilarity(newLine)

    If Len(normalizedOld) = 0 And Len(normalizedNew) = 0 Then
        Engine_LineSimilarity = 1#
        Exit Function
    End If

    If Len(normalizedOld) = 0 Or Len(normalizedNew) = 0 Then Exit Function

    If Not Engine_CanUseCharLcs(Len(normalizedOld), Len(normalizedNew)) Then
        Engine_LineSimilarity = Engine_SampledLineSimilarity(normalizedOld, normalizedNew)
        Exit Function
    End If

    commonLength = Engine_LcsLength(normalizedOld, normalizedNew)
    Engine_LineSimilarity = (2# * commonLength) / (Len(normalizedOld) + Len(normalizedNew))
End Function

Private Function Engine_NormalizeLineForSimilarity(ByVal lineText As String) As String
    Engine_NormalizeLineForSimilarity = LCase$(Trim$(lineText))
End Function

Private Function Engine_SampledLineSimilarity( _
    ByVal oldLine As String, _
    ByVal newLine As String _
) As Double
    Dim oldSample As String
    Dim newSample As String
    Dim commonLength As Long

    oldSample = Engine_LineSample(oldLine)
    newSample = Engine_LineSample(newLine)

    If Len(oldSample) = 0 Or Len(newSample) = 0 Then Exit Function

    commonLength = Engine_LcsLength(oldSample, newSample)
    Engine_SampledLineSimilarity = (2# * commonLength) / (Len(oldSample) + Len(newSample))
End Function

Private Function Engine_LineSample(ByVal lineText As String) As String
    If Len(lineText) <= 500 Then
        Engine_LineSample = lineText
    Else
        Engine_LineSample = Left$(lineText, 250) & Right$(lineText, 250)
    End If
End Function

Private Function Engine_CanUseCharLcs( _
    ByVal oldLength As Long, _
    ByVal newLength As Long _
) As Boolean
    If oldLength = 0 Or newLength = 0 Then
        Engine_CanUseCharLcs = True
        Exit Function
    End If

    If oldLength > MAX_CHAR_DIFF_CHARS Or newLength > MAX_CHAR_DIFF_CHARS Then Exit Function
    If CDbl(oldLength) * CDbl(newLength) > MAX_CHAR_LCS_CELLS Then Exit Function

    Engine_CanUseCharLcs = True
End Function

Private Function Engine_LcsLength( _
    ByVal leftText As String, _
    ByVal rightText As String _
) As Long
    Dim dp() As Long
    Dim n As Long
    Dim m As Long
    Dim i As Long
    Dim j As Long

    n = Len(leftText)
    m = Len(rightText)

    If n = 0 Or m = 0 Then Exit Function

    ReDim dp(0 To n, 0 To m)

    For i = n - 1 To 0 Step -1
        For j = m - 1 To 0 Step -1
            If Mid$(leftText, i + 1, 1) = Mid$(rightText, j + 1, 1) Then
                dp(i, j) = dp(i + 1, j + 1) + 1
            ElseIf dp(i + 1, j) >= dp(i, j + 1) Then
                dp(i, j) = dp(i + 1, j)
            Else
                dp(i, j) = dp(i, j + 1)
            End If
        Next j
    Next i

    Engine_LcsLength = dp(0, 0)
End Function

Private Sub Engine_AppendWholeLine( _
    ByVal lineText As String, _
    ByVal styleType As Long, _
    ByRef resultText As String, _
    ByRef starts() As Long, _
    ByRef lens() As Long, _
    ByRef styles() As Long, _
    ByRef segCount As Long, _
    ByRef hasOutputLine As Boolean _
)
    Engine_BeginOutputLine resultText, starts, lens, styles, segCount, hasOutputLine
    AddSegment resultText, starts, lens, styles, segCount, lineText, styleType
End Sub

Private Sub Engine_RenderWordDiffLine( _
    ByVal oldLine As String, _
    ByVal newLine As String, _
    ByRef resultText As String, _
    ByRef starts() As Long, _
    ByRef lens() As Long, _
    ByRef styles() As Long, _
    ByRef segCount As Long, _
    ByRef hasOutputLine As Boolean _
)
    Dim oldTokens() As String
    Dim newTokens() As String
    Dim oldTokenCount As Long
    Dim newTokenCount As Long
    Dim dp() As Long
    Dim i As Long
    Dim j As Long
    Dim oldStart As Long
    Dim newStart As Long

    If Not Engine_CanUseCharLcs(Len(oldLine), Len(newLine)) Then
        Engine_AppendWholeLine oldLine, DIFF_DELETE, resultText, starts, lens, styles, segCount, hasOutputLine
        Engine_AppendWholeLine newLine, DIFF_ADD, resultText, starts, lens, styles, segCount, hasOutputLine
        Exit Sub
    End If

    Engine_TokenizeLine oldLine, oldTokens, oldTokenCount
    Engine_TokenizeLine newLine, newTokens, newTokenCount

    If Not Engine_CanUseWordLcs(oldTokenCount, newTokenCount) Then
        Engine_RenderCharDiffLine oldLine, newLine, resultText, starts, lens, styles, segCount, hasOutputLine
        Exit Sub
    End If

    Engine_BeginOutputLine resultText, starts, lens, styles, segCount, hasOutputLine

    If oldTokenCount = 0 And newTokenCount = 0 Then Exit Sub

    ReDim dp(0 To oldTokenCount, 0 To newTokenCount)

    For i = oldTokenCount - 1 To 0 Step -1
        For j = newTokenCount - 1 To 0 Step -1
            If oldTokens(i + 1) = newTokens(j + 1) Then
                dp(i, j) = dp(i + 1, j + 1) + 1
            ElseIf dp(i + 1, j) >= dp(i, j + 1) Then
                dp(i, j) = dp(i + 1, j)
            Else
                dp(i, j) = dp(i, j + 1)
            End If
        Next j
    Next i

    i = 0
    j = 0

    Do While i < oldTokenCount Or j < newTokenCount
        If i < oldTokenCount And j < newTokenCount Then
            If oldTokens(i + 1) = newTokens(j + 1) Then
                AddSegment resultText, starts, lens, styles, segCount, oldTokens(i + 1), DIFF_EQUAL
                i = i + 1
                j = j + 1
            Else
                oldStart = i + 1
                newStart = j + 1

                Do While i < oldTokenCount Or j < newTokenCount
                    If i < oldTokenCount And j < newTokenCount Then
                        If oldTokens(i + 1) = newTokens(j + 1) Then Exit Do

                        If dp(i + 1, j) >= dp(i, j + 1) Then
                            i = i + 1
                        Else
                            j = j + 1
                        End If
                    ElseIf i < oldTokenCount Then
                        i = i + 1
                    Else
                        j = j + 1
                    End If

                    If i < oldTokenCount And j < newTokenCount Then
                        If oldTokens(i + 1) = newTokens(j + 1) Then Exit Do
                    End If
                Loop

                Engine_RenderWordChangeBlock _
                    oldTokens, oldStart, i, _
                    newTokens, newStart, j, _
                    resultText, starts, lens, styles, segCount
            End If
        ElseIf i < oldTokenCount Then
            Engine_RenderWordChangeBlock _
                oldTokens, i + 1, oldTokenCount, _
                newTokens, j + 1, j, _
                resultText, starts, lens, styles, segCount
            i = oldTokenCount
        Else
            Engine_RenderWordChangeBlock _
                oldTokens, i + 1, i, _
                newTokens, j + 1, newTokenCount, _
                resultText, starts, lens, styles, segCount
            j = newTokenCount
        End If
    Loop
End Sub

Private Sub Engine_RenderWordChangeBlock( _
    ByRef oldTokens() As String, _
    ByVal oldFrom As Long, _
    ByVal oldTo As Long, _
    ByRef newTokens() As String, _
    ByVal newFrom As Long, _
    ByVal newTo As Long, _
    ByRef resultText As String, _
    ByRef starts() As Long, _
    ByRef lens() As Long, _
    ByRef styles() As Long, _
    ByRef segCount As Long _
)
    Dim oldText As String
    Dim newText As String

    oldText = Engine_JoinTokens(oldTokens, oldFrom, oldTo)
    newText = Engine_JoinTokens(newTokens, newFrom, newTo)

    If Len(oldText) = 0 Then
        AddSegment resultText, starts, lens, styles, segCount, newText, DIFF_ADD
    ElseIf Len(newText) = 0 Then
        AddSegment resultText, starts, lens, styles, segCount, oldText, DIFF_DELETE
    ElseIf Engine_CanUseCharLcs(Len(oldText), Len(newText)) Then
        Engine_RenderCharDiffText oldText, newText, resultText, starts, lens, styles, segCount
    Else
        AddSegment resultText, starts, lens, styles, segCount, oldText, DIFF_DELETE
        AddSegment resultText, starts, lens, styles, segCount, newText, DIFF_ADD
    End If
End Sub

Private Sub Engine_TokenizeLine( _
    ByVal lineText As String, _
    ByRef tokens() As String, _
    ByRef tokenCount As Long _
)
    Dim pos As Long
    Dim tokenStart As Long
    Dim currentIsSpace As Boolean
    Dim nextIsSpace As Boolean

    tokenCount = 0
    ReDim tokens(1 To 1)

    If Len(lineText) = 0 Then Exit Sub

    tokenStart = 1
    currentIsSpace = Engine_IsSpaceChar(Mid$(lineText, 1, 1))

    For pos = 2 To Len(lineText)
        nextIsSpace = Engine_IsSpaceChar(Mid$(lineText, pos, 1))

        If nextIsSpace <> currentIsSpace Then
            Engine_AddToken tokens, tokenCount, Mid$(lineText, tokenStart, pos - tokenStart)
            tokenStart = pos
            currentIsSpace = nextIsSpace
        End If
    Next pos

    Engine_AddToken tokens, tokenCount, Mid$(lineText, tokenStart)
End Sub

Private Function Engine_IsSpaceChar(ByVal ch As String) As Boolean
    Select Case ch
        Case " ", vbTab, ChrW$(160)
            Engine_IsSpaceChar = True
    End Select
End Function

Private Sub Engine_AddToken( _
    ByRef tokens() As String, _
    ByRef tokenCount As Long, _
    ByVal tokenText As String _
)
    tokenCount = tokenCount + 1
    ReDim Preserve tokens(1 To tokenCount)
    tokens(tokenCount) = tokenText
End Sub

Private Function Engine_JoinTokens( _
    ByRef tokens() As String, _
    ByVal tokenFrom As Long, _
    ByVal tokenTo As Long _
) As String
    Dim i As Long
    Dim joinedText As String

    If tokenFrom > tokenTo Then Exit Function

    For i = tokenFrom To tokenTo
        joinedText = joinedText & tokens(i)
    Next i

    Engine_JoinTokens = joinedText
End Function

Private Function Engine_CanUseWordLcs( _
    ByVal oldTokenCount As Long, _
    ByVal newTokenCount As Long _
) As Boolean
    If oldTokenCount = 0 Or newTokenCount = 0 Then
        Engine_CanUseWordLcs = True
        Exit Function
    End If

    If oldTokenCount > MAX_WORD_DIFF_TOKENS Or newTokenCount > MAX_WORD_DIFF_TOKENS Then Exit Function
    If CDbl(oldTokenCount) * CDbl(newTokenCount) > MAX_WORD_LCS_CELLS Then Exit Function

    Engine_CanUseWordLcs = True
End Function

Private Sub Engine_RenderCharDiffLine( _
    ByVal oldLine As String, _
    ByVal newLine As String, _
    ByRef resultText As String, _
    ByRef starts() As Long, _
    ByRef lens() As Long, _
    ByRef styles() As Long, _
    ByRef segCount As Long, _
    ByRef hasOutputLine As Boolean _
)
    If Not Engine_CanUseCharLcs(Len(oldLine), Len(newLine)) Then
        Engine_AppendWholeLine oldLine, DIFF_DELETE, resultText, starts, lens, styles, segCount, hasOutputLine
        Engine_AppendWholeLine newLine, DIFF_ADD, resultText, starts, lens, styles, segCount, hasOutputLine
        Exit Sub
    End If

    Engine_BeginOutputLine resultText, starts, lens, styles, segCount, hasOutputLine
    Engine_RenderCharDiffText oldLine, newLine, resultText, starts, lens, styles, segCount
End Sub

Private Sub Engine_RenderCharDiffText( _
    ByVal oldText As String, _
    ByVal newText As String, _
    ByRef resultText As String, _
    ByRef starts() As Long, _
    ByRef lens() As Long, _
    ByRef styles() As Long, _
    ByRef segCount As Long _
)
    Dim dp() As Long
    Dim n As Long
    Dim m As Long
    Dim i As Long
    Dim j As Long

    n = Len(oldText)
    m = Len(newText)

    If n = 0 And m = 0 Then Exit Sub

    If n = 0 Then
        AddSegment resultText, starts, lens, styles, segCount, newText, DIFF_ADD
        Exit Sub
    End If

    If m = 0 Then
        AddSegment resultText, starts, lens, styles, segCount, oldText, DIFF_DELETE
        Exit Sub
    End If

    ReDim dp(0 To n, 0 To m)

    For i = n - 1 To 0 Step -1
        For j = m - 1 To 0 Step -1
            If Mid$(oldText, i + 1, 1) = Mid$(newText, j + 1, 1) Then
                dp(i, j) = dp(i + 1, j + 1) + 1
            ElseIf dp(i + 1, j) >= dp(i, j + 1) Then
                dp(i, j) = dp(i + 1, j)
            Else
                dp(i, j) = dp(i, j + 1)
            End If
        Next j
    Next i

    i = 0
    j = 0

    Do While i < n Or j < m
        If i < n And j < m Then
            If Mid$(oldText, i + 1, 1) = Mid$(newText, j + 1, 1) Then
                AddSegment resultText, starts, lens, styles, segCount, Mid$(oldText, i + 1, 1), DIFF_EQUAL
                i = i + 1
                j = j + 1
            ElseIf dp(i + 1, j) >= dp(i, j + 1) Then
                AddSegment resultText, starts, lens, styles, segCount, Mid$(oldText, i + 1, 1), DIFF_DELETE
                i = i + 1
            Else
                AddSegment resultText, starts, lens, styles, segCount, Mid$(newText, j + 1, 1), DIFF_ADD
                j = j + 1
            End If
        ElseIf i < n Then
            AddSegment resultText, starts, lens, styles, segCount, Mid$(oldText, i + 1, 1), DIFF_DELETE
            i = i + 1
        Else
            AddSegment resultText, starts, lens, styles, segCount, Mid$(newText, j + 1, 1), DIFF_ADD
            j = j + 1
        End If
    Loop
End Sub

Private Sub Engine_BeginOutputLine( _
    ByRef resultText As String, _
    ByRef starts() As Long, _
    ByRef lens() As Long, _
    ByRef styles() As Long, _
    ByRef segCount As Long, _
    ByRef hasOutputLine As Boolean _
)
    If hasOutputLine Then
        AddSegment resultText, starts, lens, styles, segCount, vbLf, DIFF_EQUAL
    Else
        hasOutputLine = True
    End If
End Sub

Private Function Engine_BlockCount(ByVal blockFrom As Long, ByVal blockTo As Long) As Long
    If blockFrom <= blockTo Then
        Engine_BlockCount = blockTo - blockFrom + 1
    End If
End Function

' =========================================================
' Renderer
' =========================================================
Private Sub Renderer_Render( _
    ByVal outCell As Range, _
    ByVal resultText As String, _
    ByRef starts() As Long, _
    ByRef lens() As Long, _
    ByRef styles() As Long, _
    ByVal segCount As Long _
)
    Dim i As Long

    outCell.Value = resultText
    outCell.WrapText = True

    With outCell.Font
        .Bold = False
        .Strikethrough = False
        .ColorIndex = xlAutomatic
    End With

    If Len(resultText) = 0 Then Exit Sub

    For i = 1 To segCount
        If lens(i) > 0 Then
            With outCell.Characters(starts(i), lens(i)).Font
                Select Case styles(i)
                    Case DIFF_DELETE
                        .Strikethrough = True
                        .Color = vbRed
                        .Bold = False

                    Case DIFF_ADD
                        .Strikethrough = False
                        .Color = vbBlue
                        .Bold = True

                    Case Else
                        .Strikethrough = False
                        .ColorIndex = xlAutomatic
                        .Bold = False
                End Select
            End With
        End If
    Next i
End Sub

' =========================================================
' Utilities
' =========================================================
Private Function NormalizeText(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    NormalizeText = s
End Function

Private Sub AddSegment( _
    ByRef resultText As String, _
    ByRef starts() As Long, _
    ByRef lens() As Long, _
    ByRef styles() As Long, _
    ByRef segCount As Long, _
    ByVal txt As String, _
    ByVal styleType As Long _
)
    If Len(txt) = 0 Then Exit Sub

    If segCount > 0 Then
        If styles(segCount) = styleType Then
            resultText = resultText & txt
            lens(segCount) = lens(segCount) + Len(txt)
            Exit Sub
        End If
    End If

    segCount = segCount + 1

    ReDim Preserve starts(1 To segCount)
    ReDim Preserve lens(1 To segCount)
    ReDim Preserve styles(1 To segCount)

    starts(segCount) = Len(resultText) + 1
    lens(segCount) = Len(txt)
    styles(segCount) = styleType

    resultText = resultText & txt
End Sub


