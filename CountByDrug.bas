Sub 집계표만들기()
'
' 집계표만들기 매크로
' 첫번째 : 부분합 출력미리보기 
' 두번째 : 호스피스용 병실 순 출력미리보기
' 바로 가기 키: Ctrl+Shift+P
'
    Dim ws As worksheet
    Set ws = CopySheet(ActiveSheet,"-집계표")

    Call CommonUtils.ClearExcessRowsAndColumns(ws)

    With ws
        .Cells.Font.Name = "Dotum"
        .Cells.Font.Size = 9
        With .UsedRange.Borders
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
        .UsedRange.Rows.RowHeight=23.2
    End With

    Call CommonUtils.DeleteRowsWithCriteria(ws, "반환상태", "반환종료")
    
    ' Define the desired order of the columns
    Dim arrColumnOrder As Variant
    arrColumnOrder = Array("No", "처방일자", "투약번호", "처방구분", "수행부서", "병실", _
                            "환자번호", "환자명", "연령", "약픔코드", "약품명", "총량")
    
    Dim dictMergeInfo As Object
    Set dictMergeInfo = CreateObject("Scripting.Dictionary")
    Dim i As Integer
    Dim colNum As Long
    Dim rngCell As Range

    ' Store merge information for headers of interest and unmerge those cells
    For i = LBound(arrColumnOrder) To UBound(arrColumnOrder)
        Set rngCell = ws.Rows(1).Find(What:=arrColumnOrder(i), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not rngCell Is Nothing Then
            colNum = rngCell.Column
            If rngCell.MergeCells Then
                ' Store the number of merged columns
                dictMergeInfo(colNum) = rngCell.MergeArea.Columns.Count
                ' Unmerge the cells
                rngCell.MergeArea.UnMerge
            Else
                ' If not merged, just store 1
                dictMergeInfo(colNum) = 1
            End If
        End If
    Next i
    ' Rearrange the columns based on arrColumnOrder
    Dim targetColNum As Long
    Dim currentColNum As Long
    Dim colRange As Range

    targetColNum = 1 ' Start from the first column

    For i = LBound(arrColumnOrder) To UBound(arrColumnOrder)
        currentColNum = CommonUtils.FindColumn(ws, arrColumnOrder(i)) ' Find the current position of the column

        If currentColNum > 0 And currentColNum <> targetColNum Then
            ' Handle merged columns
            Dim mergeSpan As Long
            If dictMergeInfo.Exists(currentColNum) Then
                mergeSpan = dictMergeInfo(currentColNum)
            Else
                mergeSpan = 1
            End If

            ' Select and cut the entire column or columns (in case of a merged header)
            Set colRange = ws.Columns(currentColNum).Resize(, mergeSpan)
            colRange.Cut
            ws.Columns(targetColNum).Insert Shift:=xlToRight
            Application.CutCopyMode = False ' Clear the clipboard

            ' Update the target position
            targetColNum = targetColNum + mergeSpan
        ElseIf currentColNum > 0 Then
            ' Update the target position if column is already in the correct place
            targetColNum = targetColNum + dictMergeInfo(currentColNum)
        End If
    Next i

    ' Reapply the merges based on the original merge information
    Dim header As Variant
    For Each header In arrColumnOrder
        colNum = CommonUtils.FindColumn(ws, header) ' Get the new column number after rearrangement
        If dictMergeInfo.Exists(colNum) Then
            mergeSpan = dictMergeInfo(colNum)
            If mergeSpan > 1 Then
                ws.Range(ws.Cells(1, colNum), ws.Cells(1, colNum + mergeSpan - 1)).Merge
            End If
        End If
    Next header
        
    Dim wardColumnNumber as Long
    Dim hospiceCount As Long

    wardColumnNumber = CommonUtils.FindColumn(ws,"수행부서")
    hospiceCount = Application.WorksheetFunction.CountIf( _
               ws.Range(ws.Cells(2, wardColumnNumber), _
               ws.Cells(ws.UsedRange.Rows.Count, wardColumnNumber)), _
               "호스피스완화의료병동")

    If hospiceCount = ws.UsedRange.Rows.Count - 1 Then
        Dim roomOrderSheet As Worksheet
        Set roomOrderSheet = CopySheet(ws,"-병실순")

        With roomOrderSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=roomOrderSheet.Cells(1, CommonUtils.FindColumn(roomOrderSheet, "병실")), Order:=xlAscending
            .SetRange roomOrderSheet.UsedRange
            .Header = xlYes
            .Apply
        End With

        Call UpdateColumnNumber(roomOrderSheet)
        MsgBox "먼저 호스피스완화의료병동 병실순 출력화면입니다.", vbExclamation
        Call ShowPrintPreview(roomOrderSheet)
    ElseIf hospiceCount > 0 And hospiceCount < ws.UsedRange.Rows.Count - 1 Then
        MsgBox "수행부서에 호스피스완화의료병동과 다른 부서가 섞여있습니다. 병동순 출력은 되지 않습니다.", vbExclamation
    End If

    Call UpdateColumnNumber(ws)
    ' Apply subtotal
    ws.UsedRange.Subtotal GroupBy:=colSubtotalGroupBy, Function:=xlSum, TotalList:=Array(colTotal), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    MsgBox "약물별 집계표 출력화면입니다.", vbExclamation
    Call CommonUtils.ShowPrintPreview(ws)

End Sub

Sub UpdateColumnNumber(ws As Worksheet)
    Dim startRow As Long
    startRow = 2

    With ws
        Dim dataRange As Range
        Dim dataArray() As Variant
        Dim i As Long

        Set dataRange = .Range("A" & startRow & ":A" & .Cells(.Rows.Count, "A").End(xlUp).Row)
        dataArray = dataRange.Value

        For i = LBound(dataArray, 1) To UBound(dataArray, 1)
            dataArray(i, 1) = i + 1
        Next i

        dataRange.Value = dataArray
    End With
End Sub