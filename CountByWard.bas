Sub 병동별개수세기()
'
' 병동별개수세기 매크로
' 바로 가기 키: Ctrl+Shift+V
'
    Dim ws As Worksheet
    Set ws = CopySheet(ActiveSheet,"-Copy")

    Call ClearExcessRowsAndColumns(ws)
    
    Dim foundCell As Range
    Set foundCell = ws.Rows(1).Find(What:="총량", LookIn:=xlValues, LookAt:=xlWhole)

    ws.Columns(ColumnLetter(foundCell.Column + 1) & ":" & ColumnLetter(ws.Columns.Count)).Delete
     
    Call DeleteRowsWithCriteria(ws, "반환상태", "반환종료")
    
    Dim customOrder As Variant
    customOrder = Array("7층0병동", "7층1병동", "7층2병동", "신생아실", _
                        "9층0병동", "9층2병동", "분만실", "10층1병동", "10층2병동", "뇌졸중집중치료실", _
                        "8층0병동", "8층1병동", "8층2병동", "5층3병동", _
                        "소아중환자실", "신경계중환자실", "신생아중환자실", "심장계중환자실", "외과중환자실", "혈액계중환자실", _
                        "11층1병동", "11층2병동", "12층1병동", "12층2병동", _
                        "15층1병동", "15층2병동", _
                        "13층1병동", "13층2병동", "14층1병동", "14층2병동", _
                        "16층1병동", "16층2병동", _
                        "17층1병동", "17층2병동", "18층1병동", "18층2병동", _
                        "19층1병동", "19층2병동", "내과중환자실", _
                        "20층2병동", "21층1병동", "21층2병동")

    Call CreateSortOrderColumn(ws, "수행부서", customOrder)

    Dim ptSheet As Worksheet
    Set ptSheet = CopySheet(ActiveSheet,"-Pivot")

    Call CommonUtils.SetupAndDisplayPivotTable(ws, ptSheet)
    Call CommonUtils.ShowPrintPreview(ptSheet, True, True)

End Sub

Sub CreateSortOrderColumn(ws As Worksheet, headerName As String, customOrder As Variant)
    Dim headerColumn As Long
    headerColumn = FindColumn(ws, headerName)
    If headerColumn = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, headerColumn).End(xlUp).Row

    Dim dataColumn As Variant
    dataColumn = ws.Range(ws.Cells(2, headerColumn), ws.Cells(lastRow, headerColumn)).Value

    Dim sortOrderArray() As Variant
    ReDim sortOrderArray(1 To UBound(dataColumn, 1), 1 To 1)

    Dim i As Long, orderIndex As Variant
    For i = 1 To UBound(dataColumn, 1)
        orderIndex = Application.Match(dataColumn(i, 1), customOrder, False)
        sortOrderArray(i, 1) = IIf(Not IsError(orderIndex), orderIndex, 1000)
    Next i

    Dim sortOrderColumn As Long
    sortOrderColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    ws.Cells(1, sortOrderColumn).Value = "SortOrder"
    ws.Range(ws.Cells(2, sortOrderColumn), ws.Cells(lastRow, sortOrderColumn)).Value = sortOrderArray
End Sub

Sub SetupAndDisplayPivotTable(ws As Worksheet, ptSheet As Worksheet)
    ptSheet.Activate
    
    Dim ptLastRow As Long
    ptLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim ptLastCol As Long
    ptLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim ptRange As Range
    Set ptRange = ws.Range(ws.Cells(1, 1), ws.Cells(ptLastRow, ptLastCol))

    Dim ptCache As PivotCache
    Set ptCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ptRange)

    Dim pt As PivotTable
    Set pt = ptCache.CreatePivotTable(TableDestination:=ptSheet.Cells(1, 1), TableName:="CountByWard")

    With pt
        .PivotFields("SortOrder").Orientation = xlRowField
        .PivotFields("SortOrder").Position = 1
        .PivotFields("수행부서").Orientation = xlRowField
        .PivotFields("수행부서").Position = 2
        .PivotFields("약품코드").Orientation = xlDataField
        .PivotFields("처방일자").Orientation = xlPageField

        ' Check if today's date exists in the "처방일자" field
        Dim todayDate As String
        todayDate = Format(Date, "yyyy-mm-dd")
        Dim itemExists As Boolean
        itemExists = False

        Dim pItem As PivotItem
        On Error Resume Next ' Ignore error if item not found
        Set pItem = .PivotFields("처방일자").PivotItems(todayDate)
        If Not pItem Is Nothing Then
            itemExists = True
        End If
        On Error GoTo 0 ' Reset error handling

        If itemExists Then
            .PivotFields("처방일자").CurrentPage = todayDate
        End If
        
        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowHeaders = False
   End With
    
    Dim pf As PivotField
    For Each pf In pt.PivotFields
        pf.Subtotals(1) = False
    Next pf
    Dim totalRow As Long
    totalRow = 3
    Do While ptSheet.Cells(totalRow, 1).Value <> "총합계" And totalRow <= ptSheet.Cells(ptSheet.Rows.Count, 1).End(xlUp).Row
        totalRow = totalRow + 1
    Loop
    If ptSheet.Cells(totalRow, 1).Value = "총합계" Then
        ptSheet.Range("A3:A" & totalRow - 1).Font.Color = RGB(255, 255, 255)
    End If
End Sub
