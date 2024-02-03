Function ColLetter(colNum As Long) As String
    ColLetter = Replace(Cells(1, colNum).Address(False, False), 1, "")
End Function

Function FindColumn(ws As Worksheet, headerName As String) As Long
    Dim foundCell As Range
    Set foundCell = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
    If Not foundCell Is Nothing Then
        FindColumn = foundCell.Column
    Else
        FindColumn = 0 ' Return 0 or handle error if header is not found
    End If
End Function

Sub DeleteRowsWithCriteria(targetSheet As Worksheet, headerName As String, criteria As String)
    Dim columnNumber As Long
    Dim columnLetter As String
    
    columnNumber = FindColumn(targetSheet, headerName)
    If columnNumber <> 0 Then
        columnLetter = ColLetter(columnNumber)
        With targetSheet.Columns(columnLetter & ":" & columnLetter)
            .AutoFilter Field:=1, Criteria1:=criteria
            If .Parent.AutoFilterMode Then
                .Parent.AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End If
            .Parent.AutoFilterMode = False
        End With
    End If
End Sub

Sub UpdateNoColumn(ByRef worksheet As Worksheet)
    Dim rowCounter As Long
    Dim endRow As Long
    Dim startRow As Long
    startRow = 2
    
    With worksheet
        endRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        Dim dataArray() As Variant
        dataArray = .Range(.Cells(startRow, 1), .Cells(endRow, 1)).Value
        For rowCounter = LBound(dataArray, 1) To UBound(dataArray, 1)
            dataArray(rowCounter, 1) = rowCounter - startRow + 1
        Next rowCounter
        .Range(.Cells(startRow, 1), .Cells(endRow, 1)).Value = dataArray
    End With
End Sub

Sub ConfigurePageSettings(ws As Worksheet)
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA5
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    ws.PrintPreview
End Sub

Sub 집계표만들기()
'
' 집계표만들기 매크로
' 첫번째 : 부분합 출력미리보기 두번째 : 호스피스용 병실 순 출력미리보기 ver240203
'
' 바로 가기 키: Ctrl+Shift+P
'
    Dim ws As Worksheet
    Dim colTotal As Long

    Set ws = ActiveSheet

    Call DeleteRowsWithCriteria(ws, "반환상태", "반환종료")
          
    colTotal = FindColumn(ws, "총량")

    Dim startColToDelete As Long
    startColToDelete = colTotal + 2

    Dim startColLetter As String
    startColLetter = ColLetter(startColToDelete)

    ' Delete columns from startColLetter to the end
    ws.Columns(startColLetter & ":" & ColLetter(ws.UsedRange.Columns.Count)).Delete
     
    ' Find columns "No", "처방일자", and "투약번호"
    Dim colNo As Long, colPrescriptionDate As Long, colMedicationNo As Long
    colNo = FindColumn(ws, "No")
    colPrescriptionDate = FindColumn(ws, "처방일자")

    ' Delete columns between "No" and "처방일자"
    If colNo < colPrescriptionDate - 1 Then
        ws.Columns(ColLetter(colNo + 1) & ":" & ColLetter(colPrescriptionDate - 1)).Delete
        colPrescriptionDate = FindColumn(ws, "처방일자")
    End If

    ' Delete columns between "처방일자" and "투약번호"
    colMedicationNo = FindColumn(ws, "투약번호")
    If colPrescriptionDate < colMedicationNo - 1 Then
        ws.Columns(ColLetter(colPrescriptionDate + 1) & ":" & ColLetter(colMedicationNo - 1)).Delete
    End If
          
    ' Find the columns with "약품명" and "총량" in the first row
    Dim colDrugName As Long
    colDrugName = FindColumn(ws, "약품명")

    colTotal = FindColumn(ws, "총량")

    ' Sort the data based on the found columns
    With ws.Sort
        .SortFields.Clear
        ' Adding "약품명" column to sort (ascending)
        .SortFields.Add Key:=Cells(1, colDrugName), Order:=xlAscending
        ' Adding "총량" column to sort (descending)
        .SortFields.Add Key:=Cells(1, colTotal), Order:=xlDescending

        .SetRange ws.UsedRange ' Setting the range to the used range of the worksheet
        .Header = xlYes ' The first row contains headers
        .Apply ' Applying the sort
    End With
        
    Call UpdateNoColumn(ws)

    ' Find the column with "약품명" after all operations that might shift columns
    Dim colSubtotalGroupBy As Long
    Dim finalHeaderCell As Range
    Dim lastColumnAfterDeletions As Long

    lastColumnAfterDeletions = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column ' Last used column after deletions

    For Each finalHeaderCell In ws.Range(ws.Cells(1, 1), ws.Cells(1, lastColumnAfterDeletions))
        If finalHeaderCell.Value = "약품명" Then
            colSubtotalGroupBy = finalHeaderCell.Column
            Exit For
        End If
    Next finalHeaderCell

    ' Check if "약품명" column is found, and if not, raise an error or handle it appropriately
    If colSubtotalGroupBy = 0 Then
        MsgBox """약품명"" column not found. Cannot apply subtotal.", vbExclamation
        Exit Sub ' Or handle the error appropriately
    End If

    ' Apply subtotal
    ws.UsedRange.Subtotal GroupBy:=colSubtotalGroupBy, Function:=xlSum, TotalList:=Array(colTotal), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True

    ' Showing the print preview page with settings applied
    Call ConfigurePageSettings(ws)
        
    Dim colDepartment As Long
    Dim departmentValue As String

    ' Find the column number for "수행부서"
    For Each headerCell In ws.Range("1:1")
        If headerCell.Value = "수행부서" Then
            colDepartment = headerCell.Column
            Exit For
        End If
    Next headerCell

    ' Check the value in the second row of the "수행부서" column
    departmentValue = ws.Cells(2, colDepartment).Value

    If departmentValue = "호스피스완화의료병동" Then

    ' Remove any existing subtotals
    Dim dataRange As Range
    Set dataRange = ws.UsedRange
    dataRange.RemoveSubtotal

    ' Sort by the "병실" column
    Dim colRoom As Long
    For Each headerCell In ws.Range("1:1")
        If headerCell.Value = "병실" Then
            colRoom = headerCell.Column
            Exit For
        End If
    Next headerCell

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Cells(1, colRoom), Order:=xlAscending
        .SetRange ws.UsedRange
        .Header = xlYes
        .Apply
    End With

    Call UpdateNoColumn(ws)

    ' Showing the print preview page with settings applied for "병실" sorting
    Call ConfigurePageSettings(ws)

    End If

End Sub
