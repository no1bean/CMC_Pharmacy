' Function to convert column number to column letter
Function ColLetter(colNum As Long) As String
    Dim vArr
    vArr = Split(Cells(1, colNum).Address(True, False), "$")
    ColLetter = vArr(0)
End Function


Sub 집계표만들기()
'
' 집계표만들기 매크로
' 첫번째 : 부분합 출력미리보기 두번째 : 호스피스용 병실 순 출력미리보기 ver231224
'
' 바로 가기 키: Ctrl+Shift+P
'
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    Dim cell As Range

    Set ws = ActiveSheet



    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row 'Find the last row with data in column G

    'Loop through each cell in column G from the last row to the first one
    For Each cell In ws.Range("G1:G" & lastRow)
        If cell.Value = "반환종료" Then
            If rng Is Nothing Then
                Set rng = cell
            Else
                Set rng = Union(rng, cell)
            End If
        End If
    Next cell
    If Not rng Is Nothing Then
         rng.EntireRow.Delete
    End If
          
          
    ' Find the cell containing "총량" in the first row
    Set foundCell = ws.Rows(1).Find(What:="총량", LookIn:=xlValues, LookAt:=xlWhole)

    ' Assuming foundCell is correctly set to the cell containing "총량"
    Dim startColToDelete As Long
    startColToDelete = foundCell.Column + 2

    ' Convert column number to letter
    Dim startColLetter As String
    startColLetter = ColLetter(startColToDelete)

    ' Delete columns from startColLetter to the end
     ws.Columns(startColLetter & ":" & ColLetter(ws.Columns.Count)).Delete
     
       
    
    ' Find columns "No", "처방일자", and "투약번호"
    Dim colNo As Long, colPrescriptionDate As Long, colMedicationNo As Long
    colNo = ws.Rows(1).Find(What:="No", LookIn:=xlValues, LookAt:=xlWhole).Column
    colPrescriptionDate = ws.Rows(1).Find(What:="처방일자", LookIn:=xlValues, LookAt:=xlWhole).Column

    ' Delete columns between "No" and "처방일자"
    If colNo < colPrescriptionDate - 1 Then
        ws.Columns(ColLetter(colNo + 1) & ":" & ColLetter(colPrescriptionDate - 1)).Delete
    End If

    colPrescriptionDate = ws.Rows(1).Find(What:="처방일자", LookIn:=xlValues, LookAt:=xlWhole).Column
    colMedicationNo = ws.Rows(1).Find(What:="투약번호", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    ' Delete columns between "처방일자" and "투약번호"
    If colPrescriptionDate < colMedicationNo - 1 Then
        ws.Columns(ColLetter(colPrescriptionDate + 1) & ":" & ColLetter(colMedicationNo - 1)).Delete
    End If
    
    
    
          
    ' Find the columns with "약품명" and "총량" in the first row
    Dim colDrugName As Long
    Dim colTotal As Long
    Dim headerCell As Range

    For Each headerCell In ws.Range("1:1")
        If headerCell.Value = "약품명" Then
            colDrugName = headerCell.Column
        ElseIf headerCell.Value = "총량" Then
            colTotal = headerCell.Column
        End If
    Next headerCell

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
        



Dim i As Long
Dim startRow As Long
Dim endRow As Long

' Assuming your data starts at row 2, change this if your data starts at a different row
startRow = 2

' Find the last row with data in the "No" column after all operations
endRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' Loop through each row and set the "No" column to the sequential number
For i = startRow To endRow
    ws.Cells(i, 1).Value = i - startRow + 1
Next i



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

    
    ' Setting up the page and print settings
    With ws.PageSetup ' Changed from ActiveSheet to ws to maintain consistency
        .Orientation = xlLandscape ' Setting orientation to landscape
        .PaperSize = xlPaperA5 ' Setting paper size to A5
        .Zoom = False ' Must set to False before FitToPagesWide or FitToPagesTall
        .FitToPagesWide = 1 ' Fitting all columns on one page wide
        .FitToPagesTall = False ' Allowing more than one page tall, if necessary
    End With

    ' Showing the print preview page with settings applied
    ws.PrintPreview ' Changed from ActiveSheet to ws to maintain consistency

    
    
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

' Assuming your data starts at row 2, change this if your data starts at a different row
startRow = 2

' Find the last row with data in the "No" column after all operations
endRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' Loop through each row and set the "No" column to the sequential number
For i = startRow To endRow
    ws.Cells(i, 1).Value = i - startRow + 1
Next i



' Setting up the page and print settings for "병실" sorting
With ws.PageSetup
    .Orientation = xlLandscape
    .PaperSize = xlPaperA5
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = False
End With

' Showing the print preview page with settings applied for "병실" sorting
ws.PrintPreview

End If

End Sub
