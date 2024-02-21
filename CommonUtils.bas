Sub ClearExcessRowsAndColumns(ws As Worksheet)
    Dim lastUsedRow As Long, lastUsedCol As Long
    Dim lastShapeRow As Long, lastShapeCol As Long
    Dim usedRange As Range, areaRange As Range
    Dim shape As Shape

    If ActiveWorkbook Is Nothing Then Exit Sub

    If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then ws.Unprotect ""

    On Error Resume Next
    Set usedRange = ws.UsedRange
    Set areaRange = Union(usedRange.SpecialCells(xlCellTypeConstants), usedRange.SpecialCells(xlCellTypeFormulas))
    If Not areaRange Is Nothing Then
        lastUsedRow = areaRange.Rows(areaRange.Rows.Count).Row
        lastUsedCol = areaRange.Columns(areaRange.Columns.Count).Column
    End If

    ' If lastUsedRow and lastUsedCol are still 0, set them to the last row and column of usedRange
    If lastUsedRow = 0 Then lastUsedRow = usedRange.Rows(usedRange.Rows.Count).Row
    If lastUsedCol = 0 Then lastUsedCol = usedRange.Columns(usedRange.Columns.Count).Column

    For Each shape In ws.Shapes
        lastShapeRow = shape.BottomRightCell.Row
        lastShapeCol = shape.BottomRightCell.Column
        If lastShapeCol > lastUsedCol Then lastUsedCol = lastShapeCol
        If lastShapeRow > lastUsedRow Then lastUsedRow = lastShapeRow
    Next shape

    If lastUsedRow < ws.Rows.Count Then
        With ws.Rows(lastUsedRow + 1 & ":" & ws.Rows.Count)
            .Hidden = False
            .Clear
        End With
    End If

    If lastUsedCol < ws.Columns.Count Then
        With ws.Columns(lastUsedCol + 1 & ":" & ws.Columns.Count)
            .Hidden = False
            .Clear
        End With
    End If

End Sub

Function CopySheet(sourceSheet As Worksheet, nameSuffix As String) As Worksheet
    Dim newSheet As Worksheet
    Dim baseName As String
    Dim newSheetName As String
    Dim counter As Integer
    
    sourceSheet.Copy After:=Sheets(Sheets.Count)
    Set newSheet = ActiveSheet
    
    baseName = sourceSheet.Name & nameSuffix
    newSheetName = baseName

    counter = 1
    While SheetExists(newSheetName)
        newSheetName = baseName & counter
        counter = counter + 1
    Wend
    
    newSheet.Name = newSheetName
    Set CopySheet = newSheet
End Function

Function ColumnLetter(colNum As Long) As String
    Dim result As String
    Dim modCol As Integer
    
    result = ""
    While colNum > 0
        modCol = (colNum - 1) Mod 26
        result = Chr(65 + modCol) & result
        colNum = (colNum - modCol) \ 26
    Wend
    
    ColumnLetter = result
End Function

Sub DeleteRowsWithCriteria(ws As Worksheet, headerName As String, criteria As String)
    Dim columnNumber As Long
    Dim columnLetter As String
    
    columnNumber = FindColumn(ws, headerName)
    If columnNumber <> 0 Then
        columnLetter = ColumnLetter(columnNumber)
        With ws.Columns(columnLetter & ":" & columnLetter)
            .AutoFilter Field:=1, Criteria1:=criteria
            If .Parent.AutoFilterMode Then
                .Parent.AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End If
            .Parent.AutoFilterMode = False
        End With
    End If
End Sub

Function FindColumn(ws As Worksheet, headerName As String) As Long
    Dim foundCell As Range
    Set foundCell = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
    If Not foundCell Is Nothing Then
        FindColumn = foundCell.Column
    Else
        FindColumn = 0 ' Return 0 or handle error if header is not found
    End If
End Function

Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    Dim sheet As Object
    Set sheet = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not sheet Is Nothing
    On Error GoTo 0
End Function

Sub ShowPrintPreview(ws As Worksheet, Optional isPortrait As Boolean = False, Optional fitToOnePage As Boolean = False)
    With ws.PageSetup
        .Orientation = IIf(isPortrait, xlPortrait, xlLandscape)
        .PaperSize = xlPaperA5
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = IIf(fitToOnePage, 1, False)
    End With
    ws.PrintPreview
End Sub