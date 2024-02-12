' Function to convert column number to column letter
Function ColLetter(colNum As Long) As String
    Dim vArr
    vArr = Split(Cells(1, colNum).Address(True, False), "$")
    ColLetter = vArr(0)
End Function

Sub 병동별개수세기()
'
' 병동별개수세기 매크로
' 바로 가기 키: Ctrl+Shift+V
'
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    Set ws = ActiveSheet
    ws.Rows(1).UnMerge
    Call DeleteRowsWithCriteria(ws, "반환상태", "반환종료")
    
    ' Define your custom order for "수행부서"
    Dim customOrder As Variant
    customOrder = Array("7층0병동", "7층1병동", "7층2병동", "신생아실", _
                        "9층0병동", "9층2병동", "분만실", "10층1병동", "10층2병동", "뇌졸중집중치료실", _
                        "8층0병동", "8층1병동", "8층2병동", "5층3병동", _
                        "소아중환자실", "신경계중환자실", "신생아중환자실", "심장계중환자실", "외과중환자실", "혈액계중환자실" _
                        "11층1병동", "11층2병동", "12층1병동", "12층2병동", _
                        "15층1병동", "15층2병동", _
                        "13층1병동", "13층2병동", "14층1병동", "14층2병동", _
                        "16층1병동", "16층2병동", _
                        "17층1병동", "17층2병동", "18층1병동", "18층2병동", _
                        "19층1병동", "19층2병동", "내과중환자실", _
                        "20층2병동", "21층1병동", "21층2병동")

    ' Create a PivotTable
    Dim pt As PivotTable
    Dim ptCache As PivotCache
    Dim ptSheet As Worksheet
    Dim ptRange As Range

    ' Create a new worksheet for PivotTable in the active workbook
    Set ptSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ptSheet.Activate

    ' Define the range for PivotTable
    Dim ptLastRow As Long
    Dim ptLastCol As Long
    ptLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ptLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set ptRange = ws.Range(ws.Cells(1, 1), ws.Cells(ptLastRow, ptLastCol))

    ' Create PivotCache and PivotTable in the active workbook
    Set ptCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ptRange)
    Set pt = ptCache.CreatePivotTable(TableDestination:=ptSheet.Cells(1, 1), TableName:="MyPivotTable")

    ' Set up the PivotTable
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

        ' Set the current page to today's date if it exists
        If itemExists Then
            .PivotFields("처방일자").CurrentPage = todayDate
            End If
   End With
   pt.PivotFields("SortOrder").Orientation = xlHidden     
        
   ' Setting up the page and print settings for the PivotTable sheet
   With ptSheet.PageSetup
       .Orientation = xlPortrait
       .PaperSize = xlPaperA5
       .Zoom = False
       .FitToPagesWide = 1
       .FitToPagesTall = 1
   End With

   ' Showing the print preview page with settings applied for the PivotTable sheet
   ptSheet.PrintPreview

End Sub

Sub DeleteRowsWithCriteria(ws As worksheet, headerName As String, criteria As String)
    Dim columnNumber As Long
    Dim columnLetter As String
    
    columnNumber = FindColumn(ws, headerName)
    If columnNumber <> 0 Then
        columnLetter = ColLetter(columnNumber)
        With ws.Columns(columnLetter & ":" & columnLetter)
            .AutoFilter Field:=1, Criteria1:=criteria
            If .Parent.AutoFilterMode Then
                .Parent.AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End If
            .Parent.AutoFilterMode = False
        End With
    End If
End Sub

Function ColLetter(colNum As Long) As String
    ColLetter = Replace(Cells(1, colNum).Address(False, False), 1, "")
End Function

Function FindColumn(ws As worksheet, ByVal headerName As String) As Long
    Dim foundCell As Range
    Set foundCell = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
    If Not foundCell Is Nothing Then
        FindColumn = foundCell.Column
    Else
        FindColumn = 0 ' Return 0 or handle error if header is not found
    End If
End Function