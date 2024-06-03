Attribute VB_Name = "Module2"
Const FOLDER_PATH_EDIT = "C:\BrityLockNLock\4. Edit\P17_셀아웃_데이터\"
Const FOLDER_PATH_TEMPLATE = "C:\BrityLockNLock\2. Template\P17_셀아웃_데이터\"
Const FOLDER_PATH_DOWN = "C:\BrityLockNLock\3. Download\P17_셀아웃_데이터_작성_자동화\"
Const FILE_NAME_BI = "업로드 리스트_YY년 NN주차_BI 업로드.xlsx"
Const FILE_NAME_PARKING = "주차별데이터_MM월NN주차_내부정리용.xlsx"
Const SHEET_SALES_DATA = "매출데이터"

'company name
Const HOMEPLUS = "홈플러스"
Const LOTTE_MART = "롯데마트"
Const COUPANG = "쿠팡"
Const E_MART = "이마트"
'company name

Const MATERIAL_FILE = "★자재검증리스트.xlsx"
Const MATERIAL_FOLDER = "02_자재검증리스트"
 

Sub test()
    Dim wbBIFile As Workbook
    Set wbBIFile = Workbooks.Open(FOLDER_PATH_EDIT + FILE_NAME_BI)
    
    Dim wbParkingFile As Workbook
    Set wbParkingFile = Workbooks.Open(FOLDER_PATH_EDIT + FILE_NAME_PARKING)
    ActiveWindow.WindowState = xlMinimized
    wbBIFile.Activate

    'slide 32
    '3. Activate ‘매출데이터’ Sheet
    Dim wsSalesData As Worksheet
    Set wsSalesData = wbBIFile.Worksheets(SHEET_SALES_DATA)
    
    '4. Insert datas from ‘주차별데이터_MM월NN주차_내부정리용.xlsx” file
    Dim wsHomeplus As Worksheet
    Set wsHomeplus = wbParkingFile.Worksheets(HOMEPLUS)
    wsHomeplus.AutoFilterMode = False
    
    Dim lastVisibleRow As Long
    Dim lastVisibleCol As Long
    lastVisibleRow = wsHomeplus.Cells(wsHomeplus.Rows.count, 3).End(xlUp).Row
    
    'copy company data
    Dim dataCopy As Range
    Set dataCopy = wsHomeplus.Range("C3" + ":" + "C" + CStr(lastVisibleRow))
    dataCopy.Copy
    
    'paste data
    lastVisibleRow = wsSalesData.Cells(wsSalesData.Rows.count, 1).End(xlUp).Row
    Set targetCell = wsSalesData.Cells(2, 10)
    targetCell.PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    
    'copy the rest data
    lastVisibleRow = wsHomeplus.Cells(wsHomeplus.Rows.count, 7).End(xlUp).Row
    Set dataCopy = wsHomeplus.Range("G3" + ":" + "M" + CStr(lastVisibleRow))
    dataCopy.Copy
    
    'paste data
    lastVisibleRow = wsSalesData.Cells(wsSalesData.Rows.count, 3).End(xlUp).Row
    Set targetCell = wsSalesData.Cells(2, 3)
    targetCell.PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    'slide 32


    'slide 33 - 34
    Dim wbMaterial As Workbook
    Set wbMaterial = Workbooks.Open(FOLDER_PATH_TEMPLATE + MATERIAL_FOLDER + "\" + MATERIAL_FILE)
    
    
    lastVisibleRow = wsSalesData.Cells(wsSalesData.Rows.count, 3).End(xlUp).Row
    wsSalesData.Range("A2:" + "A" + CStr(lastVisibleRow)).formula = "=VLOOKUP(J2,[" + MATERIAL_FILE + "]업체코드!$B:$D,2,FALSE)"
    wsSalesData.Range("B2:" + "B" + CStr(lastVisibleRow)).formula = "=VLOOKUP(J2,[" + MATERIAL_FILE + "]업체코드!$B:$D,3,FALSE)"
    wsSalesData.Range("A2:" + "B" + CStr(lastVisibleRow)).Copy
    Set targetCell = wsSalesData.Cells(2, 1)
    targetCell.PasteSpecial Paste:=xlPasteValues
    wsSalesData.Columns("J").Delete
    wsSalesData.Activate
    
    
    
    'filter and delete text 0 for column F and column G
    wsSalesData.AutoFilterMode = False
    
    wsSalesData.Range("A1:I1").AutoFilter
    wsSalesData.Range("$A$1:$I$1").AutoFilter Field:=6, Criteria1:="0"
    
    Dim dataRange As Range
    Dim visibleCells As Range
    
     With wsSalesData.AutoFilter.Range
        ' Adjust to exclude the header row
        Set dataRange = .Resize(.Rows.count - 1).Offset(1, 0)
    End With
    
    On Error Resume Next
        Set visibleCells = dataRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleCells Is Nothing Then
        firstVisibleRow = visibleCells.Areas(1).Rows(1).Row
        lastVisibleRow = visibleCells.Areas(visibleCells.Areas.count).Rows(visibleCells.Areas(visibleCells.Areas.count).Rows.count).Row
            
        ' Select the range of visible data excluding the header
        wsSalesData.Range(wsSalesData.Cells(firstVisibleRow, dataRange.Columns(6).Column), wsSalesData.Cells(lastVisibleRow, dataRange.Columns(6).Column)).value = ""
        wsSalesData.AutoFilterMode = False
    End If
    
    
    wsSalesData.Range("A1:I1").AutoFilter
    wsSalesData.Range("$A$1:$I$1").AutoFilter Field:=7, Criteria1:="0"
    
     With wsSalesData.AutoFilter.Range
        ' Adjust to exclude the header row
        Set dataRange = .Resize(.Rows.count - 1).Offset(1, 0)
    End With
    
    On Error Resume Next
        Set visibleCells = dataRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleCells Is Nothing Then
        firstVisibleRow = visibleCells.Areas(1).Rows(1).Row
        lastVisibleRow = visibleCells.Areas(visibleCells.Areas.count).Rows(visibleCells.Areas(visibleCells.Areas.count).Rows.count).Row
        ' Select the range of visible data excluding the header
        wsSalesData.Range(wsSalesData.Cells(firstVisibleRow, dataRange.Columns(7).Column), wsSalesData.Cells(lastVisibleRow, dataRange.Columns(7).Column)).value = ""
        
        wsSalesData.AutoFilterMode = False
    End If
    'filter and delete text 0 for column F and column G
    
    'slide 33 - 34
    
    
    
    'slide 35
    wbParkingFile.Close
    Dim homeplusFile As String
    homeplusFile = HOMEPLUS + "_0528_재고.xlsx"
    
    Dim wbHomePlus As Workbook
    Set wbHomePlus = Workbooks.Open(FOLDER_PATH_DOWN + homeplusFile)
    
    'delete 'Sheet2'
    Dim tempWs As Worksheet
    For Each tempWs In wbHomePlus.Worksheets
        If tempWs.Name = "Sheet2" Then
            tempWs.Delete
            Exit For ' Exit the loop since sheet is found
        End If
    Next tempWs
    
    
    'CREATE PIVOT TABLE
    Dim wsPivotTable As Worksheet
    Dim pivotTableSheetName As String
    Dim pivotTableName As String
    Dim pivotDataRange As Range
    Dim pivotCache As pivotCache
    Dim pivot As PivotTable
    
    pivotTableSheetName = "Sheet2"
    pivotTableName = "PivotTable" + Format(DateAdd("d", -1, Date), "MMdd")
    
    lastVisibleRow = wbHomePlus.Sheets("Sheet1").Cells(wbHomePlus.Sheets("Sheet1").Rows.count, 9).End(xlUp).Row
    lastVisibleCol = wbHomePlus.Sheets("Sheet1").Cells(9, wbHomePlus.Sheets("Sheet1").Columns.count).End(xlToLeft).Column
    Application.CutCopyMode = False
    
    Set wsPivotTable = wbHomePlus.Sheets.Add
    wsPivotTable.Name = pivotTableSheetName
    
    Set pivotDataRange = wbHomePlus.Sheets("Sheet1").Range(wbHomePlus.Sheets("Sheet1").Cells(9, 1), wbHomePlus.Sheets("Sheet1").Cells(lastVisibleRow - 1, lastVisibleCol))
    Set pivotCache = wbHomePlus.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotDataRange.Address(ReferenceStyle:=xlR1C1, External:=True))
    Set pivot = wsPivotTable.PivotTables.Add(pivotCache:=pivotCache, TableDestination:=wsPivotTable.Cells(3, 1), TableName:=pivotTableName)
    
    With pivot
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    
    With wsPivotTable.PivotTables(pivotTableName).pivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    wsPivotTable.PivotTables(pivotTableName).RepeatAllLabels xlRepeatLabels
    With wsPivotTable.PivotTables(pivotTableName).PivotFields("상품(SKU)")
        .Orientation = xlRowField
        .Position = 1
    End With
    wsPivotTable.PivotTables(pivotTableName).AddDataField wsPivotTable.PivotTables(pivotTableName).PivotFields("가용재고(수량)"), "Sum of 가용재고(수량)", xlSum
    'CREATE PIVOT TABLE
    'slide 35
    
    
    
    'slide 36
    '5. Copy and Paste As Numeric to [자재데이터] Sheet
    lastVisibleRow = wbHomePlus.Sheets(pivotTableSheetName).Cells(wbHomePlus.Sheets(pivotTableSheetName).Rows.count, 1).End(xlUp).Row
    wbHomePlus.Sheets(pivotTableSheetName).Range("A4:A" + CStr(lastVisibleRow - 1)).Copy
    wbBIFile.Sheets("자재데이터").Cells(2, 4).PasteSpecial Paste:=xlPasteValues
    ' Clear the clipboard
    Application.CutCopyMode = False
    
    lastVisibleRow = wbHomePlus.Sheets(pivotTableSheetName).Cells(wbHomePlus.Sheets(pivotTableSheetName).Rows.count, 2).End(xlUp).Row
    wbHomePlus.Sheets(pivotTableSheetName).Range("B4:B" + CStr(lastVisibleRow - 1)).Copy
    wbBIFile.Sheets("자재데이터").Cells(2, 8).PasteSpecial Paste:=xlPasteValues
    ' Clear the clipboard
    Application.CutCopyMode = False
    'slide 36
    
    
    'slide 37
    '7. Insert Company Text to Column "I"
    wbBIFile.Sheets("자재데이터").Range("I2:I" + CStr(lastVisibleRow)).value = HOMEPLUS
    
    '8. Close “홈플러스_재고” file.
    wbHomePlus.Close
    
    'slide 37
End Sub
