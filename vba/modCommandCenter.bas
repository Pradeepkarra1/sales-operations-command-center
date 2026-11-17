Attribute VB_Name = "modCommandCenter"
Option Explicit

' =====================================================================
'  Sales & Operations Command Center - automation macros
'  Run any Public Sub from Alt+F8, or use the buttons on Control_Panel.
'  Start with BuildPivotReports (it also builds the Control_Panel sheet).
' =====================================================================

Private Const SRC_SHEET As String = "Sales_Data"
Private Const SRC_TABLE As String = "tblSales"
Private Const PANEL As String = "Control_Panel"
Private Const APP_TITLE As String = "Command Center"

' Runs when the workbook is opened with macros enabled.
Public Sub Auto_Open()
    If Not SheetExists("Pivot_Sales") Then
        If MsgBox("Build the pivot tables, slicers and control panel now?" & vbCrLf & _
                  "You can also run BuildPivotReports later from Alt+F8.", _
                  vbYesNo + vbQuestion, APP_TITLE) = vbYes Then
            BuildPivotReports
        End If
    End If
End Sub

' ---------------------------------------------------------------------
'  1. Pivot tables, pivot charts and slicers
' ---------------------------------------------------------------------
Public Sub BuildPivotReports()
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim ws As Worksheet
    Dim pts As Collection
    Dim sheetList As Variant
    Dim i As Long

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Remove earlier versions so the macro can be re-run safely
    For i = ThisWorkbook.SlicerCaches.Count To 1 Step -1
        ThisWorkbook.SlicerCaches(i).Delete
    Next i
    sheetList = Array("Pivot_Sales", "Pivot_Dealers", "Pivot_Products", _
                      "Pivot_Monthly", "Pivot_Customers", "Pivot_Delivery")
    For i = LBound(sheetList) To UBound(sheetList)
        DeleteSheet CStr(sheetList(i))
    Next i

    ' One shared pivot cache on the sales table: smaller file, shared slicers
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SRC_TABLE)
    Set pts = New Collection

    ' --- Pivot_Sales: segment and region by fiscal year
    Set ws = NewReportSheet("Pivot_Sales", "Net sales by segment and by region")
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A5"), TableName:="ptSegment")
    SetField pt, "Segment", xlRowField
    SetField pt, "Fiscal Year", xlColumnField
    AddValue pt, "Net Sales", "Net Sales ($)", xlSum, "$#,##0"
    StylePivot pt
    pts.Add pt
    AddPivotChart ws, pt, "J14", xlColumnClustered, "Net sales by segment and fiscal year"

    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A16"), TableName:="ptRegion")
    SetField pt, "Region", xlRowField
    SetField pt, "Fiscal Year", xlColumnField
    AddValue pt, "Net Sales", "Net Sales ($)", xlSum, "$#,##0"
    StylePivot pt
    pts.Add pt

    ' --- Pivot_Dealers: dealer scorecard
    Set ws = NewReportSheet("Pivot_Dealers", "Dealer scorecard")
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A5"), TableName:="ptDealers")
    SetField pt, "Region", xlRowField, 1
    SetField pt, "Dealer ID", xlRowField, 2
    AddValue pt, "Order ID", "Orders", xlCount, "#,##0"
    AddValue pt, "Net Sales", "Net Sales ($)", xlSum, "$#,##0"
    AddValue pt, "Gross Margin", "Gross Margin ($)", xlSum, "$#,##0"
    AddValue pt, "Discount %", "Avg Discount", xlAverage, "0.0%"
    AddValue pt, "On Time", "On-Time Rate", xlAverage, "0.0%"
    StylePivot pt
    pts.Add pt

    ' --- Pivot_Products: product hierarchy
    Set ws = NewReportSheet("Pivot_Products", "Product performance")
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A5"), TableName:="ptProducts")
    SetField pt, "Segment", xlRowField, 1
    SetField pt, "Category", xlRowField, 2
    SetField pt, "Product ID", xlRowField, 3
    AddValue pt, "Units", "Units Sold", xlSum, "#,##0"
    AddValue pt, "Net Sales", "Net Sales ($)", xlSum, "$#,##0"
    AddValue pt, "Gross Margin", "Gross Margin ($)", xlSum, "$#,##0"
    StylePivot pt
    pts.Add pt

    ' --- Pivot_Monthly: monthly trend
    Set ws = NewReportSheet("Pivot_Monthly", "Monthly net sales trend")
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A5"), TableName:="ptMonthly")
    SetField pt, "Month", xlRowField
    SetField pt, "Segment", xlColumnField
    AddValue pt, "Net Sales", "Net Sales ($)", xlSum, "$#,##0"
    StylePivot pt
    pts.Add pt
    AddPivotChart ws, pt, "J14", xlLine, "Monthly net sales by segment"

    ' --- Pivot_Customers: customer mix
    Set ws = NewReportSheet("Pivot_Customers", "Customer type mix")
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A5"), TableName:="ptCustomers")
    SetField pt, "Customer Type", xlRowField
    AddValue pt, "Order ID", "Orders", xlCount, "#,##0"
    AddValue pt, "Net Sales", "Net Sales ($)", xlSum, "$#,##0"
    AddValue pt, "Gross Margin", "Gross Margin ($)", xlSum, "$#,##0"
    AddValue pt, "Discount %", "Avg Discount", xlAverage, "0.0%"
    StylePivot pt
    pts.Add pt
    AddPivotChart ws, pt, "J14", xlBarClustered, "Customer type mix"

    ' --- Pivot_Delivery: on-time delivery
    Set ws = NewReportSheet("Pivot_Delivery", "On-time delivery rate")
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A5"), TableName:="ptDelivery")
    SetField pt, "Region", xlRowField
    SetField pt, "Segment", xlColumnField
    AddValue pt, "On Time", "On-Time Rate", xlAverage, "0.0%"
    StylePivot pt
    pts.Add pt

    AddSharedSlicers pts, sheetList
    BuildControlPanel

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Built " & pts.Count & " pivot tables on 6 sheets, with 4 slicers on each sheet." & vbCrLf & _
           "A slicer click filters every pivot report in the workbook.", vbInformation, APP_TITLE
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "BuildPivotReports stopped: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub AddSharedSlicers(pts As Collection, sheetList As Variant)
    Dim fieldList As Variant
    Dim f As Long, p As Long, s As Long
    Dim sc As SlicerCache
    Dim firstPt As PivotTable
    Dim ws As Worksheet

    fieldList = Array("Fiscal Year", "Region", "Segment", "Customer Type")
    Set firstPt = pts(1)
    For f = LBound(fieldList) To UBound(fieldList)
        Set sc = ThisWorkbook.SlicerCaches.Add2(firstPt, CStr(fieldList(f)))
        ' Connect the same slicer cache to every pivot table
        For p = 2 To pts.Count
            sc.PivotTables.AddPivotTable pts(p)
        Next p
        ' Show the slicer on every report sheet
        For s = LBound(sheetList) To UBound(sheetList)
            Set ws = ThisWorkbook.Worksheets(CStr(sheetList(s)))
            sc.Slicers.Add SlicerDestination:=ws, _
                           Name:="slc" & f & "_" & CStr(sheetList(s)), _
                           Caption:=CStr(fieldList(f)), _
                           Top:=ws.Range("J2").Top, _
                           Left:=ws.Range("J2").Left + f * 150, _
                           Width:=140, Height:=165
        Next s
    Next f
End Sub

Private Sub SetField(pt As PivotTable, fieldName As String, orient As XlPivotFieldOrientation, _
                     Optional pos As Long = 0)
    With pt.PivotFields(fieldName)
        .Orientation = orient
        If pos > 0 Then .Position = pos
    End With
End Sub

Private Sub AddValue(pt As PivotTable, fieldName As String, label As String, _
                     fn As XlConsolidationFunction, fmt As String)
    Dim df As PivotField
    Set df = pt.AddDataField(pt.PivotFields(fieldName), label, fn)
    df.NumberFormat = fmt
End Sub

Private Sub StylePivot(pt As PivotTable)
    On Error Resume Next
    pt.TableStyle2 = "PivotStyleMedium2"
    pt.ShowTableStyleRowStripes = True
    On Error GoTo 0
End Sub

Private Sub AddPivotChart(ws As Worksheet, pt As PivotTable, anchor As String, _
                          kind As XlChartType, titleText As String)
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(ws.Range(anchor).Left, ws.Range(anchor).Top, 520, 270)
    With co.Chart
        .SetSourceData pt.TableRange1
        .ChartType = kind
        .HasTitle = True
        .ChartTitle.Text = titleText
    End With
End Sub

Private Function NewReportSheet(sheetName As String, heading As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    ws.Range("A1").Value = heading
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Built by the BuildPivotReports macro from table " & SRC_TABLE & _
                           ". Slicers filter every pivot report in the workbook."
    ws.Range("A2").Font.Italic = True
    Set NewReportSheet = ws
End Function

' ---------------------------------------------------------------------
'  2. Control panel with one button per macro
' ---------------------------------------------------------------------
Public Sub BuildControlPanel()
    Dim ws As Worksheet
    Dim items As Variant
    Dim i As Long, r As Long
    Dim btn As Object
    Dim oldAlerts As Boolean

    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    DeleteSheet PANEL
    Application.DisplayAlerts = oldAlerts

    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = PANEL
    ws.Range("B1").Value = "Control Panel"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 16
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D").ColumnWidth = 110

    items = Array( _
        Array("RefreshAllReports", "Refresh all reports", "Refresh every pivot table and recalculate after new orders are pasted into Sales_Data."), _
        Array("BuildPivotReports", "Rebuild pivots and slicers", "Delete and rebuild the six pivot sheets, pivot charts and shared slicers."), _
        Array("ResetAllFilters", "Reset all filters", "Clear every slicer and set the Dashboard back to FY2026 / All regions."), _
        Array("ExportDashboardPDF", "Export dashboard to PDF", "Save the Dashboard and Findings sheets as a dated PDF next to this workbook."), _
        Array("CreateDealerReports", "Create dealer reports", "Create one workbook per dealer (orders plus KPI summary) in a Dealer_Reports folder."), _
        Array("RunDataQualityCheck", "Run data quality check", "Scan Sales_Data for duplicates, blanks, formula errors and out-of-range values; log them."), _
        Array("DiscountExceptionReport", "Discount exception report", "List every order at or above a discount level you choose, sorted by discount."), _
        Array("EmailKpiSummary", "Draft KPI email", "Open an Outlook draft with the headline KPIs and top findings (nothing is sent)."))

    For i = LBound(items) To UBound(items)
        r = 4 + i * 3
        Set btn = ws.Buttons.Add(ws.Range("B" & r).Left, ws.Range("B" & r).Top, 190, 28)
        btn.OnAction = CStr(items(i)(0))
        btn.Caption = CStr(items(i)(1))
        ws.Range("D" & r).Value = CStr(items(i)(2))
    Next i
    ws.Range("B2").Value = "Last refreshed: not yet"
    ws.Activate
End Sub

' ---------------------------------------------------------------------
'  3. Refresh and reset
' ---------------------------------------------------------------------
Public Sub RefreshAllReports()
    On Error GoTo Fail
    Application.ScreenUpdating = False
    ThisWorkbook.RefreshAll
    Application.CalculateFull
    If SheetExists(PANEL) Then
        ThisWorkbook.Worksheets(PANEL).Range("B2").Value = "Last refreshed: " & Format(Now, "yyyy-mm-dd hh:nn")
    End If
    Application.ScreenUpdating = True
    MsgBox "All pivot tables refreshed and formulas recalculated.", vbInformation, APP_TITLE
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Refresh stopped: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub ResetAllFilters()
    Dim sc As SlicerCache
    On Error Resume Next
    For Each sc In ThisWorkbook.SlicerCaches
        sc.ClearManualFilter
    Next sc
    ThisWorkbook.Worksheets("Dashboard").Range("C4").Value = "FY2026"
    ThisWorkbook.Worksheets("Dashboard").Range("F4").Value = "All"
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------
'  4. Distribution: PDF, dealer workbooks, email draft
' ---------------------------------------------------------------------
Public Sub ExportDashboardPDF()
    Dim pdfPath As String
    Dim startSheet As Object

    On Error GoTo Fail
    If ThisWorkbook.Path = "" Then
        MsgBox "Save the workbook first so the PDF has a folder to go to.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & _
              "Command_Center_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    ThisWorkbook.Activate
    Set startSheet = ActiveSheet
    ThisWorkbook.Sheets(Array("Dashboard", "Findings")).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    startSheet.Select
    MsgBox "Saved: " & pdfPath, vbInformation, APP_TITLE
    Exit Sub
Fail:
    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").Select
    MsgBox "PDF export stopped: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub CreateDealerReports()
    Dim lo As ListObject
    Dim wsDealers As Worksheet, wsOrders As Worksheet, wsSummary As Worksheet
    Dim wbNew As Workbook
    Dim folderPath As String, dealerId As String, dealerName As String
    Dim r As Long, made As Long, lastRow As Long

    On Error GoTo Fail
    If ThisWorkbook.Path = "" Then
        MsgBox "Save the workbook first so the reports have a folder to go to.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    If MsgBox("Create one workbook per dealer in a Dealer_Reports folder next to this file?", _
              vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub

    folderPath = ThisWorkbook.Path & Application.PathSeparator & "Dealer_Reports"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    Set lo = ThisWorkbook.Worksheets(SRC_SHEET).ListObjects(SRC_TABLE)
    Set wsDealers = ThisWorkbook.Worksheets("Dim_Dealer")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    r = 5
    Do While Len(CStr(wsDealers.Cells(r, 1).Value)) > 0
        dealerId = CStr(wsDealers.Cells(r, 1).Value)
        dealerName = CStr(wsDealers.Cells(r, 2).Value)

        lo.Range.AutoFilter Field:=3, Criteria1:=dealerId
        Set wbNew = Workbooks.Add(xlWBATWorksheet)
        Set wsOrders = wbNew.Worksheets(1)
        wsOrders.Name = "Orders"
        lo.Range.SpecialCells(xlCellTypeVisible).Copy
        wsOrders.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        wsOrders.Rows(1).Font.Bold = True
        wsOrders.Columns.AutoFit
        lastRow = wsOrders.Cells(wsOrders.Rows.Count, 1).End(xlUp).Row

        Set wsSummary = wbNew.Worksheets.Add(Before:=wsOrders)
        wsSummary.Name = "Summary"
        WriteDealerSummary wsSummary, dealerId, dealerName, lastRow

        wbNew.SaveAs Filename:=folderPath & Application.PathSeparator & dealerId & "_report_" & _
                     Format(Date, "yyyymmdd") & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
        made = made + 1
        r = r + 1
    Loop

    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox made & " dealer reports saved in:" & vbCrLf & folderPath, vbInformation, APP_TITLE
    Exit Sub
Fail:
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Dealer reports stopped: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub WriteDealerSummary(ws As Worksheet, dealerId As String, dealerName As String, lastRow As Long)
    Dim c As Long
    Dim fy As String, col As String
    Dim fyRng As String, netRng As String, gmRng As String, listRng As String, onTimeRng As String

    ' Column letters match the Sales_Data layout: L List Value, M Net Sales,
    ' P Gross Margin, U On Time, V Fiscal Year
    fyRng = "Orders!$V$2:$V$" & lastRow
    netRng = "Orders!$M$2:$M$" & lastRow
    gmRng = "Orders!$P$2:$P$" & lastRow
    listRng = "Orders!$L$2:$L$" & lastRow
    onTimeRng = "Orders!$U$2:$U$" & lastRow

    ws.Range("A1").Value = dealerName & " (" & dealerId & ")"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A3").Value = "Metric"
    ws.Range("A4").Value = "Orders"
    ws.Range("A5").Value = "Net sales"
    ws.Range("A6").Value = "Gross margin"
    ws.Range("A7").Value = "Gross margin %"
    ws.Range("A8").Value = "Average discount"
    ws.Range("A9").Value = "On-time delivery %"

    For c = 0 To 2
        fy = "FY" & CStr(2024 + c)
        col = Chr(66 + c)
        ws.Range(col & "3").Value = fy
        ws.Range(col & "4").Formula = "=COUNTIF(" & fyRng & ",""" & fy & """)"
        ws.Range(col & "5").Formula = "=SUMIF(" & fyRng & ",""" & fy & """," & netRng & ")"
        ws.Range(col & "6").Formula = "=SUMIF(" & fyRng & ",""" & fy & """," & gmRng & ")"
        ws.Range(col & "7").Formula = "=IF(" & col & "5=0,0," & col & "6/" & col & "5)"
        ws.Range(col & "8").Formula = "=IF(SUMIF(" & fyRng & ",""" & fy & """," & listRng & ")=0,0,1-" & _
                                      col & "5/SUMIF(" & fyRng & ",""" & fy & """," & listRng & "))"
        ws.Range(col & "9").Formula = "=IF(" & col & "4=0,0,SUMIF(" & fyRng & ",""" & fy & """," & _
                                      onTimeRng & ")/" & col & "4)"
    Next c

    ws.Range("A3:D3").Font.Bold = True
    ws.Range("B4:D4").NumberFormat = "#,##0"
    ws.Range("B5:D6").NumberFormat = "$#,##0"
    ws.Range("B7:D9").NumberFormat = "0.0%"
    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B:D").ColumnWidth = 16
End Sub

Public Sub EmailKpiSummary()
    Dim ol As Object, mail As Object
    Dim d As Worksheet, f As Worksheet
    Dim body As String
    Dim i As Long

    Set d = ThisWorkbook.Worksheets("Dashboard")
    Set f = ThisWorkbook.Worksheets("Findings")

    body = "Sales & Operations summary - " & CStr(d.Range("C4").Value) & _
           ", region: " & CStr(d.Range("F4").Value) & vbCrLf & vbCrLf
    body = body & "Net sales: " & FmtMoney(d.Range("B7").Value) & vbCrLf
    body = body & "YoY growth: " & FmtPct(d.Range("D7").Value) & vbCrLf
    body = body & "Gross margin: " & FmtPct(d.Range("F7").Value) & vbCrLf
    body = body & "Average discount: " & FmtPct(d.Range("H7").Value) & vbCrLf
    body = body & "On-time delivery: " & FmtPct(d.Range("J7").Value) & vbCrLf
    body = body & "Excess inventory: " & FmtMoney(d.Range("L7").Value) & vbCrLf
    body = body & "FY2027 forecast: " & FmtMoney(d.Range("N7").Value) & vbCrLf & vbCrLf
    body = body & "What needs attention:" & vbCrLf
    For i = 1 To 5
        body = body & i & ". " & CStr(f.Range("B" & (5 + i)).Value) & vbCrLf
    Next i

    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then
        MsgBox "Outlook is not available on this computer. The summary text is shown below so you can copy it." & _
               vbCrLf & vbCrLf & Left(body, 800), vbInformation, APP_TITLE
        Exit Sub
    End If
    Set mail = ol.CreateItem(0)
    mail.Subject = "Sales & Operations summary - " & CStr(d.Range("C4").Value)
    mail.body = body
    mail.Display
End Sub

' ---------------------------------------------------------------------
'  5. Analyst tools: data quality check and discount exceptions
' ---------------------------------------------------------------------
Public Sub RunDataQualityCheck()
    Dim lo As ListObject
    Dim arr As Variant
    Dim ws As Worksheet
    Dim seen As Collection
    Dim i As Long, j As Long, outRow As Long, issues As Long
    Dim hasErr As Boolean
    Dim orderId As String

    On Error GoTo Fail
    Set lo = ThisWorkbook.Worksheets(SRC_SHEET).ListObjects(SRC_TABLE)
    arr = lo.DataBodyRange.Value
    Application.ScreenUpdating = False

    Set ws = ResetSheet("Data_Quality_Log")
    ws.Range("A1").Value = "Data quality log for " & SRC_SHEET
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A4:D4").Value = Array("Sheet row", "Order ID", "Check", "Detail")
    ws.Range("A4:D4").Font.Bold = True
    outRow = 5
    Set seen = New Collection

    For i = 1 To UBound(arr, 1)
        ' Formula or lookup errors anywhere in the row
        hasErr = False
        For j = 1 To UBound(arr, 2)
            If IsError(arr(i, j)) Then
                hasErr = True
                LogIssue ws, outRow, i + 1, "(see row)", "Formula error", _
                         "Column: " & CStr(lo.HeaderRowRange.Cells(1, j).Value)
                Exit For
            End If
        Next j

        If Not hasErr Then
            orderId = Trim(CStr(arr(i, 1)))
            If Len(orderId) = 0 Then
                LogIssue ws, outRow, i + 1, "", "Missing key", "Order ID is blank"
            Else
                On Error Resume Next
                seen.Add i, orderId
                If Err.Number <> 0 Then
                    Err.Clear
                    On Error GoTo Fail
                    LogIssue ws, outRow, i + 1, orderId, "Duplicate", "Order ID appears more than once"
                End If
                On Error GoTo Fail
            End If
            If Len(Trim(CStr(arr(i, 3)))) = 0 Then LogIssue ws, outRow, i + 1, orderId, "Missing key", "Dealer ID is blank"
            If Len(Trim(CStr(arr(i, 5)))) = 0 Then LogIssue ws, outRow, i + 1, orderId, "Missing key", "Product ID is blank"
            If Not IsDate(arr(i, 2)) Then LogIssue ws, outRow, i + 1, orderId, "Invalid date", "Order Date is not a date"
            If OutOfRange(arr(i, 9), 1, 1000) Then LogIssue ws, outRow, i + 1, orderId, "Out of range", "Units should be 1 to 1,000"
            If OutOfRange(arr(i, 11), 0, 0.25) Then LogIssue ws, outRow, i + 1, orderId, "Out of range", "Discount should be 0% to 25%"
            If OutOfRange(arr(i, 13), 0.01, 1E+12) Then LogIssue ws, outRow, i + 1, orderId, "Out of range", "Net Sales should be positive"
            If IsNumeric(arr(i, 16)) Then
                If CDbl(arr(i, 16)) < 0 Then LogIssue ws, outRow, i + 1, orderId, "Negative margin", "Order sold below cost"
            End If
            If IsDate(arr(i, 2)) And IsDate(arr(i, 19)) Then
                If CDate(arr(i, 19)) < CDate(arr(i, 2)) Then LogIssue ws, outRow, i + 1, orderId, "Date logic", "Delivered before it was ordered"
            End If
        End If
        If outRow > 5004 Then Exit For
    Next i

    issues = outRow - 5
    ws.Range("A2").Value = "Rows checked: " & UBound(arr, 1) & "   Issues logged: " & issues & _
                           IIf(outRow > 5004, " (stopped at 5,000)", "") & "   Run: " & Format(Now, "yyyy-mm-dd hh:nn")
    ws.Columns("A:D").AutoFit
    Application.ScreenUpdating = True
    MsgBox "Checked " & UBound(arr, 1) & " rows. Issues logged: " & issues & ".", vbInformation, APP_TITLE
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Data quality check stopped: " & Err.Description, vbExclamation, APP_TITLE
End Sub

Private Sub LogIssue(ws As Worksheet, ByRef outRow As Long, sheetRow As Long, orderId As String, _
                     checkName As String, detail As String)
    ws.Cells(outRow, 1).Value = sheetRow
    ws.Cells(outRow, 2).Value = orderId
    ws.Cells(outRow, 3).Value = checkName
    ws.Cells(outRow, 4).Value = detail
    outRow = outRow + 1
End Sub

Private Function OutOfRange(v As Variant, minVal As Double, maxVal As Double) As Boolean
    If IsEmpty(v) Then
        OutOfRange = True
    ElseIf Not IsNumeric(v) Then
        OutOfRange = True
    Else
        OutOfRange = (CDbl(v) < minVal Or CDbl(v) > maxVal)
    End If
End Function

Public Sub DiscountExceptionReport()
    Dim lo As ListObject
    Dim arr As Variant, answer As Variant
    Dim ws As Worksheet
    Dim limitPct As Double, totalNet As Double, totalGM As Double
    Dim i As Long, outRow As Long

    On Error GoTo Fail
    answer = Application.InputBox("List orders with a discount at or above (decimal, for example 0.15):", _
                                  "Discount exception report", 0.15, Type:=1)
    If VarType(answer) = vbBoolean Then Exit Sub
    limitPct = CDbl(answer)
    If limitPct > 1 Then limitPct = limitPct / 100

    Set lo = ThisWorkbook.Worksheets(SRC_SHEET).ListObjects(SRC_TABLE)
    arr = lo.DataBodyRange.Value
    Application.ScreenUpdating = False

    Set ws = ResetSheet("Discount_Exceptions")
    ws.Range("A1").Value = "Orders discounted " & Format(limitPct, "0.0%") & " or more"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A4:K4").Value = Array("Order ID", "Order Date", "Dealer ID", "Region", "Product ID", "Segment", _
                                    "Discount %", "Net Sales", "Gross Margin", "GM %", "Fiscal Year")
    ws.Range("A4:K4").Font.Bold = True
    outRow = 5

    For i = 1 To UBound(arr, 1)
        If IsNumeric(arr(i, 11)) And IsNumeric(arr(i, 13)) And IsNumeric(arr(i, 16)) Then
            If CDbl(arr(i, 11)) >= limitPct Then
                ws.Cells(outRow, 1).Value = arr(i, 1)
                ws.Cells(outRow, 2).Value = arr(i, 2)
                ws.Cells(outRow, 3).Value = arr(i, 3)
                ws.Cells(outRow, 4).Value = arr(i, 4)
                ws.Cells(outRow, 5).Value = arr(i, 5)
                ws.Cells(outRow, 6).Value = arr(i, 7)
                ws.Cells(outRow, 7).Value = arr(i, 11)
                ws.Cells(outRow, 8).Value = arr(i, 13)
                ws.Cells(outRow, 9).Value = arr(i, 16)
                ws.Cells(outRow, 10).Value = arr(i, 17)
                ws.Cells(outRow, 11).Value = arr(i, 22)
                totalNet = totalNet + CDbl(arr(i, 13))
                totalGM = totalGM + CDbl(arr(i, 16))
                outRow = outRow + 1
            End If
        End If
    Next i

    If outRow > 5 Then
        ws.Range("A4:K" & (outRow - 1)).Sort Key1:=ws.Range("G4"), Order1:=xlDescending, Header:=xlYes
        ws.Range("B5:B" & (outRow - 1)).NumberFormat = "mmm d, yyyy"
        ws.Range("G5:G" & (outRow - 1)).NumberFormat = "0.0%"
        ws.Range("J5:J" & (outRow - 1)).NumberFormat = "0.0%"
        ws.Range("H5:I" & (outRow - 1)).NumberFormat = "$#,##0"
    End If
    ws.Range("A2").Value = "Orders: " & (outRow - 5) & "   Net sales: " & Format(totalNet, "$#,##0") & _
                           "   Gross margin: " & Format(totalGM, "$#,##0") & _
                           IIf(totalNet > 0, "   Margin: " & Format(totalGM / IIf(totalNet = 0, 1, totalNet), "0.0%"), "")
    ws.Columns("A:K").AutoFit
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Discount report stopped: " & Err.Description, vbExclamation, APP_TITLE
End Sub

' ---------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------
Private Function SheetExists(sheetName As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (sh Is Nothing)
End Function

Private Sub DeleteSheet(sheetName As String)
    If SheetExists(sheetName) Then ThisWorkbook.Sheets(sheetName).Delete
End Sub

Private Function ResetSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim oldAlerts As Boolean
    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    DeleteSheet sheetName
    Application.DisplayAlerts = oldAlerts
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    Set ResetSheet = ws
End Function

Private Function FmtMoney(v As Variant) As String
    If IsError(v) Then
        FmtMoney = "n/a"
    ElseIf IsNumeric(v) Then
        FmtMoney = Format(CDbl(v) / 1000000, "$#,##0.0") & "M"
    Else
        FmtMoney = CStr(v)
    End If
End Function

Private Function FmtPct(v As Variant) As String
    If IsError(v) Then
        FmtPct = "n/a"
    ElseIf IsNumeric(v) Then
        FmtPct = Format(CDbl(v), "0.0%")
    Else
        FmtPct = CStr(v)
    End If
End Function
