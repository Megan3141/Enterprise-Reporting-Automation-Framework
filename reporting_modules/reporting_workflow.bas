Option Explicit

Sub Run_ReportingWorkflow(reportType As String, reportingPeriod As Long)

    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet

    Dim reportWorkbook As Workbook
    Dim destinationWorksheet As Worksheet
    Dim summaryWorksheet As Worksheet

    Dim sourcePath As String
    Dim templatePath As String
    Dim outputFolder As String

    Dim sourceStartColumn As String
    Dim sourceRows As Variant
    Dim destinationRows As Variant

    Dim baseRow As Long
    Dim lastRow As Long
    Dim cell As Range
    Dim checkValue As String
    Dim periodFound As Boolean

    Dim i As Long

    Dim outputFileName As String
    Dim pdfPath As String

    reportType = Trim(UCase(reportType))

    '-------------------------
    ' CONFIGURATION
    '-------------------------
    outputFolder = "./outputs/"
    sourcePath = "./source_data/reporting_input.xlsx"

    '-------------------------
    ' REPORT TYPE CONFIGURATION
    '-------------------------
    If reportType = "WEEKLY" Then
        templatePath = "./templates/weekly_report_template.xlsx"
        sourceStartColumn = "C"
    Else
        templatePath = "./templates/period_to_date_report_template.xlsx"
        sourceStartColumn = "L"
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    '-------------------------
    ' OPEN SOURCE DATA
    '-------------------------
    Set sourceWorkbook = Workbooks.Open(sourcePath, ReadOnly:=True, UpdateLinks:=False)
    Set sourceWorksheet = sourceWorkbook.Sheets("reporting_input")

    '-------------------------
    ' OPEN REPORT TEMPLATE
    '-------------------------
    Set reportWorkbook = Workbooks.Open(templatePath, UpdateLinks:=False)

    Set destinationWorksheet = reportWorkbook.Sheets("Report")
    Set summaryWorksheet = reportWorkbook.Sheets("Summary")

    '-------------------------
    ' FIND REPORTING PERIOD
    '-------------------------
    lastRow = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, "A").End(xlUp).Row
    periodFound = False

    For Each cell In sourceWorksheet.Range("A1:A" & lastRow)

        checkValue = Replace(Replace(Trim(UCase(CStr(cell.Text))), " ", ""), Chr(160), "")

        If checkValue = "PERIOD" & reportingPeriod Then
            baseRow = cell.Row
            periodFound = True
            Exit For
        End If

    Next cell

    If periodFound = False Then

        MsgBox "Reporting period " & reportingPeriod & " was not found in the source data."

        sourceWorkbook.Close False
        reportWorkbook.Close False

        GoTo CleanExit

    End If

    '-------------------------
    ' ROW MAPPING
    '-------------------------
    sourceRows = Array( _
        1, 2, 3, 4, 5, 6, 7, 8, _
        9, 10, 11, 12, 13, 14, 15, 16, _
        17, 18, 19, 20)

    If reportType = "WEEKLY" Then

        destinationRows = Array( _
            4, 5, 6, 7, 8, 9, 10, 11, _
            12, 13, 14, 15, 16, 17, 19, 20, _
            21, 22, 23, 24)

        destinationWorksheet.Range("C4:I24").ClearContents

    Else

        destinationRows = Array( _
            4, 5, 6, 7, 8, 9, 10, 11, _
            12, 13, 14, 15, 16, 17, 18, 19, _
            20, 21, 22, 23)

        destinationWorksheet.Range("C4:I23").ClearContents

    End If

    '-------------------------
    ' COPY VALUES ONLY
    '-------------------------
    For i = LBound(sourceRows) To UBound(sourceRows)

        destinationWorksheet.Range("C" & destinationRows(i)).Resize(1, 7).Value = _
            sourceWorksheet.Range(sourceStartColumn & (baseRow + sourceRows(i))).Resize(1, 7).Value

    Next i

    '-------------------------
    ' REFRESH SUMMARY VIEW
    '-------------------------
    With summaryWorksheet.Range("B4")
        .Formula = .Formula
    End With

    summaryWorksheet.Calculate

    '-------------------------
    ' OUTPUT FILE NAMES
    '-------------------------
    If reportType = "WEEKLY" Then
        outputFileName = "reporting_output_period_" & reportingPeriod & ".xlsx"
        pdfPath = outputFolder & "reporting_output_period_" & reportingPeriod & ".pdf"
    Else
        outputFileName = "period_to_date_reporting_output_period_" & reportingPeriod & ".xlsx"
        pdfPath = outputFolder & "period_to_date_reporting_output_period_" & reportingPeriod & ".pdf"
    End If

    '-------------------------
    ' SAVE OUTPUT COPY
    '-------------------------
    If Dir(outputFolder & outputFileName) <> "" Then
        Kill outputFolder & outputFileName
    End If

    reportWorkbook.SaveCopyAs outputFolder & outputFileName

    '-------------------------
    ' EXPORT PDF
    '-------------------------
    summaryWorksheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    sourceWorkbook.Close False
    reportWorkbook.Close False

CleanExit:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True

End Sub


Sub Run_Weekly_And_Period_To_Date_Reports()

    Dim periodInput As String
    Dim reportingPeriod As Long

    periodInput = InputBox("Enter Reporting Period")
    If periodInput = "" Then Exit Sub

    reportingPeriod = CLng(periodInput)

    Run_ReportingWorkflow "WEEKLY", reportingPeriod
    Run_ReportingWorkflow "PERIOD_TO_DATE", reportingPeriod

    MsgBox "Reporting workflows complete."

End Sub
