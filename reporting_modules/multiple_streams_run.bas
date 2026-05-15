Option Explicit

Sub Run_ReportingStream(streamName As String, reportType As String, reportingPeriod As Long)

    Dim sourceWorkbook As Workbook
    Dim reportWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim reportWorksheet As Worksheet
    Dim summaryWorksheet As Worksheet

    Dim sourcePath As String
    Dim templatePath As String
    Dim outputFolder As String

    Dim periodText As String
    Dim periodCell As Range
    Dim startRow As Long

    Dim copyColumns As String
    Dim outputFileName As String
    Dim pdfPath As String

    periodText = "PERIOD" & reportingPeriod

    sourcePath = "./source_data/reporting_input.xlsx"
    outputFolder = "./outputs/"

    If reportType = "WEEKLY" Then

        templatePath = "./templates/weekly_report_template.xlsx"
        copyColumns = "C:H"

        outputFileName = streamName & "_weekly_period_" & reportingPeriod & ".xlsx"
        pdfPath = outputFolder & streamName & "_weekly_period_" & reportingPeriod & ".pdf"

    ElseIf reportType = "PERIOD_TO_DATE" Then

        templatePath = "./templates/period_to_date_report_template.xlsx"
        copyColumns = "N:S"

        outputFileName = streamName & "_period_to_date_period_" & reportingPeriod & ".xlsx"
        pdfPath = outputFolder & streamName & "_period_to_date_period_" & reportingPeriod & ".pdf"

    Else

        MsgBox "Invalid report type: " & reportType
        Exit Sub

    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    Set sourceWorkbook = Workbooks.Open(sourcePath, ReadOnly:=True, UpdateLinks:=False)
    Set reportWorkbook = Workbooks.Open(templatePath, UpdateLinks:=False)

    Set sourceWorksheet = sourceWorkbook.Sheets(streamName)
    Set reportWorksheet = reportWorkbook.Sheets("Report")
    Set summaryWorksheet = reportWorkbook.Sheets("Summary")

    Set periodCell = sourceWorksheet.Columns("D").Find( _
        What:=periodText, _
        LookIn:=xlValues, _
        LookAt:=xlWhole _
    )

    If periodCell Is Nothing Then

        MsgBox periodText & " not found for stream " & streamName

        GoTo CleanExit

    End If

    startRow = periodCell.Row + 3

    sourceWorksheet.Range( _
        Split(copyColumns, ":")(0) & startRow & ":" & _
        Split(copyColumns, ":")(1) & startRow + 84 _
    ).Copy

    reportWorksheet.Range("E4").PasteSpecial Paste:=xlPasteValues

    Application.CutCopyMode = False

    summaryWorksheet.Range("B3").Formula = summaryWorksheet.Range("B3").Formula

    summaryWorksheet.Calculate

    reportWorkbook.SaveCopyAs outputFolder & outputFileName

    summaryWorksheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

CleanExit:

    On Error Resume Next

    sourceWorkbook.Close SaveChanges:=False
    reportWorkbook.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True

End Sub