Option Explicit

Sub Run_All_Workflows()

    Dim periodInput As String
    Dim reportingPeriod As Long

    periodInput = InputBox("Enter Reporting Period")

    If periodInput = "" Then Exit Sub

    If Not IsNumeric(periodInput) Then
        MsgBox "Please enter a valid reporting period."
        Exit Sub
    End If

    reportingPeriod = CLng(periodInput)

    Run_ReportingWorkflow "WEEKLY", reportingPeriod
    Run_ReportingWorkflow "PERIOD_TO_DATE", reportingPeriod

    MsgBox "All reporting workflows complete for period " & reportingPeriod

End Sub