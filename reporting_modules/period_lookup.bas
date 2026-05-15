Option Explicit

Function FindReportingPeriodRow( _
    ByVal sourceWorksheet As Worksheet, _
    ByVal lookupColumn As String, _
    ByVal periodNumber As Long, _
    ByVal periodFormat As String _
) As Long

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim expectedValue As String
    Dim actualValue As String

    Select Case UCase(Trim(periodFormat))

        Case "WEEK_TEXT"
            expectedValue = "WEEK " & periodNumber

        Case "WK_PREFIX"
            expectedValue = "WK" & periodNumber

        Case "PERIOD_PREFIX"
            expectedValue = "PERIOD" & periodNumber

        Case Else
            FindReportingPeriodRow = 0
            Exit Function

    End Select

    lastRow = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, lookupColumn).End(xlUp).Row

    For rowIndex = 1 To lastRow

        actualValue = UCase(Trim(CStr(sourceWorksheet.Cells(rowIndex, lookupColumn).Value)))

        If actualValue = expectedValue Then
            FindReportingPeriodRow = rowIndex
            Exit Function
        End If

    Next rowIndex

    FindReportingPeriodRow = 0

End Function