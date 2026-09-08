Attribute VB_Name = "Module_ReportAutomation"
'==============================================================================
' Performance KPI Dashboard & Reporting Automation
' Module: Module_ReportAutomation
'
' Purpose:
'   Automates the recurring monthly HR KPI report cycle:
'     1. Refresh all Power Query connections (raw HRIS export -> cleaned tables)
'     2. Refresh every PivotTable / PivotChart built on fact_monthly_kpi
'     3. Re-stamp the "Report Generated" timestamp + reporting month on the
'        Summary sheet
'     4. Export the Summary sheet to a dated PDF into \Reports\<Year>\
'     5. Optionally attach + email the PDF to the distribution list on the
'        Config sheet via Outlook
'
' Timing: No reproducible benchmark is included. Measure the complete
' manual and automated workflows before claiming a time-saving percentage.
'==============================================================================

Option Explicit

Public Sub RunMonthlyReport()

    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo CleanFail

    RefreshAllQueries
    RefreshAllPivots
    StampReportMetadata
    Dim pdfPath As String
    pdfPath = ExportSummaryToPDF()

    ' Flip to True + configure Config sheet to auto-email each run
    If Range("Config!EmailOnRun").Value = True Then
        EmailReport pdfPath
    End If

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Monthly KPI report refreshed in " & Format(Timer - startTime, "0.0") & _
           " seconds." & vbCrLf & "PDF saved to: " & pdfPath, vbInformation, "Report Automation"
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "RunMonthlyReport failed: " & Err.Description, vbCritical, "Report Automation"
End Sub

'------------------------------------------------------------------------
' Refreshes every Power Query connection in the workbook (background = False
' so we wait for completion before moving to the pivot refresh step)
'------------------------------------------------------------------------
Private Sub RefreshAllQueries()
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        On Error Resume Next
        conn.OLEDBConnection.BackgroundQuery = False
        conn.Refresh
        On Error GoTo 0
    Next conn
End Sub

'------------------------------------------------------------------------
' Refreshes every PivotTable on every worksheet (dashboard pivots + the
' hidden staging pivots that feed the charts)
'------------------------------------------------------------------------
Private Sub RefreshAllPivots()
    Dim ws As Worksheet
    Dim pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    ThisWorkbook.RefreshAll
End Sub

'------------------------------------------------------------------------
' Updates the "Report Generated" timestamp and reporting period label on
' the Summary sheet so every export is self-documenting
'------------------------------------------------------------------------
Private Sub StampReportMetadata()
    With ThisWorkbook.Worksheets("Summary")
        .Range("ReportGeneratedOn").Value = Now
        .Range("ReportingPeriod").Value = Format(DateAdd("m", -1, Date), "mmmm yyyy")
    End With
End Sub

'------------------------------------------------------------------------
' Exports the Summary sheet to a dated PDF under \Reports\<Year>\
' Returns the full path written.
'------------------------------------------------------------------------
Private Function ExportSummaryToPDF() As String
    Dim folderPath As String, fileName As String, fullPath As String
    folderPath = ThisWorkbook.Path & "\Reports\" & Format(Date, "yyyy") & "\"

    If Dir(folderPath, vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\Reports"
        MkDir folderPath
    End If

    fileName = "HR_KPI_Report_" & Format(DateAdd("m", -1, Date), "yyyy_mm") & ".pdf"
    fullPath = folderPath & fileName

    ThisWorkbook.Worksheets("Summary").ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=fullPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ExportSummaryToPDF = fullPath
End Function

'------------------------------------------------------------------------
' Emails the generated PDF to the distribution list on the Config sheet
' using Outlook (late-bound so this compiles even without an Outlook
' reference set)
'------------------------------------------------------------------------
Private Sub EmailReport(ByVal pdfPath As String)
    Dim OutApp As Object, OutMail As Object
    Dim recipients As String

    recipients = ThisWorkbook.Worksheets("Config").Range("DistributionList").Value
    If Len(Trim(recipients)) = 0 Then Exit Sub

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = recipients
        .Subject = "HR Performance KPI Report - " & Format(DateAdd("m", -1, Date), "mmmm yyyy")
        .Body = "Hi team," & vbCrLf & vbCrLf & _
                "Attached is the automated monthly HR performance KPI report." & vbCrLf & _
                "Generated on " & Format(Now, "mmmm d, yyyy h:mm AM/PM") & "." & vbCrLf & vbCrLf & _
                "Regards," & vbCrLf & "HR Analytics"
        .Attachments.Add pdfPath
        .Send
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
