Attribute VB_Name = "modReports"
Option Explicit

'*****************************************************************
' modReports
' by Matthew Hickson (BMG)
' written: 03/16/2001
' updated: 03/19/2001
'
' Purpose:
' To house (HTML) reporting functions of BMGBugTracker software
'*****************************************************************

Public Sub ApplicationReport(Optional pPreview As Boolean = True)
On Error GoTo CleanUp
   
   Dim cnBugDatabase As ADODB.Connection
   Dim rsBugList As ADODB.Recordset
   Dim objReport As CHTMLWrapper
   Dim lngBugCnt As Long
   Dim objTempApp As CApplication
   Dim objTempStaff As CStaffMember
   
   'Open Connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsBugList = New ADODB.Recordset
   rsBugList.Open "SELECT * FROM BugData ORDER BY AppID, ID", cnBugDatabase
   
   If Not rsBugList.BOF Then
      Set objReport = New CHTMLWrapper
   
      With objReport
         .OpenHTMLDoc g_strAppPath & "application.html"
         .OpenHTMLHeader "Application Status for " & Format$(Now, "mm/dd/yyyy")
         .CloseHTMLHeader
         
         .OpenHTMLBody
         .WriteHeader 1, "Application Status for " & Format$(Now, "mm/dd/yyyy")
         .InsertHorizontalRule
         
         .StartTable 1
         
         'Parse report data
                 
         'Print Table Headings
         .StartTableRow
         .WriteTableHeader "Application"
         .WriteTableHeader "Bug ID"
         .WriteTableHeader "Type"
         .WriteTableHeader "Severity"
         .WriteTableHeader "Status"
         .WriteTableHeader "Assigned To"
         .WriteTableHeader "Reported By"
         .WriteTableHeader "Date Reported"
         .WriteTableHeader "Behaviour Summary"
         .EndTableRow
         
         Set objTempApp = New CApplication
         Set objTempStaff = New CStaffMember
         
         lngBugCnt = 0
         
         Do While Not rsBugList.EOF
            lngBugCnt = lngBugCnt + 1
         
            .StartTableRow
            
            .InsertComment "APPLICATION: " & objTempApp.ResolveApplicationName(rsBugList.Fields("AppID"))
            
            .WriteTableData objTempApp.ResolveApplicationName(rsBugList.Fields("AppID"))
            .WriteTableData Format$(rsBugList.Fields("ID"), "0000000")
            .WriteTableData rsBugList.Fields("TypeID")
            .WriteTableData rsBugList.Fields("SeverityID")
            .WriteTableData rsBugList.Fields("StatusID")
            .WriteTableData objTempStaff.ResolveStaffMemberName(rsBugList.Fields("AssignedID"))
            .WriteTableData rsBugList.Fields("ReportedBy")
            .WriteTableData Format$(rsBugList.Fields("DateReported"), "mmm-dd-yyyy")
            .WriteTableData Left$(rsBugList.Fields("ObservedBehaviour"), 80) & "..."
            
            .EndTableRow
            
            rsBugList.MoveNext
         Loop
         .EndTable
                     
         Set objTempApp = Nothing
         Set objTempStaff = Nothing
                     
         .InsertHorizontalRule
         .WriteItalicText lngBugCnt & " bugs reported"
         .InsertBreak
         .WriteItalicText "End of Application Status for " & Format$(Now, "mm/dd/yyyy")
      End With
      
      'Show the report (with default web browser
      Shell "start " & g_strAppPath & "application.html", vbHide
   Else
      MsgBox "No status to report!", , "Status"
   End If
   
   
CleanUp:
   rsBugList.Close
   cnBugDatabase.Close
   
   Set objReport = Nothing
   Set rsBugList = Nothing
   Set cnBugDatabase = Nothing
End Sub

Public Sub BugListReport(Optional pPreview As Boolean = True)
On Error GoTo CleanUp
   
   Dim cnBugDatabase As ADODB.Connection
   Dim rsBugList As ADODB.Recordset
   Dim objReport As CHTMLWrapper
   Dim lngBugCnt As Long
   Dim objTempApp As CApplication
   
   'Open Connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsBugList = New ADODB.Recordset
   rsBugList.Open "SELECT * FROM BugData ORDER BY ID", cnBugDatabase
   
   If Not rsBugList.BOF Then
      Set objReport = New CHTMLWrapper
   
      With objReport
         .OpenHTMLDoc g_strAppPath & "buglist.html"
         .OpenHTMLHeader "Bug List for " & Format$(Now, "mm/dd/yyyy")
         .CloseHTMLHeader
         
         .OpenHTMLBody
         .WriteHeader 1, "Bug List for " & Format$(Now, "mm/dd/yyyy")
         .InsertHorizontalRule
         
         .StartTable 1
         
         'Parse report data
                 
         'Print Table Headings
         .StartTableRow
         .WriteTableHeader "ID"
         .WriteTableHeader "Application"
         .WriteTableHeader "Type"
         .WriteTableHeader "Severity"
         .WriteTableHeader "Status"
         .WriteTableHeader "Assigned To"
         .WriteTableHeader "Reported By"
         .WriteTableHeader "Date Reported"
         .WriteTableHeader "Behaviour Summary"
         .EndTableRow
         
         Set objTempApp = New CApplication
         
         lngBugCnt = 0
         
         Do While Not rsBugList.EOF
            lngBugCnt = lngBugCnt + 1
         
            .StartTableRow
            
            .InsertComment "BUG ID: " & Format$(rsBugList.Fields("ID"), "0000000")
            
            .WriteTableData Format$(rsBugList.Fields("ID"), "0000000")
            .WriteTableData objTempApp.ResolveApplicationName(rsBugList.Fields("AppID"))
            .WriteTableData rsBugList.Fields("TypeID")
            .WriteTableData rsBugList.Fields("SeverityID")
            .WriteTableData rsBugList.Fields("StatusID")
            .WriteTableData rsBugList.Fields("AssignedID")
            .WriteTableData rsBugList.Fields("ReportedBy")
            .WriteTableData Format$(rsBugList.Fields("DateReported"), "mmm-dd-yyyy")
            .WriteTableData Left$(rsBugList.Fields("ObservedBehaviour"), 80) & "..."
            
            .EndTableRow
            
            rsBugList.MoveNext
         Loop
         .EndTable
                     
         Set objTempApp = Nothing
                     
         .InsertHorizontalRule
         .WriteItalicText lngBugCnt & " bugs reported"
         .InsertBreak
         .WriteItalicText "End of Bug List for " & Format$(Now, "mm/dd/yyyy")
      End With
   
      'Show the report (with default web browser
      Shell "start " & g_strAppPath & "buglist.html", vbHide
   Else
      MsgBox "No status to report!", , "Status"
   End If
   
CleanUp:
   rsBugList.Close
   cnBugDatabase.Close
   
   Set objReport = Nothing
   Set rsBugList = Nothing
   Set cnBugDatabase = Nothing
End Sub

Public Sub DeveloperStatusReport(Optional pPreview As Boolean = True)
On Error GoTo CleanUp
   
   Dim cnBugDatabase As ADODB.Connection
   Dim rsBugList As ADODB.Recordset
   Dim objReport As CHTMLWrapper
   Dim lngBugCnt As Long
   Dim objTempApp As CApplication
   
   'Open Connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsBugList = New ADODB.Recordset
   rsBugList.Open "SELECT * FROM BugData ORDER BY AssignedID, AppID, ID", cnBugDatabase
   
   If Not rsBugList.BOF Then
      Set objReport = New CHTMLWrapper
   
      With objReport
         .OpenHTMLDoc g_strAppPath & "developer.html"
         .OpenHTMLHeader "Developer Status List for " & Format$(Now, "mm/dd/yyyy")
         .CloseHTMLHeader
         
         .OpenHTMLBody
         .WriteHeader 1, "Developer Status List for " & Format$(Now, "mm/dd/yyyy")
         .InsertHorizontalRule
         
         .StartTable 1
         
         'Parse report data
                 
         'Print Table Headings
         .StartTableRow
         .WriteTableHeader "Assigned To"
         .WriteTableHeader "Application"
         .WriteTableHeader "ID"
         .WriteTableHeader "Type"
         .WriteTableHeader "Severity"
         .WriteTableHeader "Status"
         .WriteTableHeader "Reported By"
         .WriteTableHeader "Date Reported"
         .WriteTableHeader "Behaviour Summary"
         .EndTableRow
         
         Set objTempApp = New CApplication
         
         lngBugCnt = 0
         
         Do While Not rsBugList.EOF
            lngBugCnt = lngBugCnt + 1
         
            .StartTableRow
            
            .WriteTableData rsBugList.Fields("AssignedID")
            .WriteTableData objTempApp.ResolveApplicationName(rsBugList.Fields("AppID"))
            .WriteTableData Format$(rsBugList.Fields("ID"), "0000000")
            .WriteTableData rsBugList.Fields("TypeID")
            .WriteTableData rsBugList.Fields("SeverityID")
            .WriteTableData rsBugList.Fields("StatusID")
            .WriteTableData rsBugList.Fields("ReportedBy")
            .WriteTableData Format$(rsBugList.Fields("DateReported"), "mmm-dd-yyyy")
            .WriteTableData Left$(rsBugList.Fields("ObservedBehaviour"), 80) & "..."
            
            .EndTableRow
            
            rsBugList.MoveNext
         Loop
         .EndTable
                     
         Set objTempApp = Nothing
                     
         .InsertHorizontalRule
         .WriteItalicText lngBugCnt & " bugs reported"
         .InsertBreak
         .WriteItalicText "End of Developer Status List for " & Format$(Now, "mm/dd/yyyy")
      End With
   
      'Show the report (with default web browser
      Shell "start " & g_strAppPath & "developer.html", vbHide
   Else
      MsgBox "No status to report!", , "Status"
   End If
      
CleanUp:
   rsBugList.Close
   cnBugDatabase.Close
   
   Set objReport = Nothing
   Set rsBugList = Nothing
   Set cnBugDatabase = Nothing
End Sub

