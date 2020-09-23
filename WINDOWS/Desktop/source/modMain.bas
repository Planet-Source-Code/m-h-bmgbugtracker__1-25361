Attribute VB_Name = "modMain"
Option Explicit

'*****************************************************************
' modMain
' by Matthew Hickson (BMG)
' written: 03/16/2001
' updated: 03/19/2001
'
' Purpose:
' To prepare and launch BMGBugTracker software
'*****************************************************************

Public g_strAppPath As String

Public Sub Main()
   'Fix App path if necessary
   g_strAppPath = App.Path
   If Right$(g_strAppPath, 1) <> "\" Then
      g_strAppPath = g_strAppPath & "\"
   End If
   
   'Prime Connection string for database functions
   g_strConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
                           "Persist Security Info=False;" & _
                           "Data Source=" & g_strAppPath & "bugdata.mdb"
   
   Dim objTempBugType As CBugType
   Dim objTempSeverity As CSeverity
   Dim objTempStaff As CStaffMember
   Dim objTempStatus As CStatus
   Dim objTempApplication As CApplication
   
   'Pass BugTypes to Main form
   Set objTempBugType = New CBugType
   frmMain.BugTypeList = objTempBugType.GetBugTypes
   Set objTempBugType = Nothing
   
   'Pass Severity Levels to Main form
   Set objTempSeverity = New CSeverity
   frmMain.SeverityList = objTempSeverity.GetSeverityLevels
   Set objTempSeverity = Nothing

   'Pass Staff Members to Main form
   Set objTempStaff = New CStaffMember
   frmMain.StaffList = objTempStaff.GetStaffMembers
   Set objTempStaff = Nothing

   'Pass Status Levels to Main form
   Set objTempStatus = New CStatus
   frmMain.StatusList = objTempStatus.GetStatusLevels
   Set objTempStatus = Nothing

   'Pass Software Packages to Main form
   Set objTempApplication = New CApplication
   frmMain.ApplicationList = objTempApplication.GetAllApplications
   Set objTempApplication = Nothing
      
   'Show main form
   frmMain.Show
   
   'END OF PROGRAM (well, control passed to frmMain)
End Sub
