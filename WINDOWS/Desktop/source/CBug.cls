VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************
' CBug
' by Matthew Hickson (BMG)
' written: 03/14/2001
' updated: 03/15/2001 - MDH
'
' Purpose:
' To model a bug for the BMGBugTracker software
'*****************************************************************

Private lngID As Long
Private objApplication As CApplication
Private objBugType As CBugType
Private objSeverity As CSeverity
Private objStatus As CStatus
Private objAssignedTo As CStaffMember
Private strReportedBy As String
Private dteReported As Date
Private strStepsToReproduce As String
Private strExpectedBehaviour As String
Private strObservedBehaviour As String

'*****************************************************************
' Class Functions / Subs
'*****************************************************************
Public Function GetBugList() As String
On Error GoTo CleanUp
   
   Dim cnBugDatabase As ADODB.Connection
   Dim rsBugList As ADODB.Recordset
   
   'Open connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsBugList = New ADODB.Recordset
   rsBugList.Open "SELECT ID FROM BugData ORDER BY ID", _
                  cnBugDatabase
                          
   'Return apps as string
   GetBugList = rsBugList.GetString
   
CleanUp:
   'Clean up
   rsBugList.Close
   cnBugDatabase.Close
   
   Set rsBugList = Nothing
   Set cnBugDatabase = Nothing
End Function

Public Function LoadBug(BugToLoad As Long) As Boolean
   Dim cnBugDatabase As ADODB.Connection
   Dim rsBug As ADODB.Recordset
   Dim objTempApp As CApplication
   
   'Default of LoadBug is False
   LoadBug = False

   'Open connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsBug = New ADODB.Recordset
   rsBug.Open "SELECT ID, AppID, TypeID, SeverityID, StatusID, AssignedID, " & _
                  "ReportedBy, DateReported, " & _
                  "StepsToReproduce, ExpectedBehaviour, ObservedBehaviour " & _
                  "FROM BugData " & _
                  "WHERE ID=" & BugToLoad & " " & _
                  "ORDER BY ID", _
                  cnBugDatabase

   Set objTempApp = New CApplication
   
   'Connect Recordset to item
   With rsBug
      If Not .BOF Then
         Me.ID = .Fields("ID")
         Me.Application = objTempApp.ResolveApplicationName(.Fields("AppID"))
         Me.BugType = .Fields("TypeID")
         Me.Severity = .Fields("SeverityID")
         Me.Status = .Fields("StatusID")
         Me.AssignedTo = objAssignedTo.ResolveStaffMemberName(.Fields("AssignedID"))
         Me.ReportedBy = .Fields("ReportedBy")
         Me.Reported = .Fields("DateReported")
         Me.StepsToReproduce = .Fields("StepsToReproduce")
         Me.ExpectedBehaviour = .Fields("ExpectedBehaviour")
         Me.ObservedBehaviour = .Fields("ObservedBehaviour")
         
         'We got the bug loaded
         LoadBug = True
      End If
   End With
   
   Set objTempApp = New CApplication

   'Clean up
   rsBug.Close
   cnBugDatabase.Close
   
   Set rsBug = Nothing
   Set cnBugDatabase = Nothing
End Function

Public Function SaveBug() As Boolean
   Dim cnBugDatabase As ADODB.Connection
   Dim strSaveSQL As String
   Dim lngRecordsAffected As Long
   Dim objTempApp As CApplication
   
   Set objTempApp = New CApplication
   
   'Assume Save Fails
   SaveBug = False
   
   'Build SQL for save / update
   If Me.ID = -1 Then 'Fresh bug
      strSaveSQL = "INSERT INTO BugData(" & _
                   "AppID, TypeID, SeverityID, StatusID, AssignedID, " & _
                   "ReportedBy, DateReported, " & _
                   "StepsToReproduce, ExpectedBehaviour, ObservedBehaviour" & _
                   ") SELECT " & _
                   objTempApp.ResolveApplicationID(Me.Application) & " AS bAID, " & _
                   Me.BugType & " AS bTID, " & _
                   Me.Severity & " AS bSID, " & _
                   Me.Status & " AS bStID, " & _
                   objAssignedTo.ResolveStaffMemberID(Me.AssignedTo) & " AS bAsID, '" & _
                   Me.ReportedBy & "' AS bReportedBy, #" & _
                   Me.Reported & "# AS bReported, '" & _
                   Replace$(Me.StepsToReproduce, "'", "''") & "' AS bSteps, '" & _
                   Replace$(Me.ExpectedBehaviour, "'", "''") & "' AS bExpect, '" & _
                   Replace$(Me.ObservedBehaviour, "'", "''") & "' AS bObserve"
   Else 'Saving edited bug
      strSaveSQL = "UPDATE BugData " & _
                   "SET " & _
                   "AppID=" & objTempApp.ResolveApplicationID(Me.Application) & ", " & _
                   "TypeID=" & Me.BugType & ", " & _
                   "SeverityID=" & Me.Severity & ", " & _
                   "StatusID=" & Me.Status & ", " & _
                   "AssignedID=" & objAssignedTo.ResolveStaffMemberID(Me.AssignedTo) & ", " & _
                   "ReportedBy='" & Me.ReportedBy & "', " & _
                   "DateReported=#" & Me.Reported & "#, " & _
                   "StepsToReproduce='" & Replace$(Me.StepsToReproduce, "'", "''") & "', " & _
                   "ExpectedBehaviour='" & Replace$(Me.ExpectedBehaviour, "'", "''") & "', " & _
                   "ObservedBehaviour='" & Replace$(Me.ObservedBehaviour, "'", "''") & "' " & _
                   "WHERE ID=" & Me.ID
   End If
   
   Set objTempApp = Nothing
   
   'Execute SQL
   With cnBugDatabase
      Set cnBugDatabase = New ADODB.Connection
      cnBugDatabase.Open g_strConnectionString
      cnBugDatabase.Execute strSaveSQL, lngRecordsAffected
   End With
   
   'Bug saved properly
   If lngRecordsAffected = 1 Then
      SaveBug = True
   End If
   
   'Clean up
   cnBugDatabase.Close
   Set cnBugDatabase = Nothing
End Function

Public Function FindBugID(pLengthOfID As Long) As String
   Dim cnBugDatabase As ADODB.Connection
   Dim rsBugID As ADODB.Recordset
   Dim objTempApp As CApplication
   
   'Open connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   Set objTempApp = New CApplication
   
   'This should be sufficient to identify a bug (and return it's ID)
   Set rsBugID = New ADODB.Recordset
   rsBugID.Open "SELECT ID " & _
                  "FROM BugData " & _
                  "WHERE " & _
                  "(AppID=" & objTempApp.ResolveApplicationID(Me.Application) & ") " & _
                  "AND (TypeID=" & Me.BugType & ") " & _
                  "AND (SeverityID=" & Me.Severity & ") " & _
                  "AND (StatusID=" & Me.Status & ") " & _
                  "AND (ReportedBy='" & Me.ReportedBy & "') " & _
                  "AND (DateReported=#" & Me.Reported & "#) ", _
                  cnBugDatabase

   Set objTempApp = New CApplication
   
   'Return ID
   FindBugID = "" & Format$(rsBugID.Fields("ID"), String(pLengthOfID, "0"))

   'Clean up
   rsBugID.Close
   cnBugDatabase.Close
   
   Set rsBugID = Nothing
   Set cnBugDatabase = Nothing
End Function

'*****************************************************************
' Properties
'*****************************************************************
Public Property Let ID(pID As Long)
   lngID = pID
End Property

Public Property Get ID() As Long
   ID = lngID
End Property

Public Property Let Application(pApplication As String)
   objApplication.Description = pApplication
End Property

Public Property Get Application() As String
   Application = objApplication.Description
End Property

Public Property Let BugType(pBugType As BMGBugTypes)
   objBugType.ID = pBugType
End Property

Public Property Get BugType() As BMGBugTypes
   BugType = objBugType.ID
End Property

Public Property Let Severity(pSeverity As BMGBugSeverityTypes)
   objSeverity.ID = pSeverity
End Property

Public Property Get Severity() As BMGBugSeverityTypes
   Severity = objSeverity.ID
End Property

Public Property Let Status(pStatus As BMGBugStatusTypes)
   objStatus.ID = pStatus
End Property

Public Property Get Status() As BMGBugStatusTypes
   Status = objStatus.ID
End Property

Public Property Let ReportedBy(pReportedBy As String)
   strReportedBy = pReportedBy
End Property

Public Property Get ReportedBy() As String
   ReportedBy = strReportedBy
End Property

Public Property Let AssignedTo(pAssignedTo As String)
   objAssignedTo.Description = pAssignedTo
End Property

Public Property Get AssignedTo() As String
   AssignedTo = objAssignedTo.Description
End Property

Public Property Let Reported(pReported As Date)
   dteReported = pReported
End Property

Public Property Get Reported() As Date
   Reported = dteReported
End Property

Public Property Let StepsToReproduce(pStepsToReproduce As String)
   strStepsToReproduce = pStepsToReproduce
End Property

Public Property Get StepsToReproduce() As String
   StepsToReproduce = strStepsToReproduce
End Property

Public Property Let ExpectedBehaviour(pExpectedBehaviour As String)
   strExpectedBehaviour = pExpectedBehaviour
End Property

Public Property Get ExpectedBehaviour() As String
   ExpectedBehaviour = strExpectedBehaviour
End Property

Public Property Let ObservedBehaviour(pObservedBehaviour As String)
   strObservedBehaviour = pObservedBehaviour
End Property

Public Property Get ObservedBehaviour() As String
   ObservedBehaviour = strObservedBehaviour
End Property

'*****************************************************************
' Class Constructor / Destructor
'*****************************************************************
Private Sub Class_Initialize()
   lngID = -1
   Set objApplication = New CApplication
   Set objBugType = New CBugType
   Set objSeverity = New CSeverity
   Set objStatus = New CStatus
   Set objAssignedTo = New CStaffMember
   strReportedBy = ""
   dteReported = Now
   strStepsToReproduce = ""
   strExpectedBehaviour = ""
   strObservedBehaviour = ""
End Sub

Private Sub Class_Terminate()
   Set objApplication = Nothing
   Set objBugType = Nothing
   Set objSeverity = Nothing
   Set objStatus = Nothing
   Set objAssignedTo = Nothing
End Sub
