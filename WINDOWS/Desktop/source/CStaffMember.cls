VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStaffMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************
' CStaffMember
' by Matthew Hickson (BMG)
' written: 03/14/2001
' updated: --/--/----
'
' Purpose:
' To model a staff member for the BMGBugTracker software
'*****************************************************************

Private lngID As Long
Private strDescription As String

'*****************************************************************
' Class Functions / Subs
'*****************************************************************
Public Function GetStaffMembers() As String
On Error GoTo CleanUp

   Dim cnBugDatabase As ADODB.Connection
   Dim rsStaffMembersList As ADODB.Recordset
   
   'Open connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsStaffMembersList = New ADODB.Recordset
   rsStaffMembersList.Open "SELECT Description FROM Staff ORDER BY ID", _
                           cnBugDatabase
                          
   'Return apps as string
   GetStaffMembers = rsStaffMembersList.GetString
   
CleanUp:
   'Clean up
   rsStaffMembersList.Close
   cnBugDatabase.Close
   
   Set rsStaffMembersList = Nothing
   Set cnBugDatabase = Nothing
End Function

Public Function ResolveStaffMemberID(pStaffMemberName As String) As Long
On Error GoTo CleanUp

   Dim cnBugDatabase As ADODB.Connection
   Dim rsStaffMemberID As ADODB.Recordset
   
   'Open connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsStaffMemberID = New ADODB.Recordset
   rsStaffMemberID.Open "SELECT ID FROM Staff WHERE Description='" & pStaffMemberName & "'", _
                        cnBugDatabase
                          
   
   ResolveStaffMemberID = rsStaffMemberID.Fields("ID")
   
CleanUp:
   'Clean up
   rsStaffMemberID.Close
   cnBugDatabase.Close
   
   Set rsStaffMemberID = Nothing
   Set cnBugDatabase = Nothing
End Function

Public Function ResolveStaffMemberName(pStaffMemberID As String) As String
On Error GoTo CleanUp

   Dim cnBugDatabase As ADODB.Connection
   Dim rsStaffMemberName As ADODB.Recordset
   
   'Open connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsStaffMemberName = New ADODB.Recordset
   rsStaffMemberName.Open "SELECT Description FROM Staff WHERE ID=" & pStaffMemberID, _
                          cnBugDatabase
                          
   
   ResolveStaffMemberName = rsStaffMemberName.Fields("Description")
   
CleanUp:
   'Clean up
   rsStaffMemberName.Close
   cnBugDatabase.Close
   
   Set rsStaffMemberName = Nothing
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

Public Property Let Description(pDescription As String)
   strDescription = pDescription
End Property

Public Property Get Description() As String
   Description = strDescription
End Property

'*****************************************************************
' Class Constructor / Destructor
'*****************************************************************
Private Sub Class_Initialize()
   lngID = 0
   strDescription = ""
End Sub

Private Sub Class_Terminate()
   'Nothing
End Sub


