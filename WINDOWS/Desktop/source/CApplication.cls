VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************
' CApplication
' by Matthew Hickson (BMG)
' written: 03/14/2001
' updated: 03/15/2001 - MDH
'
' Purpose:
' To model an application for the BMGBugTracker software
'*****************************************************************

Private lngID As Long
Private strDescription As String

'*****************************************************************
' Class Functions / Subs
'*****************************************************************
Public Function GetAllApplications() As String
On Error GoTo CleanUp
   
   Dim cnBugDatabase As ADODB.Connection
   Dim rsApplicationList As ADODB.Recordset
   
   'Open connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsApplicationList = New ADODB.Recordset
   rsApplicationList.Open "SELECT Description FROM Applications ORDER BY Description", _
                          cnBugDatabase
                          
   'Return apps as string
   GetAllApplications = rsApplicationList.GetString
   
CleanUp:
   'Clean up
   rsApplicationList.Close
   cnBugDatabase.Close
   
   Set rsApplicationList = Nothing
   Set cnBugDatabase = Nothing
End Function

Public Function ResolveApplicationID(pApplicationName As String) As Long
On Error GoTo CleanUp
   
   Dim cnBugDatabase As ADODB.Connection
   Dim rsApplicationID As ADODB.Recordset
   
   'Open connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsApplicationID = New ADODB.Recordset
   rsApplicationID.Open "SELECT ID FROM Applications WHERE Description='" & pApplicationName & "'", _
                          cnBugDatabase
                          
   'Return apps as string
   ResolveApplicationID = rsApplicationID.Fields("ID")
   
CleanUp:
   'Clean up
   rsApplicationID.Close
   cnBugDatabase.Close
   
   Set rsApplicationID = Nothing
   Set cnBugDatabase = Nothing
End Function

Public Function ResolveApplicationName(pApplicationID As Long) As String
On Error GoTo CleanUp
   
   Dim cnBugDatabase As ADODB.Connection
   Dim rsApplicationName As ADODB.Recordset
   
   'Open connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsApplicationName = New ADODB.Recordset
   rsApplicationName.Open "SELECT Description FROM Applications WHERE ID=" & pApplicationID, _
                          cnBugDatabase
                          
   'Return apps as string
   ResolveApplicationName = rsApplicationName.Fields("Description")
   
CleanUp:
   'Clean up
   rsApplicationName.Close
   cnBugDatabase.Close
   
   Set rsApplicationName = Nothing
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
