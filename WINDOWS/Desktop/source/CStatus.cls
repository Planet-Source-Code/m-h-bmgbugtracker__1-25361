VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************
' CStatus
' by Matthew Hickson (BMG)
' written: 03/14/2001
' updated: 03/16/2001 - MDH
'
' Purpose:
' To model a status level for the BMGBugTracker software
'*****************************************************************

Public Enum BMGBugStatusTypes
   enmUndefined = 0
   enmFixed = 1
   enmWillNotFix = 2
   enmPostponed = 3
   enmNotReproducable = 4
   enmDuplicate = 5
   enmByDesign = 6
   enmUnassigned = 7
End Enum

Private enmID As BMGBugStatusTypes
Private strDescription As String

'*****************************************************************
' Class Functions / Subs
'*****************************************************************
Public Function GetStatusLevels() As String
On Error GoTo CleanUp
   
   Dim cnBugDatabase As ADODB.Connection
   Dim rsStatusLevelList As ADODB.Recordset
   
   'Open connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsStatusLevelList = New ADODB.Recordset
   rsStatusLevelList.Open "SELECT Description FROM BugStatus ORDER BY ID", _
                          cnBugDatabase
                          
   'Return apps as string
   GetStatusLevels = rsStatusLevelList.GetString
   
CleanUp:
   'Clean up
   rsStatusLevelList.Close
   cnBugDatabase.Close
   
   Set rsStatusLevelList = Nothing
   Set cnBugDatabase = Nothing
End Function

'*****************************************************************
' Properties
'*****************************************************************
Public Property Let ID(pID As BMGBugStatusTypes)
   enmID = pID
End Property

Public Property Get ID() As BMGBugStatusTypes
   ID = enmID
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
   enmID = enmUndefined
   strDescription = ""
End Sub

Private Sub Class_Terminate()
   'Nothing
End Sub


