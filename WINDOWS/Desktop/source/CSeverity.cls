VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSeverity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************
' CSeverity
' by Matthew Hickson (BMG)
' written: 03/14/2001
' updated: --/--/----
'
' Purpose:
' To model a severity level for the BMGBugTracker software
'*****************************************************************

Public Enum BMGBugSeverityTypes
   enmUndefined = 0
   enmLow = 1
   enmModerate = 2
   enmHigh = 3
   enmVeryHigh = 4
End Enum

Private enmID As BMGBugSeverityTypes
Private strDescription As String

'*****************************************************************
' Class Functions / Subs
'*****************************************************************
Public Function GetSeverityLevels() As String
On Error GoTo CleanUp
   
   Dim cnBugDatabase As ADODB.Connection
   Dim rsSeverityLevelList As ADODB.Recordset
   
   'Open connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsSeverityLevelList = New ADODB.Recordset
   rsSeverityLevelList.Open "SELECT Description FROM BugSeverity ORDER BY ID", _
                            cnBugDatabase
                          
   'Return apps as string
   GetSeverityLevels = rsSeverityLevelList.GetString
   
CleanUp:
   'Clean up
   rsSeverityLevelList.Close
   cnBugDatabase.Close
   
   Set rsSeverityLevelList = Nothing
   Set cnBugDatabase = Nothing
End Function

'*****************************************************************
' Properties
'*****************************************************************
Public Property Let ID(pID As BMGBugSeverityTypes)
   enmID = pID
End Property

Public Property Get ID() As BMGBugSeverityTypes
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

