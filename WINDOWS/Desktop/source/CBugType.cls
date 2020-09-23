VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBugType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************
' CBugType
' by Matthew Hickson (BMG)
' written: 03/14/2001
' updated: --/--/----
'
' Purpose:
' To model a bug type for the BMGBugTracker software
'*****************************************************************

Public Enum BMGBugTypes
   enmUndefined = 0
   enmBug = 1
   enmFeatureRequest = 2
   enmAlterationRequest = 3
End Enum

Private enmID As BMGBugTypes
Private strDescription As String

'*****************************************************************
' Class Functions / Subs
'*****************************************************************
Public Function GetBugTypes() As String
On Error GoTo CleanUp
   
   Dim cnBugDatabase As ADODB.Connection
   Dim rsBugTypeList As ADODB.Recordset
   
   'Open connection
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Get records
   Set rsBugTypeList = New ADODB.Recordset
   rsBugTypeList.Open "SELECT Description FROM BugType ORDER BY ID", _
                      cnBugDatabase
                          
   'Return apps as string
   GetBugTypes = rsBugTypeList.GetString
   
CleanUp:
   'Clean up
   rsBugTypeList.Close
   cnBugDatabase.Close
   
   Set rsBugTypeList = Nothing
   Set cnBugDatabase = Nothing
End Function

'*****************************************************************
' Properties
'*****************************************************************
Public Property Let ID(pID As BMGBugTypes)
   enmID = pID
End Property

Public Property Get ID() As BMGBugTypes
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


