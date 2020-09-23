Attribute VB_Name = "modUtility"
Option Explicit

'*****************************************************************
' modUtility
' by Matthew Hickson (BMG)
' written: 03/14/2001
' updated: 03/15/2001 - MDH
'
' Purpose:
' To house utility functions of BMGBugTracker software
'*****************************************************************

Declare Function SendMessage _
         Lib "user32" _
         Alias "SendMessageA" ( _
         ByVal hwnd As Long, _
         ByVal wMsg As Long, _
         ByVal wParam As Integer, _
         ByVal lParam As Any _
         ) As Long
         
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_LIMITTEXT = &H141
Public Const CB_ERR = (-1)

Public g_strConnectionString As String

Public Function SearchCombo(cb As ComboBox, sItem As String) As Long
   SearchCombo = SendMessage(cb.hwnd, CB_FINDSTRINGEXACT, -1, ByVal sItem)
End Function
