VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFindBug 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Bug"
   ClientHeight    =   2520
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   8292
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   8292
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEndDate 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   492
      Left            =   5040
      TabIndex        =   9
      Top             =   1920
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   492
      Left            =   7200
      TabIndex        =   11
      Top             =   1920
      Width           =   972
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   492
      Left            =   6120
      TabIndex        =   10
      Top             =   1920
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Caption         =   "Results"
      Height          =   1812
      Left            =   3000
      TabIndex        =   7
      Top             =   0
      Width           =   5172
      Begin MSFlexGridLib.MSFlexGrid msfFindResults 
         Height          =   1452
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4932
         _ExtentX        =   8700
         _ExtentY        =   2561
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedRows       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.TextBox txtKeyWords 
      Height          =   288
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2772
   End
   Begin VB.ComboBox cboApplication 
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   2772
   End
   Begin VB.TextBox txtStartDate 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "End Date:"
      Height          =   492
      Left            =   1560
      TabIndex        =   12
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "Keywords:"
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "Application:"
      Height          =   492
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1212
   End
   Begin VB.Label Label7 
      Caption         =   "Start Date:"
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1212
   End
End
Attribute VB_Name = "frmFindBug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************
' frmFindBug
' by Matthew Hickson (BMG)
' written: 03/30/2001
' updated: 04/02/2001
'
' Purpose:
' Search form of BMGBugTracker software (for finding Bug IDs)
'*****************************************************************

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()
   Dim cn As ADODB.Connection
   Dim rs As ADODB.Recordset
   Dim objTempApp As CApplication
   Dim lngAppID As Long
   Dim dteStartDate As Date
   Dim dteEndDate As Date
   Dim dteSwapTemp As Date
   Dim astrKeywords() As String
   Dim lngKeywordCtr As Long
   Dim strSQL As String
   
   msfFindResults.Clear
   msfFindResults.Rows = 0
   
   strSQL = ""
   strSQL = strSQL & "SELECT ID, AppID, DateReported, StepsToReproduce "
   strSQL = strSQL & "FROM BugData "
   strSQL = strSQL & "WHERE (1=1) " 'Always true therefore, always records
   
   'Get app identification
   If cboApplication.Text <> "" Then
      Set objTempApp = New CApplication
      lngAppID = objTempApp.ResolveApplicationID(cboApplication.Text)
      Set objTempApp = Nothing
      
      strSQL = strSQL & "AND (AppID = " & lngAppID & ") "
   Else
      lngAppID = -1
   End If
   
   'Get date range
   If IsDate(txtStartDate.Text) Then
      dteStartDate = CDate(txtStartDate.Text)
      
      If IsDate(txtEndDate.Text) Then
         dteEndDate = CDate(txtEndDate.Text)
         
         'Properly set dates (no good if start date is later than end date)
         If dteStartDate > dteEndDate Then
            dteSwapTemp = dteStartDate
            dteStartDate = dteEndDate
            dteEndDate = dteSwapTemp
            
            txtStartDate.Text = Format(dteStartDate, "mm/dd/yyyy")
            txtEndDate.Text = Format(dteEndDate, "mm/dd/yyyy")
         End If
         
         strSQL = strSQL & "AND (DateReported >=#" & dteStartDate & "#) "
         strSQL = strSQL & "AND (DateReported <=#" & dteEndDate & "#) "
      End If
   End If
   
   'Get keywords...
   If Trim$(txtKeyWords.Text) <> "" Then
      astrKeywords = Split(txtKeyWords.Text, " ")
   
      For lngKeywordCtr = LBound(astrKeywords) To UBound(astrKeywords)
         strSQL = strSQL & "AND (InStr(1, StepsToReproduce, '" & astrKeywords(lngKeywordCtr) & "') > 0) "
      Next lngKeywordCtr
   End If
   
   'Finish off SQL
   strSQL = strSQL & "ORDER BY ID;"
   
   'Get Data
   Set cn = New ADODB.Connection
   Set rs = New ADODB.Recordset
   
   cn.Open g_strConnectionString
   rs.Open strSQL, cn
   
   'Show results
   Set objTempApp = New CApplication
   With rs
      Do While Not .EOF
         msfFindResults.AddItem _
            Format$(.Fields("ID"), "0000000") & vbTab & _
            objTempApp.ResolveApplicationName(.Fields("AppID")) & vbTab & _
            Format$(.Fields("DateReported"), "mm/dd/yyyy") & vbTab & _
            .Fields("StepsToReproduce")
      
         .MoveNext
      Loop
   End With
   Set objTempApp = Nothing
   
   Set rs = Nothing
   Set cn = Nothing
End Sub

Private Sub cmdOk_Click()
   With msfFindResults
      If .RowSel >= 0 Then
         frmMain.FoundBugID = .TextMatrix(.RowSel, 0)
      End If
   End With

   Unload Me
End Sub

Private Sub Form_Load()
   Dim lngItemCtr As Long

   'Populate Application List
   For lngItemCtr = 1 To frmMain.cboApplication.ListCount - 1
      cboApplication.AddItem frmMain.cboApplication.List(lngItemCtr)
   Next lngItemCtr
   
   'Init Grid to nothing
   With msfFindResults
      .Clear
      .Rows = 0
      
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      
      .ColWidth(0) = .Width * 0.15
      .ColWidth(1) = .Width * 0.25
      .ColWidth(2) = .Width * 0.15
      .ColWidth(3) = .Width * 0.4
   End With
End Sub

Private Sub msfFindResults_DblClick()
   With msfFindResults
      If .RowSel >= 0 Then
         frmMain.FoundBugID = .TextMatrix(.RowSel, 0)
      End If
   End With

   Unload Me
End Sub
