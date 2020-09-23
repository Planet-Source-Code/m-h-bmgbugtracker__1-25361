VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMG Bug Tracker"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboAssignedTo 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox txtReportedBy 
      Height          =   285
      Left            =   120
      TabIndex        =   19
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox txtObservedBehaviour 
      Height          =   885
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   4680
      Width           =   4455
   End
   Begin VB.TextBox txtExpectedBehaviour 
      Height          =   885
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   3480
      Width           =   4455
   End
   Begin VB.TextBox txtStepsToReproduce 
      Height          =   885
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   2280
      Width           =   4455
   End
   Begin VB.ComboBox cboBugType 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtReported 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.ComboBox cboApplication 
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1680
      Width           =   4455
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cboSeverity 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtBugID 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Assigned To:"
      Height          =   495
      Left            =   2400
      TabIndex        =   20
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Reported By:"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Observed Behaviour:"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label10 
      Caption         =   "Expected Behaviour:"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label9 
      Caption         =   "Steps To Reproduce Bug:"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "Date Reported:"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Type of Bug:"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Application:"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Current Status:"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Severity:"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Bug ID:"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewBug 
         Caption         =   "&New Bug"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpenBug 
         Caption         =   "&Open Bug"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSaveBug 
         Caption         =   "&Save Bug"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPrintBug 
         Caption         =   "&Print Bug"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuFindBug 
         Caption         =   "&Find Bug"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuApplicationStatus 
         Caption         =   "Application Status"
      End
      Begin VB.Menu mnuBugList 
         Caption         =   "Bug List"
      End
      Begin VB.Menu mnuDeveloperStatus 
         Caption         =   "Developer Status"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuAddApplication 
         Caption         =   "Add Application"
      End
      Begin VB.Menu mnuAddDeveloper 
         Caption         =   "Add Developer"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************
' frmMain
' by Matthew Hickson (BMG)
' written: 03/14/2001
' updated: 03/15/2001 - MDH
' updated: 03/16/2001 - MDH
'
' Purpose:
' Main form of BMGBugTracker software (for user input)
'*****************************************************************

Private objBug As CBug
Private strBugTypeList As String
Private strSeverityList As String
Private strStaffList As String
Private strStatusList As String
Private strApplicationList As String
Private lngFoundBugID As Long

Public Property Let FoundBugID(pFoundBugID As Long)
   lngFoundBugID = pFoundBugID
End Property

Public Property Get FoundBugID() As Long
   FoundBugID = lngFoundBugID
End Property

Public Property Let BugTypeList(pBugTypeList As String)
   strBugTypeList = pBugTypeList
End Property

Public Property Get BugTypeList() As String
   BugTypeList = strBugTypeList
End Property

Public Property Let SeverityList(pSeverityList As String)
   strSeverityList = pSeverityList
End Property

Public Property Get SeverityList() As String
   SeverityList = strSeverityList
End Property

Public Property Let StaffList(pStaffList As String)
   strStaffList = pStaffList
End Property

Public Property Get StaffList() As String
   StaffList = strStaffList
End Property

Public Property Let StatusList(pStatusList As String)
   strStatusList = pStatusList
End Property

Public Property Get StatusList() As String
   StatusList = strStatusList
End Property

Public Property Let ApplicationList(pApplicationList As String)
   strApplicationList = pApplicationList
End Property

Public Property Get ApplicationList() As String
   ApplicationList = strApplicationList
End Property

Private Sub ClearForm()
   Dim objCurrentCtrl As Control
   
   For Each objCurrentCtrl In Me.Controls
      If (TypeOf objCurrentCtrl Is TextBox) Then
         objCurrentCtrl.Text = ""
      ElseIf (TypeOf objCurrentCtrl Is ComboBox) Then
         If objCurrentCtrl.ListCount > 0 Then
            objCurrentCtrl.ListIndex = 0
         End If
      End If
   Next objCurrentCtrl
   
   txtReported.Text = Format$(Now, "mm/dd/yyyy")
   
   'Make sure we start at the beginning of the form
   If txtReported.Visible Then
      txtReported.SetFocus
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Dim arrTempValues() As String
   Dim lngLB As Long
   Dim lngUB As Long
   Dim lngItemCtr As Long
   
   'Set up a new bug object
   Set objBug = New CBug
   
   'Populate Bug Type Combo
   If strBugTypeList <> "" Then
      arrTempValues = Split(strBugTypeList, vbCr)
      lngLB = LBound(arrTempValues)
      lngUB = UBound(arrTempValues) - 1
      
      For lngItemCtr = lngLB To lngUB
         cboBugType.AddItem arrTempValues(lngItemCtr)
      Next lngItemCtr
   End If
   
   'Populate Severity Combo
   If strSeverityList <> "" Then
      arrTempValues = Split(strSeverityList, vbCr)
      lngLB = LBound(arrTempValues)
      lngUB = UBound(arrTempValues) - 1
      
      For lngItemCtr = lngLB To lngUB
         cboSeverity.AddItem arrTempValues(lngItemCtr)
      Next lngItemCtr
   End If
   
   'Populate Staff Combo (Assigned To)
   If strStaffList <> "" Then
      arrTempValues = Split(strStaffList, vbCr)
      lngLB = LBound(arrTempValues)
      lngUB = UBound(arrTempValues) - 1
      
      For lngItemCtr = lngLB To lngUB
         cboAssignedTo.AddItem arrTempValues(lngItemCtr)
      Next lngItemCtr
   End If

   'Populate Status Combo (Assigned To)
   If strStatusList <> "" Then
      arrTempValues = Split(strStatusList, vbCr)
      lngLB = LBound(arrTempValues)
      lngUB = UBound(arrTempValues) - 1
      
      For lngItemCtr = lngLB To lngUB
         cboStatus.AddItem arrTempValues(lngItemCtr)
      Next lngItemCtr
   End If

   'Populate Application List
   If strApplicationList <> "" Then
      arrTempValues = Split(strApplicationList, vbCr)
      lngLB = LBound(arrTempValues)
      lngUB = UBound(arrTempValues) - 1
      
      For lngItemCtr = lngLB To lngUB
         cboApplication.AddItem arrTempValues(lngItemCtr)
      Next lngItemCtr
   End If
   
   ClearForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set objBug = Nothing
End Sub

Private Sub mnuAbout_Click()
   MsgBox "BMG Bug Tracker" & vbNewLine & _
          "Copyright 2001" & vbNewLine, _
          vbInformation, _
          "About"
End Sub

Private Sub mnuAddApplication_Click()
On Error GoTo CleanUp
   Dim cnBugDatabase As ADODB.Connection
   Dim lngRecordsAffected As Long
   Dim strApplicationDescription As String
   
   strApplicationDescription = InputBox("Enter an application description: ", "Add Application")
   
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Perform the addition
   If Trim$(strApplicationDescription) <> "" Then
      cnBugDatabase.Execute "INSERT INTO Applications(Description) SELECT '" & strApplicationDescription & "' AS aDescription", lngRecordsAffected
      
      If lngRecordsAffected = 1 Then
         cboApplication.AddItem strApplicationDescription
      Else
         MsgBox "Application could not be added!", , "Error"
      End If
   End If
   
CleanUp:
   cnBugDatabase.Close
   Set cnBugDatabase = Nothing
End Sub

Private Sub mnuAddDeveloper_Click()
On Error GoTo CleanUp
   Dim cnBugDatabase As ADODB.Connection
   Dim lngRecordsAffected As Long
   Dim strDeveloper As String
   
   strDeveloper = InputBox("Enter a developer: ", "Add Developer")
   
   Set cnBugDatabase = New ADODB.Connection
   cnBugDatabase.Open g_strConnectionString
   
   'Perform the addition
   If Trim$(strDeveloper) <> "" Then
      cnBugDatabase.Execute "INSERT INTO Staff(Description) SELECT '" & strDeveloper & "' AS sDeveloper", lngRecordsAffected
      
      If lngRecordsAffected = 1 Then
         cboAssignedTo.AddItem strDeveloper
      Else
         MsgBox "Developer could not be added!", , "Error"
      End If
   End If
   
CleanUp:
   cnBugDatabase.Close
   Set cnBugDatabase = Nothing
End Sub

Private Sub mnuApplicationStatus_Click()
   modReports.ApplicationReport
End Sub

Private Sub mnuBugList_Click()
   modReports.BugListReport
End Sub

Private Sub mnuDeveloperStatus_Click()
   modReports.DeveloperStatusReport
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub

Private Sub mnuFindBug_Click()
   Dim BugLoaded As Boolean
   
   FoundBugID = -1

   frmFindBug.Show vbModal, Me
   
   If FoundBugID <> -1 Then
      BugLoaded = objBug.LoadBug(FoundBugID)
      
      If BugLoaded Then
         DisplayBug objBug
      Else
         MsgBox "Bug " & FoundBugID & " could not be opened!", , "Error"
      End If
   End If
End Sub

Private Sub mnuNewBug_Click()
   ClearForm
   Set objBug = Nothing
   Set objBug = New CBug
End Sub

Private Sub mnuOpenBug_Click()
   Dim BugLoaded As Boolean
   Dim strBugRequested As String
   Dim lngBugRequested As Long
   
   strBugRequested = InputBox("Enter Bug ID:", "Open Bug")
   
   If IsNumeric(Trim$(strBugRequested)) Then
      lngBugRequested = CLng(Val(strBugRequested))
      
      BugLoaded = objBug.LoadBug(lngBugRequested)
      
      If BugLoaded Then
         DisplayBug objBug
      Else
         MsgBox "Bug " & lngBugRequested & " could not be opened!", , "Error"
      End If
   End If
End Sub

Private Sub mnuPrintBug_Click()
   'Prints to default printer...
   
   'Set Printer Defaults
   Printer.FontName = "Courier New"
   Printer.FontSize = 10
   
   'Print Header
   Printer.FontBold = True
   Printer.Print vbTab & String(80, "_")
   Printer.Print ""
   Printer.Print vbTab & "BMG Individual Bug Report"
   Printer.Print vbTab & "Produced: " & Format$(Now, "mm/dd/yyyy")
   Printer.Print vbTab & String(80, "_")
   Printer.Print ""
   Printer.FontBold = False
   
   'Print Data Header
   Printer.FontBold = True
   Printer.Print vbTab & Format$("Bug ID: ", "!" & String(15, "@"));
   Printer.FontBold = False
   Printer.Print txtBugID.Text
   
   Printer.FontBold = True
   Printer.Print vbTab & Format$("Reported On: ", "!" & String(15, "@"));
   Printer.FontBold = False
   Printer.Print txtReported.Text
   
   Printer.FontBold = True
   Printer.Print vbTab & Format$("Reported By: ", "!" & String(15, "@"));
   Printer.FontBold = False
   Printer.Print txtReportedBy.Text
   
   Printer.FontBold = True
   Printer.Print vbTab & Format$("Type: ", "!" & String(15, "@"));
   Printer.FontBold = False
   Printer.Print cboBugType.Text
   
   Printer.FontBold = True
   Printer.Print vbTab & Format$("Severity: ", "!" & String(15, "@"));
   Printer.FontBold = False
   Printer.Print cboSeverity.Text
   
   Printer.FontBold = True
   Printer.Print vbTab & Format$("Status: ", "!" & String(15, "@"));
   Printer.FontBold = False
   Printer.Print cboStatus.Text
   
   Printer.FontBold = True
   Printer.Print vbTab & Format$("Application: ", "!" & String(15, "@"));
   Printer.FontBold = False
   Printer.Print cboApplication.Text
   
   Printer.FontBold = True
   Printer.Print vbTab & Format$("Assigned To: ", "!" & String(15, "@"));
   Printer.FontBold = False
   Printer.Print cboAssignedTo.Text
   
   'Print Data
   Printer.Print ""
   Printer.FontBold = True
   Printer.Print vbTab & "[Steps To Reproduce]"
   Printer.FontBold = False
   Printer.Print vbTab & Replace$(txtStepsToReproduce.Text, vbNewLine, vbNewLine & vbTab)
   
   Printer.Print ""
   Printer.FontBold = True
   Printer.Print vbTab & "[Expected Behaviour]"
   Printer.FontBold = False
   Printer.Print vbTab & Replace$(txtExpectedBehaviour.Text, vbNewLine, vbNewLine & vbTab)
   
   Printer.Print ""
   Printer.FontBold = True
   Printer.Print vbTab & "[Observed Behaviour]"
   Printer.FontBold = False
   Printer.Print vbTab & Replace$(txtObservedBehaviour.Text, vbNewLine, vbNewLine & vbTab)
   
   'Print Footer
   Printer.FontBold = True
   Printer.Print ""
   Printer.Print ""
   Printer.Print vbTab & String(80, "_")
   Printer.Print vbTab & "*** End Of Bug ***"
   Printer.FontBold = False
   
   Printer.EndDoc
   
   MsgBox "Bug Printed", , "Status"
End Sub

Private Sub mnuSaveBug_Click()
   Dim BugSaved As Boolean
   Dim iResp As Integer
   
   If ValidateForm Then
      iResp = MsgBox("Do you wish to save this bug?", vbYesNo, "Confirm")
      
      If iResp = vbYes Then
         CollectBug objBug
         BugSaved = objBug.SaveBug
         
         If BugSaved Then
            MsgBox "Bug successfully saved!", , "Status"
            txtBugID.Text = objBug.FindBugID(7)
         Else
            MsgBox "Bug not saved!", , "Error"
         End If
      End If
   Else
      'This is the only error that can occur at this point
      MsgBox "You must specify a proper date!", , "Error"
      txtReported.SetFocus
   End If
End Sub

Private Sub DisplayBug(pBug As CBug)
   ClearForm
      
   With pBug
      cboApplication.ListIndex = SearchCombo(cboApplication, .Application)
      cboAssignedTo.ListIndex = SearchCombo(cboAssignedTo, .AssignedTo)
      cboBugType.ListIndex = .BugType
      cboSeverity.ListIndex = .Severity
      cboStatus.ListIndex = .Status
      txtBugID.Text = Format$(.ID, String(7, "0"))
      txtExpectedBehaviour.Text = .ExpectedBehaviour
      txtObservedBehaviour.Text = .ObservedBehaviour
      txtReported.Text = Format(.Reported, "mm/dd/yyyy")
      txtReportedBy.Text = .ReportedBy
      txtStepsToReproduce.Text = .StepsToReproduce
   End With
End Sub

Private Sub CollectBug(pBug As CBug)
   Dim objTempStaff As CStaffMember

   Set objTempStaff = New CStaffMember

   With pBug
      If txtBugID.Text <> "" Then .ID = CLng(Val(txtBugID.Text))
      .Application = cboApplication.Text
      .AssignedTo = cboAssignedTo.Text
      .BugType = cboBugType.ListIndex
      .Severity = cboSeverity.ListIndex
      .Status = cboStatus.ListIndex
      .ExpectedBehaviour = txtExpectedBehaviour.Text
      .ObservedBehaviour = txtObservedBehaviour.Text
      .Reported = CDate(txtReported.Text)
      .ReportedBy = txtReportedBy.Text
      .StepsToReproduce = txtStepsToReproduce.Text
   End With

   Set objTempStaff = Nothing
End Sub

Private Function ValidateForm() As Boolean
      'Assume form won't validate
      ValidateForm = False

      'Check items on form
      If Not IsDate(txtReported.Text) Then Exit Function
      
      'We made it through, all is valid
      ValidateForm = True
End Function
