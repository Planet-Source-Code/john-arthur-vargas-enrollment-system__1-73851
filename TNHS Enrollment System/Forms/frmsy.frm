VERSION 5.00
Begin VB.Form frmsy 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "S"
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   ScaleHeight     =   5130
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox txt12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox txt13 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txt11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3240
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   3840
   End
   Begin TNHSES.lvButtons_H cmdenroll 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   4440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Caption         =   "New School Year"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   12648384
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454016
   End
   Begin TNHSES.lvButtons_H Command3 
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   12648384
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454016
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4455
      Begin VB.TextBox txt1st 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txt3rd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txt2nd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txt1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   0
         Width           =   2295
      End
      Begin VB.TextBox txt4th 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total No. of Sections:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "School Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Sections per Year Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "2nd year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "1st year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         Height          =   1575
         Left            =   0
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "4th year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "3rd year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2160
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Total no of Students Per year Level Limit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   25
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "3rd year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2520
      TabIndex        =   24
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "4th year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2520
      TabIndex        =   23
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "1st year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   600
      TabIndex        =   22
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "2nd year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   600
      TabIndex        =   21
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      Height          =   1575
      Left            =   120
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   4200
      Picture         =   "frmsy.frx":0000
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "TNHS Enrollment System-School Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   6480
      Picture         =   "frmsy.frx":2FAD
      Top             =   60
      Width           =   420
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   5100
      Index           =   0
      Left            =   0
      Top             =   15
      Width           =   4665
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   -120
      Picture         =   "frmsy.frx":5F5A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4875
   End
End
Attribute VB_Name = "frmsy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim displaytotal As Integer
Private Sub cmdenroll_Click()
If cmdenroll.Caption = "New School Year" Then
If userlevel = "Admin" Then
If Year(Now) = sy + 1 Then
If MsgBox("Warning!" & vbNewLine & "Issuing new school year means that all records for schedule will be reset." & vbNewLine & "Are you Sure you Want to proceed?" & vbNewLine & "New School Year: " & sy + 1 & "-" & sy + 2, vbCritical + vbYesNo) = vbYes Then
MsgBox "New School Year Form." & vbNewLine & "Input No. of sections per year level." & vbNewLine & "Click finish to implement effects to new school year"
cmdenroll.Caption = "Finish"
Command3.Caption = "Cancel"
txt1st.Locked = False
txt2nd.Locked = False
txt3rd.Locked = False
txt4th.Locked = False
txt11.Locked = False
txt12.Locked = False
txt13.Locked = False
txt14.Locked = False
txt1st.SetFocus
txt1.Text = sy + 1 & "-" & sy + 2
End If
Else
MsgBox "You Cannot Change school year because the date is not the preceeding school year"
End If
Else
MsgBox "You dont have the rights to access this process"
End If
Else
If displaytotal > 42 Then
MsgBox "The number of sections has been exceeded"
Else
If txt1st.Text < 3 Or txt2nd.Text < 3 Or txt3rd.Text < 3 Or txt4th.Text < 3 Then
MsgBox "Invalid number of section"
Else
Dim msgko As String
If MsgBox("Proceed?", vbYesNo + vbQuestion) = vbYes Then
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblstudents"
Do Until RS.EOF
RS.Fields("EnrollType") = "Regular"
RS.Fields("Evaluated") = "No"
If RS.Fields("status") = "Enrolled" Then
RS.Fields("status") = "Not Enrolled"
ElseIf RS.Fields("status") = "Not Enrolled" Then
RS.Fields("status") = "Dropped"
End If
RS.Update
RS.MoveNext
Loop
RS.Close

Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblprevsubjsched"
Do Until RS.EOF
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblprevsubjsched"
RS1.AddNew
RS1.Fields(0) = RS.Fields(0)
RS1.Fields(1) = RS.Fields(1)
RS1.Fields(2) = RS.Fields(2)
RS1.Fields(3) = RS.Fields(3)
RS1.Fields(4) = RS.Fields(4)
RS1.Fields(5) = RS.Fields(5)
RS1.Update
RS1.Close
RS.Delete adAffectCurrent
RS.MoveNext
Loop
RS.Close

Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblsubjsched"
Do Until RS.EOF
RS.Delete adAffectCurrent
RS.MoveNext
Loop
Dim a As Integer
For a = 1 To txt1st.Text
RS.AddNew
RS.Fields(0) = "1st-" & a
RS.Fields(2) = "Not Set"
RS.Update
Next a
a = 0
For a = 1 To txt2nd.Text
RS.AddNew
RS.Fields(0) = "2nd-" & a
RS.Fields(2) = "Not Set"
RS.Update
Next a
a = 0
For a = 1 To txt3rd.Text
RS.AddNew
RS.Fields(0) = "3rd-" & a
RS.Fields(2) = "Not Set"
RS.Update
Next a
a = 0
For a = 1 To txt4th.Text
RS.AddNew
RS.Fields(0) = "4th-" & a
RS.Fields(2) = "Not Set"
RS.Update
Next a
RS.Close
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from roommap"
Do Until RS1.EOF
RS1.Fields(2) = "vacant"
RS1.Fields(3) = "vacant"
RS1.MoveNext
Loop
RS1.Close
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblsched"
Do Until RS1.EOF
RS1.Delete adAffectCurrent
RS1.MoveNext
Loop
RS1.Close
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblvalsched"
Do Until RS1.EOF
RS1.Delete adAffectCurrent
RS1.MoveNext
Loop
RS1.Close
usrlog ("Changed The School Year of TNHS")
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblyrsection"
RS1.Fields(0) = txt1.Text
RS1.Fields(1) = txt1st.Text
RS1.Fields(2) = txt2nd.Text
RS1.Fields(3) = txt3rd.Text
RS1.Fields(4) = txt4th.Text
RS1.Fields(5) = txt11.Text
RS1.Fields(6) = txt12.Text
RS1.Fields(7) = txt13.Text
RS1.Fields(8) = txt14.Text
RS1.Update
RS1.Close
firstsec = txt1st.Text
secondsec = txt2nd.Text
thirdsec = txt3rd.Text
fourthsec = txt4th.Text
sy = Mid(txt1.Text, 1, 4)
If MsgBox("New school year has been set, Notify Teachers?", vbYesNo + vbQuestion) = vbYes Then
FormGsm.textMessage.Text = "Greetings!" & vbNewLine & "We would like to inform you that new school year has been set."
FormGsm.textMessage.Text = FormGsm.textMessage.Text & vbNewLine & "Please evaluate your handled students if you have advisory class."
FormGsm.textMessage.Text = FormGsm.textMessage.Text & vbNewLine & "Thank you."
FormGsm.Text2.Text = "From: TNHS Administrator"
FormGsm.Command1.Enabled = False
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from users Order by Sname"
Do Until RS.EOF
If RS.Fields("Cnumber") <> Empty Then

FormGsm.ListView1.ListItems.Add , , RS.Fields("fname") & " " & RS.Fields("sname")
FormGsm.ListView1.ListItems.Item(FormGsm.ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("Cnumber")

End If
RS.MoveNext
Loop
RS.Close
End If
FormGsm.Show 1
Unload Me
txt1st.Locked = True
txt2nd.Locked = True
txt3rd.Locked = True
txt4th.Locked = True
txt11.Locked = True
txt12.Locked = True
txt13.Locked = True
txt14.Locked = True

End If
End If
End If
End If
End Sub

Private Sub Txt1st_KeyPress(KeyAscii As Integer)
 If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub
Private Sub Txt2nd_KeyPress(KeyAscii As Integer)
 If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub
Private Sub Txt3rd_KeyPress(KeyAscii As Integer)
 If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub
Private Sub Txt4th_KeyPress(KeyAscii As Integer)
 If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub
Private Sub Command3_Click()
If Command3.Caption = "Cancel" Then
MsgBox "Application for new school year has been cancelled"
cmdenroll.Caption = "New School Year"
Command3.Caption = "Exit"
txt1.Text = sy & "-" & Int(sy + 1)
txt1st.Text = firstsec
txt2nd.Text = secondsec
txt3rd.Text = thirdsec
txt4th.Text = fourthsec
Else
Unload Me
End If
End Sub



Private Sub Form_Load()
txt1.Text = sy & "-" & Int(sy + 1)
txt1st.Text = firstsec
txt2nd.Text = secondsec
txt3rd.Text = thirdsec
txt4th.Text = fourthsec
txt11.Text = firstsec1
txt12.Text = secondsec1
txt13.Text = thirdsec1
txt14.Text = fourthsec1
End Sub

Private Sub Image5_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If txt1st.Text = "" Then txt1st.Text = 0
If txt2nd.Text = "" Then txt2nd.Text = 0
If txt3rd.Text = "" Then txt3rd.Text = 0
If txt4th.Text = "" Then txt4th.Text = 0

displaytotal = Int(txt1st.Text) + Int(txt2nd.Text) + Int(txt3rd.Text) + Int(txt4th.Text)
Label8.Caption = displaytotal & "/42"
End Sub
