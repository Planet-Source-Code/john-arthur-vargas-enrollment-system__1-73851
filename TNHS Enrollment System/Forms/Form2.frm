VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmsecurity 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "TNHS Enrollment System- LOGIN"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6705
   StartUpPosition =   1  'CenterOwner
   Begin TNHSES.lvButtons_H command1 
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Login"
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
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   6240
      Top             =   1800
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   10
      Width           =   6705
      _cx             =   11827
      _cy             =   714
      FlashVars       =   ""
      Movie           =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\login-header.swf"
      Src             =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\login-header.swf"
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1920
      Width           =   2655
   End
   Begin TNHSES.lvButtons_H Command2 
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Exit"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "User Information "
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2895
   End
   Begin VB.Image Image3 
      Height          =   1455
      Left            =   480
      Picture         =   "Form2.frx":164A
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   0
      Picture         =   "Form2.frx":BADD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6705
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   3390
      Left            =   10
      Top             =   0
      Width           =   6690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1740
      Left            =   -600
      Picture         =   "Form2.frx":18043
      Stretch         =   -1  'True
      Top             =   550
      Width           =   7320
   End
End
Attribute VB_Name = "frmsecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Sub Command1_Click()
Call validator
End Sub

Private Sub Command2_Click()
If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion) = vbYes Then
End
End If
End Sub

Private Sub Form_Load()
frmsecurity.Width = 205
ShockwaveFlash1.Movie = App.Path & "\flash\login-header.swf"
ShockwaveFlash1.Play
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call validator
End If
End Sub

Private Sub Timer1_Timer()
If frmsecurity.Width = 6705 Then
Timer1.Enabled = False
Else
frmsecurity.Left = frmsecurity.Left - 250
frmsecurity.Width = frmsecurity.Width + 500
End If
End Sub
Sub Clears()
Text1.Text = Clear
Text2.Text = Clear
Text1.SetFocus
End Sub
Sub validator()
Dim Con As New ADODB.Connection
Dim RS As New ADODB.Recordset
Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\db1.mdb;Persist Security Info=False"
Con.Open
With RS
    .ActiveConnection = Con
    .CursorLocation = 3
    .LockType = 2
End With
RS.Open "select * from users"
RS.Find ("Username='" & Text1.Text & "'")
If RS.EOF = False Then
If RS.Fields("Password") = Text2.Text Then
If RS.Fields("status") = "active" Then
userid = RS.Fields(0)
username = RS.Fields(1)
password = RS.Fields(2)
fname = RS.Fields(4)
sname = RS.Fields(3)
names = RS.Fields(4) & " " & RS.Fields(3)
userlevel = RS.Fields(5)
usrsubject = RS.Fields(6)
address = RS.Fields(7) & " "
number = RS.Fields(8) & " "
RS.Close
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from tblprevsubjsched Where adviserid='" & userid & "'"
valfloater = RS.EOF
If RS.EOF = False Then prevyear = RS.Fields(0)
RS.Close

usrlog ("Logged in to the System")
MsgBox "Logged in"
Unload Me
frmmain.Show
Else
MsgBox "Your account is inactive. Contact system administrator to activate your account."
Call Clears
End If
Else
MsgBox "invalid password"
Call Clears
End If
Else
MsgBox "username does not exist"
Call Clears
End If
End Sub
