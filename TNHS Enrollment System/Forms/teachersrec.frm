VERSION 5.00
Begin VB.Form teachersrec 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtsname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtfname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
   Begin VB.ComboBox cmbsubject 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "teachersrec.frx":0000
      Left            =   1920
      List            =   "teachersrec.frx":001C
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
   Begin TNHSES.lvButtons_H Command1 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Save"
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
      cGradient       =   12648447
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454143
   End
   Begin TNHSES.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      cGradient       =   12648447
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454143
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   4080
      Picture         =   "teachersrec.frx":006C
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Surname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject to Teach"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1935
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher's Record"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2925
      Left            =   15
      Top             =   0
      Width           =   4545
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "teachersrec.frx":3019
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10515
   End
End
Attribute VB_Name = "teachersrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If cmbsubject.Text = "Not Set" Then
MsgBox "Subject to Teach is a Required Field"
Else
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from users Where ID='" & txtid.Text & "'"
        With RS
        If .EOF = False Then
         RS.Fields("Sname") = txtsname.Text
         RS.Fields("Fname") = txtfname.Text
        RS.Fields("Subject") = cmbsubject.Text
        MsgBox "Teacher's Record Updated"
        RS.Update
        End If
        End With
        RS.Close
        Unload Me
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Form_Load()
   If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from users Where ID='" & frmmanageuser.txtid.Text & "'"
        With RS
        If .EOF = False Then
        txtsname.Text = RS.Fields("ID")
        txtsname.Text = RS.Fields("Sname")
        txtfname.Text = RS.Fields("Fname")
        cmbsubject.Text = RS.Fields("Subject")
        End If
        End With
        RS.Close
End Sub



Private Sub Image5_Click()
If cmbsubject.Text <> "Not Set" Then
Unload Me
Else
MsgBox "Input Preferred subject to teach"
End If
End Sub

Private Sub lvButtons_H1_Click()
If cmbsubject.Text <> "Not Set" Then
Unload Me
Else
MsgBox "Input Preferred subject to teach"
End If
End Sub

