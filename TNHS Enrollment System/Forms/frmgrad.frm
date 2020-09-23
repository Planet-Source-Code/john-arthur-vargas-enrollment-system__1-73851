VERSION 5.00
Begin VB.Form frmgrad 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   LinkTopic       =   "Form4"
   ScaleHeight     =   2760
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1080
      Width           =   1500
   End
   Begin VB.TextBox Text1 
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
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1560
      Width           =   540
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin TNHSES.lvButtons_H cmdenroll 
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Enroll"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Cancel"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Previous SY Average Grade"
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
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Year Graduated"
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
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Student no."
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
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BorderWidth     =   2
      Height          =   2745
      Index           =   0
      Left            =   15
      Top             =   15
      Width           =   3420
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   2880
      Picture         =   "frmgrad.frx":0000
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Graduate Form"
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
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmgrad.frx":2FAD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3435
   End
End
Attribute VB_Name = "frmgrad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdenroll_Click()
If Text1.Text = Clear Then
MsgBox "Input Previous SY Grade first"
Text1.SetFocus
Else
Call dbConnection
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblstudents Where studno='" & txt1.Text & "'"
With RS
If .EOF = False Then
Call dbConnection
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblgradstudent"
RS1.AddNew
For a = 0 To 18
RS1.Fields(a) = RS.Fields(a)
Next a
RS1.Fields(19) = Text1.Text
RS1.Fields(20) = Text2.Text
RS1.Update
RS1.Close
RS.Delete adAffectCurrent
End If
RS.Close
End With
MsgBox "Added to Graduate List"
usrlog ("Added New Graduate with Student No.  '" & txt1.Text & "'")
Unload Me
End If
End Sub
Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text2.Text = sy
txt1.Text = frmviewstudrec.List.SelectedItem.Text
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub

