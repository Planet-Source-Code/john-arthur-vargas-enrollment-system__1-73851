VERSION 5.00
Begin VB.Form frmviewteacherschedpop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form10"
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   LinkTopic       =   "Form10"
   ScaleHeight     =   3945
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt6 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txt5 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txt2 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txt4 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txt3 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin TNHSES.lvButtons_H Command3 
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   3360
      Width           =   1455
      _ExtentX        =   2566
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
      cBhover         =   16777088
      LockHover       =   1
      cGradient       =   16777152
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   16776960
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Section to Teach"
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
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
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
      TabIndex        =   10
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
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
      TabIndex        =   8
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
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
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher's Name"
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
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   3930
      Index           =   0
      Left            =   15
      Top             =   15
      Width           =   4470
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   3960
      Picture         =   "frmvewteacherschedpop.frx":0000
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher's Schedule"
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
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   0
      Picture         =   "frmvewteacherschedpop.frx":2FAD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8715
   End
End
Attribute VB_Name = "frmviewteacherschedpop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
If formboolean = 0 Then
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblsched Where teacher='" & teacher & "'"
If RS1.EOF = False Then
Do Until RS1.EOF
txt1.Text = teacher
txt2.Text = RS1.Fields("subject")
If viewteacherpopsession = RS1.Fields(1) And viewteacherpopclass = RS1.Fields(4) Then
txt2.Text = RS1.Fields("subject")
txt3.Text = "Occupied"
txt4.Text = RS1.Fields(4)
txt5.Text = RS1.Fields(0)
txt6.Text = RS1.Fields(5)
Exit Sub
Else
txt3.Text = "Vacant"
End If
RS1.MoveNext
Loop
Else
txt1.Text = teacher
txt2.Text = frmaddsched.List1.Text
txt3.Text = "Vacant"
End If
RS1.Close
Else
Label20.Caption = "Values Teacher's Schedule"
Label3.Caption = "Day"
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblvalsched Where teacher='" & valteacher & "'"
If RS1.EOF = False Then
Do Until RS1.EOF
txt1.Text = valteacher
txt2.Text = "Values"
If viewteacherpopsession = RS1.Fields(1) And viewteacherpopclass = RS1.Fields(4) And days = RS1.Fields("day") Then
txt2.Text = "Values"
txt3.Text = "Occupied"
txt4.Text = RS1.Fields("day")
txt5.Text = RS1.Fields(0)
txt6.Text = RS1.Fields(5)
Exit Sub
Else
txt3.Text = "Vacant"
End If
RS1.MoveNext
Loop
Else
txt1.Text = teacher
txt2.Text = "Values"
txt3.Text = "Vacant"
End If
RS1.Close
End If
End Sub

Private Sub Image5_Click()
Unload Me
End Sub
