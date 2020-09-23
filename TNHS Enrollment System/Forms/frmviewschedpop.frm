VERSION 5.00
Begin VB.Form frmviewschedpop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form10"
   ClientHeight    =   4440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   LinkTopic       =   "Form10"
   ScaleHeight     =   4440
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Index           =   6
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3360
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
      Index           =   2
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
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
      Index           =   3
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1920
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
      Index           =   1
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
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
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
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
      Index           =   4
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2400
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
      Index           =   5
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2880
      Width           =   2295
   End
   Begin TNHSES.lvButtons_H Command3 
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   3840
      Width           =   1575
      _ExtentX        =   2355
      _ExtentY        =   1085
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Day"
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
      TabIndex        =   14
      Top             =   3360
      Width           =   1455
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
      TabIndex        =   12
      Top             =   105
      Width           =   1695
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   3960
      Picture         =   "frmviewschedpop.frx":0000
      Top             =   60
      Width           =   420
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   4410
      Index           =   1
      Left            =   15
      Top             =   15
      Width           =   4470
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
      TabIndex        =   11
      Top             =   480
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
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
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
      TabIndex        =   9
      Top             =   2400
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
      Top             =   1560
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
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Room"
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
   Begin VB.Image Image6 
      Height          =   375
      Left            =   0
      Picture         =   "frmviewschedpop.frx":2FAD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8715
   End
End
Attribute VB_Name = "frmviewschedpop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2

    If boolpop = 1 Then
    If frmroomsched.Img3(mapid).Picture = frmroomsched.imagelist1(1).ListImages(mapid).Picture Then
    txt1(0).Text = teacher
    txt1(1).Text = "Vacant"
    txt1(2).Text = usrsubject
    txt1(3).Text = pid
    Exit Sub
    End If
    If usrsubject <> "Values Educ." Then
    Label7.Visible = False
    txt1(6).Visible = False
    RS.Open "select * from tblsched Where room='" & pid & "' And teacherid='" & teacherids & "'"
    Else
    RS.Open "select * from tblvalsched Where room='" & pid & "' And teacherid='" & teacherids & "' And day='" & days & "'"
    End If
    ElseIf boolpop = 2 Then
    If sessions = "Morning" Then
    If frmroomsched.lblmorns(mapid).BackColor = &HFFFFC0 Then
    txt1(0).Text = teacher
    txt1(1).Text = "Vacant"
    txt1(2).Text = usrsubject
    Exit Sub
    End If
    Else
    If frmroomsched.lblafts(mapid).BackColor = &HFFFFC0 Then
    txt1(0).Text = teacher
    txt1(1).Text = "Vacant"
    txt1(2).Text = usrsubject
    Exit Sub
     End If
    End If
    If usrsubject <> "Values Educ." Then
       Label7.Visible = False
    txt1(6).Visible = False
    RS.Open "select * from tblsched Where class='" & pid & "' And teacherid='" & teacherids & "'"
    RS.Find ("session='" & sessions & "'")
    Else
    RS.Open "select * from tblvalsched Where class='" & pid & "' And teacherid='" & teacherids & "' And day='" & days & "'"
    RS.Find ("session='" & sessions & "'")
    End If
    End If
If RS.EOF = False Then
        txt1(0).Text = RS.Fields(7)
        txt1(1).Text = "Occupied"
        txt1(2).Text = usrsubject
        txt1(3).Text = RS.Fields(3)
        txt1(4).Text = RS.Fields(0)
        txt1(5).Text = RS.Fields(5)
    If usrsubject = "Values Educ." Then
    txt1(6).Text = RS.Fields(8)
    End If
    
End If
RS.Close
End Sub

Private Sub Image5_Click()
Unload Me
End Sub

