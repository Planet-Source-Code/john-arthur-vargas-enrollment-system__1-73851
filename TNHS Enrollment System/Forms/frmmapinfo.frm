VERSION 5.00
Begin VB.Form frmmapinfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form10"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3360
   LinkTopic       =   "Form10"
   ScaleHeight     =   3000
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
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
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin TNHSES.lvButtons_H Command4 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "Select Room"
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
   Begin TNHSES.lvButtons_H Command3 
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2778
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupied By:"
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
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Room No."
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
      Top             =   600
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BorderWidth     =   2
      Height          =   2990
      Index           =   0
      Left            =   10
      Top             =   10
      Width           =   3350
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Information"
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
      Width           =   1575
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   2880
      Picture         =   "frmmapinfo.frx":0000
      Top             =   60
      Width           =   420
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   0
      Picture         =   "frmmapinfo.frx":2FAD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3435
   End
End
Attribute VB_Name = "frmmapinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sessionstr As String
Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If MsgBox("Are you sure you want to select this room?", vbYesNo + vbQuestion) = vbYes Then
frmschedsetup.Text1.Text = txt1.Text
Unload Me
Unload frmmappop
End If
End Sub

Private Sub Form_Load()
Command4.Enabled = True
Call dbConnection
Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from roommap Where ID=" & Int(mapid)
If RS.EOF = False Then
With RS
txt1.Text = .Fields(1)
If session = 1 Then
 If .Fields(2) = "vacant" Then
 txt2.Text = "vacant"
 txt3.Text = Clear
 Else
 txt2.Text = "occupied"
 txt3.Text = .Fields(2)
 Command4.Enabled = False
 End If
Else
 If .Fields(3) = "vacant" Then
 txt2.Text = "vacant"
 txt3.Text = Clear
 Else
 txt2.Text = "occupied"
 txt3.Text = .Fields(3)
 Command4.Enabled = False
 End If
End If
End With
End If
End Sub

Private Sub Image5_Click()
Unload Me
End Sub
