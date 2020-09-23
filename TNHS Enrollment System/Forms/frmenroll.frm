VERSION 5.00
Begin VB.Form frmenroll 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   LinkTopic       =   "Form4"
   ScaleHeight     =   3240
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   9
      Top             =   2040
      Width           =   540
   End
   Begin VB.ComboBox cmb12 
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
      Height          =   360
      Left            =   1320
      TabIndex        =   6
      Text            =   "Select.."
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox cmb11 
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
      Height          =   360
      ItemData        =   "frmenroll.frx":0000
      Left            =   1320
      List            =   "frmenroll.frx":0010
      TabIndex        =   5
      Text            =   "Select.."
      Top             =   1080
      Width           =   2055
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
      Width           =   2055
   End
   Begin TNHSES.lvButtons_H cmdenroll 
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Evaluate"
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
      TabIndex        =   8
      Top             =   2640
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
      TabIndex        =   10
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
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
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BorderWidth     =   2
      Height          =   3230
      Index           =   0
      Left            =   10
      Top             =   10
      Width           =   3425
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   2880
      Picture         =   "frmenroll.frx":0028
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Evaluate Student Form"
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
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmenroll.frx":2FD5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3435
   End
End
Attribute VB_Name = "frmenroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb11_Click()
Call updatesec
End Sub

Private Sub cmdenroll_Click()
If cmb11.Text = "Select.." Or cmb12.Text = "Select.." Or Text1.Text = "" Then
MsgBox "Grade, Year and Section are required fields"
Else

If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblstudents Where studno='" & txt1.Text & "'"
With RS
If .EOF = False Then
    .Fields(0) = txt1.Text
        .Fields(6) = frmregular.txt7.Text
        .Fields(10) = cmb11.Text
        .Fields(11) = cmb12.Text
        '.Fields(13) = "Not Enrolled"
        .Fields("evaluated") = "Yes"
        .Fields("yrsec") = cmb11.Text & "-" & cmb12.Text
        .Fields(19) = Text1.Text
.Update
End If
.Close
End With
      usrlog ("Added to evaluation list of Regular Student with Student Number:" & txt1.Text)

MsgBox "Student Added to Evaluation List"
i_Clear.cLearMe frmregular
frmregular.Image4.Picture = frmmain.imagelist1.ListImages(51).Picture
frmregular.txt6.Text = "__/__/____"
frmregular.txt11.Text = "_________"
frmregular.cmdenroll.Enabled = False
frmregular.Command4.Enabled = False
frmregular.Command3.Caption = "Exit"
Unload Me
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call dbConnection
txt1.Text = frmregular.txt1.Text
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblstudents Where studno='" & txt1.Text & "'"
With RS
If .EOF = False Then
Call updatesec
If .Fields(13) = "Enrolled" Then
cmb11.Text = .Fields(10)
cmb12.Text = .Fields(11)
Call updatesec
Else

If .Fields(10) = "1st" Then
cmb11.Text = "2nd"
ElseIf .Fields(10) = "2nd" Then
cmb11.Text = "3rd"
ElseIf .Fields(10) = "3rd" Then
cmb11.Text = "4th"
End If
Call updatesec
End If
End If
.Close
End With
End Sub
Sub updatesec()
cmb12.Clear
If cmb11.Text = "2nd" Then
For a = 1 To secondsec
cmb12.AddItem a
Next a
ElseIf cmb11.Text = "3rd" Then
For a = 1 To thirdsec
cmb12.AddItem a
Next a
ElseIf cmb11.Text = "4th" Then
For a = 1 To fourthsec
cmb12.AddItem a
Next a
End If
cmb12.Text = "Select.."
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub
