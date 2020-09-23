VERSION 5.00
Begin VB.Form frmwebcam 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer vTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin TNHSES.lvButtons_H Command4 
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      Caption         =   "Capture Image"
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
      cBhover         =   12648384
      cGradient       =   12648384
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454016
   End
   Begin TNHSES.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   2760
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
      cBhover         =   12648384
      cGradient       =   12648384
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454016
   End
   Begin TNHSES.lvButtons_H lvButtons_H2 
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "Set As Picture"
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
      cBhover         =   12648384
      cGradient       =   12648384
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454016
   End
   Begin VB.Shape Shape3 
      Height          =   1935
      Left            =   3720
      Top             =   720
      Width           =   2175
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Image"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   4710
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   6440
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   5880
      Picture         =   "frmwebcam.frx":0000
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Image Capture"
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
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmwebcam.frx":2FAD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9195
   End
   Begin VB.Shape Shape2 
      Height          =   3015
      Left            =   240
      Top             =   480
      Width           =   3015
   End
   Begin VB.Image PicWebcam 
      Height          =   3000
      Left            =   240
      Picture         =   "frmwebcam.frx":6C0B
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3000
   End
End
Attribute VB_Name = "frmwebcam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long

Private mCapHwnd As Long
Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054



Private Sub Command4_Click()
MsgBox "Image Captured"
Image4.Picture = PicWebcam.Picture
End Sub

Private Sub Form_Load()
    If viewstudrec = "freshmen" Then
Image4.Picture = frmfreshmen.Image4.Picture
ElseIf viewstudrec = "regular" Then
Image4.Picture = frmregular.Image4.Picture
ElseIf viewstudrec = "user" Then
Image4.Picture = frmmanageuser.Image4.Picture
ElseIf viewstudrec = "user1" Then
Image4.Picture = frmeditinfo.Image4.Picture
Else
Image4.Picture = frmtransferee.Image4.Picture
End If

    mCapHwnd = capCreateCaptureWindow("Picture Capture", 0, 0, 0, 50, 50, Me.hWnd, 0)
    DoEvents
    If capDriverConnect(mCapHwnd, 0) = True Then

        vTimer.Enabled = True
    Else
        MsgBox "Capture device not installed", vbOKOnly, "Capture device Error"
        Command4.Enabled = False
    End If
End Sub

Private Sub Image2_Click()
Unload Me
End Sub
Private Sub lvButtons_H1_Click()
Unload Me
End Sub

Private Sub lvButtons_H2_Click()
If MsgBox("Set as Picture?", vbYesNo + vbQuestion) = vbYes Then
If viewstudrec = "freshmen" Then
frmfreshmen.Image4.Picture = Image4.Picture
ElseIf viewstudrec = "regular" Then
frmregular.Image4.Picture = Image4.Picture
ElseIf viewstudrec = "user" Then
frmmanageuser.Image4.Picture = Image4.Picture
ElseIf viewstudrec = "user1" Then
frmeditinfo.Image4.Picture = Image4.Picture
Else
frmtransferee.Image4.Picture = Image4.Picture
End If
Unload Me
End If
End Sub


Private Sub vTimer_Timer()
    DoEvents
    SendMessage mCapHwnd, GET_FRAME, 0, 0
    SendMessage mCapHwnd, COPY, 0, 0
    PicWebcam.Picture = Clipboard.GetData
    Clipboard.Clear

End Sub
