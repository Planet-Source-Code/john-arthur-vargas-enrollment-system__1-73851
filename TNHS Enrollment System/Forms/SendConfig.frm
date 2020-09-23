VERSION 5.00
Begin VB.Form FormSendConfig 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Send Options"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Advanced Send Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      Begin VB.CheckBox checkUDH 
         BackColor       =   &H00400000&
         Caption         =   "User Data Header (UDH)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CheckBox checkMultipart 
         BackColor       =   &H00400000&
         Caption         =   "Allow &multipart messages"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox checkFlash 
         BackColor       =   &H00400000&
         Caption         =   "&Immediate Display"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox checkReport 
         BackColor       =   &H00400000&
         Caption         =   "Request &delivery report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Width           =   2055
      End
   End
   Begin TNHSES.lvButtons_H OkButton 
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   2760
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
      cBhover         =   16777088
      LockHover       =   1
      cGradient       =   16777152
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   16776960
   End
   Begin TNHSES.lvButtons_H CancelButton 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "OK"
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
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   3375
      Left            =   10
      Top             =   10
      Width           =   4485
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Send Options"
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4215
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   3840
      Picture         =   "SendConfig.frx":0000
      Top             =   60
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   -2880
      Picture         =   "SendConfig.frx":2FAD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16155
   End
End
Attribute VB_Name = "FormSendConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
    Me.Hide
End Sub

Private Sub OKButton_Click()
    Me.Hide
End Sub
