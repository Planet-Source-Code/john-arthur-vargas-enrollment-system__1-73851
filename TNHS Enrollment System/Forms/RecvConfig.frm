VERSION 5.00
Begin VB.Form FormGsmRecvConfig 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   4815
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Advanced Receive Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4575
      Begin VB.CheckBox checkDelete 
         BackColor       =   &H00400000&
         Caption         =   "&Delete messages after receive"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   3000
      End
      Begin VB.ComboBox comboStore 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   360
         TabIndex        =   1
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Message Storage:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   2175
      End
   End
   Begin TNHSES.lvButtons_H OkButton 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
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
      TabIndex        =   5
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
      Height          =   3375
      Left            =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Receive Options"
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
      TabIndex        =   6
      Top             =   120
      Width           =   4215
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   4080
      Picture         =   "RecvConfig.frx":0000
      Top             =   120
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   -4200
      Picture         =   "RecvConfig.frx":2FAD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16155
   End
End
Attribute VB_Name = "FormGsmRecvConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    comboStore.AddItem ("All messages")
    comboStore.AddItem ("SM - SIM Memory")
    comboStore.AddItem ("ME - Device Memory")
    comboStore.AddItem ("MT - SIM & Device Memory")
   
    comboStore.ListIndex = 0
End Sub

Private Sub OKButton_Click()
    Me.Hide
End Sub


