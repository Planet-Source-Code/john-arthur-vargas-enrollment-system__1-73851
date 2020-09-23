VERSION 5.00
Begin VB.Form frmaddvalsched 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form10"
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8820
   LinkTopic       =   "Form10"
   Moveable        =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optday 
      BackColor       =   &H00400000&
      Caption         =   "Tuesday"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   12
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton optday 
      BackColor       =   &H00400000&
      Caption         =   "Monday"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.OptionButton optday 
      BackColor       =   &H00400000&
      Caption         =   "Wednesday"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin VB.OptionButton optday 
      BackColor       =   &H00400000&
      Caption         =   "Thursday"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton optday 
      BackColor       =   &H00400000&
      Caption         =   "Friday"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4920
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      ItemData        =   "frmaddvalsched.frx":0000
      Left            =   600
      List            =   "frmaddvalsched.frx":000D
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin TNHSES.lvButtons_H Command3 
      Height          =   495
      Left            =   6600
      TabIndex        =   43
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
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
   Begin TNHSES.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   6600
      TabIndex        =   44
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Select Schedule"
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
   Begin VB.Shape Shape3 
      BackColor       =   &H00400000&
      Height          =   2220
      Left            =   120
      Top             =   480
      Width           =   8535
   End
   Begin VB.Label lblclass 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Click Color Bar to view Schedule"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   5760
      TabIndex        =   45
      Top             =   4560
      Width           =   2805
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   4695
      X2              =   4695
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   4695
      X2              =   4695
      Y1              =   4200
      Y2              =   4560
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4350
      X2              =   4350
      Y1              =   4200
      Y2              =   4560
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   4200
      Y2              =   4560
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   5235
      X2              =   5235
      Y1              =   4200
      Y2              =   4560
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   3450
      X2              =   3450
      Y1              =   4200
      Y2              =   4560
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   2565
      X2              =   2565
      Y1              =   4200
      Y2              =   4560
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   7365
      X2              =   7365
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   5580
      X2              =   5580
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   5235
      X2              =   5235
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4350
      X2              =   4350
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   3450
      X2              =   3450
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2565
      X2              =   2565
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   5580
      X2              =   5580
      Y1              =   4200
      Y2              =   4560
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   7365
      X2              =   7365
      Y1              =   4200
      Y2              =   4560
   End
   Begin VB.Label lblaft 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   4
      Left            =   4695
      TabIndex        =   42
      Top             =   4200
      Width           =   525
   End
   Begin VB.Label lblaft 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   5
      Left            =   5580
      TabIndex        =   41
      Top             =   4200
      Width           =   885
   End
   Begin VB.Label lblaft 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   8
      Left            =   5235
      TabIndex        =   40
      Top             =   4200
      Width           =   330
   End
   Begin VB.Label lblaft 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   0
      Left            =   4350
      TabIndex        =   39
      Top             =   4200
      Width           =   345
   End
   Begin VB.Label lblaft 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   7
      Left            =   7365
      TabIndex        =   38
      Top             =   4200
      Width           =   900
   End
   Begin VB.Label lblaft 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   6
      Left            =   6480
      TabIndex        =   37
      Top             =   4200
      Width           =   870
   End
   Begin VB.Label lblaft 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   3
      Left            =   3450
      TabIndex        =   36
      Top             =   4200
      Width           =   885
   End
   Begin VB.Label lblaft 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   2
      Left            =   2565
      TabIndex        =   35
      Top             =   4200
      Width           =   870
   End
   Begin VB.Label lblaft 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   1
      Left            =   1680
      TabIndex        =   34
      Top             =   4200
      Width           =   870
   End
   Begin VB.Label lblmorn 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   4
      Left            =   4695
      TabIndex        =   33
      Top             =   3480
      Width           =   525
   End
   Begin VB.Label lblmorn 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   5
      Left            =   5580
      TabIndex        =   32
      Top             =   3480
      Width           =   885
   End
   Begin VB.Label lblmorn 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   8
      Left            =   5235
      TabIndex        =   31
      Top             =   3480
      Width           =   330
   End
   Begin VB.Label lblmorn 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   0
      Left            =   4350
      TabIndex        =   30
      Top             =   3480
      Width           =   330
   End
   Begin VB.Label lblmorn 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   7
      Left            =   7365
      TabIndex        =   29
      Top             =   3480
      Width           =   900
   End
   Begin VB.Label lblmorn 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   6
      Left            =   6480
      TabIndex        =   28
      Top             =   3480
      Width           =   870
   End
   Begin VB.Label lblmorn 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   3
      Left            =   3450
      TabIndex        =   27
      Top             =   3480
      Width           =   885
   End
   Begin VB.Label lblmorn 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   2
      Left            =   2565
      TabIndex        =   26
      Top             =   3480
      Width           =   870
   End
   Begin VB.Label lblmorn 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   1
      Left            =   1680
      TabIndex        =   25
      Top             =   3480
      Width           =   870
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Left            =   240
      TabIndex        =   24
      Top             =   4200
      Width           =   1260
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Afternoon(pm)"
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
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Label lblclass 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "12:30          1:20             2:10            3:00  3:20    3:50  4:10           5:00              5:50             6:40"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   1560
      TabIndex        =   22
      Top             =   3960
      Width           =   7005
   End
   Begin VB.Label lblclass 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "6:00             6:50           7:40             8:30 8:50     9:20  9:40          10:30           11:20          12:10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   7
      Left            =   1560
      TabIndex        =   21
      Top             =   3240
      Width           =   7005
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Morning(am)"
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
      Index           =   18
      Left            =   240
      TabIndex        =   20
      Top             =   3240
      Width           =   1260
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Left            =   240
      TabIndex        =   19
      Top             =   3480
      Width           =   1260
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   17
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lblaft 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   300
      Index           =   100
      Left            =   5880
      TabIndex        =   16
      Top             =   2400
      Width           =   405
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Vacant"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   15
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblaft 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Height          =   300
      Index           =   1000
      Left            =   7080
      TabIndex        =   14
      Top             =   2400
      Width           =   405
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupied"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Substitute Subject Teacher:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "|Status:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Values Teacher:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject to Substitute"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   8280
      Picture         =   "frmaddvalsched.frx":002D
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Values Schedule"
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BorderWidth     =   2
      Height          =   5030
      Index           =   0
      Left            =   10
      Top             =   10
      Width           =   8810
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   0
      Picture         =   "frmaddvalsched.frx":2FDA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8835
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      Height          =   2220
      Left            =   120
      Top             =   2760
      Width           =   8520
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Picture         =   "frmaddvalsched.frx":6BD0
      Stretch         =   -1  'True
      Top             =   480
      Width           =   8535
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      Height          =   2325
      Left            =   120
      Top             =   480
      Width           =   8535
   End
End
Attribute VB_Name = "frmaddvalsched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim titser As String, day As String, teacher As String, Class As Integer, sessionstr As String

Sub dayloader()
If day = "" Then
MsgBox "Select day"
Else

    For a = 0 To 8
    lblmorn(a).BackColor = &HFFFFC0
    lblaft(a).BackColor = &HFFFFC0
    Next a
    Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
    RS.Open "select * from tblvalsched Where teacher='" & teacher & "'"
    With RS
Do Until .EOF
    If .Fields("day") = day Then
        If .Fields(1) = "Morning" Then
        Select Case .Fields(4)
            Case "1st": lblmorn(1).BackColor = &H8080FF
            Case "2nd": lblmorn(2).BackColor = &H8080FF
            Case "3rd": lblmorn(3).BackColor = &H8080FF
            Case "4th":
            lblmorn(4).BackColor = &H8080FF
                If Right(.Fields(5), 2) = "20" Then
                lblmorn(0).BackColor = &H8080FF
                Else
                lblmorn(8).BackColor = &H8080FF
                End If
            Case "5th": lblmorn(5).BackColor = &H8080FF
            Case "6th": lblmorn(6).BackColor = &H8080FF
            Case "7th": lblmorn(7).BackColor = &H8080FF
            End Select
    Else
           Select Case .Fields(4)
            Case "1st": lblaft(1).BackColor = &H8080FF
            Case "2nd": lblaft(2).BackColor = &H8080FF
            Case "3rd": lblaft(3).BackColor = &H8080FF
            Case "4th":
            lblaft(4).BackColor = &H8080FF
            If Right(.Fields(5), 2) = "50" Then
            lblaft(0).BackColor = &H8080FF
            Else
                lblaft(8).BackColor = &H8080FF
            End If
            Case "5th": lblaft(5).BackColor = &H8080FF
            Case "6th": lblaft(6).BackColor = &H8080FF
            Case "7th": lblaft(7).BackColor = &H8080FF
            End Select
    End If
    End If
.MoveNext
Loop
.Close
End With
Call sessioncheck
Dim availability As Integer
availability = 0
For a = 0 To 8
If lblmorn(a).BackColor = &HFF& Or lblaft(a).BackColor = &HFF& Then availability = 1
Next a
If availability = 1 Then
Label11.Caption = "unavailable"
Else
Label11.Caption = "available"
End If
End If

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
day = ""
optday(0).Enabled = False
optday(1).Enabled = False
optday(2).Enabled = False
optday(3).Enabled = False
optday(4).Enabled = False
If session = 1 Then
sessionstr = "Morning"
Else
sessionstr = "Afternoon"
End If
teacher = frmschedsetup.Combo1.Text
Label3.Caption = "Values teacher:" & teacher
valteacher = teacher
For a = 1 To 2
If a <> indsched Then
If List1.List(2) = frmschedsetup.lblvalsubject(a).Caption Then List1.RemoveItem 2
If List1.List(1) = frmschedsetup.lblvalsubject(a).Caption Then List1.RemoveItem 1
If List1.List(0) = frmschedsetup.lblvalsubject(a).Caption Then List1.RemoveItem 0
End If
Next a
End Sub

Private Sub Image5_Click()
Unload Me
End Sub

Private Sub lblaft_Click(Index As Integer)
If teacher = "" Then
MsgBox "no teacher has been selected"
Else
formboolean = 1
days = day
Select Case Index
Case 1: viewteacherpopclass = "1st"
Case 2: viewteacherpopclass = "2nd"
Case 3: viewteacherpopclass = "3rd"
Case 4: viewteacherpopclass = "4th"
Case 5: viewteacherpopclass = "5th"
Case 6: viewteacherpopclass = "6th"
Case 7: viewteacherpopclass = "7th"
Case 8: viewteacherpopclass = "4th"
Case 0: viewteacherpopclass = "4th"
End Select
viewteacherpopsession = "Afternoon"
frmviewteacherschedpop.Show 1
End If
End Sub

Private Sub lblmorn_Click(Index As Integer)
If teacher = "" Then
MsgBox "no teacher has been selected"
Else
formboolean = 1
days = day
Select Case Index
Case 1: viewteacherpopclass = "1st"
Case 2: viewteacherpopclass = "2nd"
Case 3: viewteacherpopclass = "3rd"
Case 4: viewteacherpopclass = "4th"
Case 5: viewteacherpopclass = "5th"
Case 6: viewteacherpopclass = "6th"
Case 7: viewteacherpopclass = "7th"
Case 0: viewteacherpopclass = "4th"
Case 8: viewteacherpopclass = "4th"
End Select
viewteacherpopsession = "Morning"
frmviewteacherschedpop.Show 1
End If
End Sub

Private Sub List1_Click()
For a = 0 To 7
If frmschedsetup.lblsubject(a).Caption = List1.Text Then
Label7.Caption = frmschedsetup.lblteacher(a).Caption
Label10.Caption = frmschedsetup.lbltime(a).Caption
valsubteacherid(indsched) = teacherid(a)
If Mid(frmschedsetup.lblclass(a).Caption, 1, 3) = "1st" Then classint = 1
If Mid(frmschedsetup.lblclass(a).Caption, 1, 3) = "2nd" Then classint = 2
If Mid(frmschedsetup.lblclass(a).Caption, 1, 3) = "3rd" Then classint = 3
If Mid(frmschedsetup.lblclass(a).Caption, 1, 3) = "4th" Then classint = 4
If Mid(frmschedsetup.lblclass(a).Caption, 1, 3) = "5th" Then classint = 5
If Mid(frmschedsetup.lblclass(a).Caption, 1, 3) = "6th" Then classint = 6
If Mid(frmschedsetup.lblclass(a).Caption, 1, 3) = "7th" Then classint = 7
classes(indsched) = Mid(frmschedsetup.lblclass(a).Caption, 1, 3)
End If
Next a
 For a = 0 To 8
    lblmorn(a).BackColor = &HFFFFC0
    lblaft(a).BackColor = &HFFFFC0
    Next a
Call sessioncheck
Call dayloader
optday(0).Enabled = True
optday(1).Enabled = True
optday(2).Enabled = True
optday(3).Enabled = True
optday(4).Enabled = True
End Sub
Sub sessioncheck()

If sessionstr = "Morning" Then
If lblmorn(classint).BackColor = &H8080FF Then
lblmorn(classint).BackColor = &HFF&
Else
lblmorn(classint).BackColor = &HFFFF00
If Label10.Caption = "8:50-9:40" Then
lblmorn(8).BackColor = &HFFFF00
ElseIf Label10.Caption = "8:30-9:20" Then
lblmorn(0).BackColor = &HFFFF00
End If
End If
Else
If lblaft(classint).BackColor = &H8080FF Then
lblaft(classint).BackColor = &HFF&
Else
lblaft(classint).BackColor = &HFFFF00
If Label10.Caption = "3:20-4:10" Then
lblaft(8).BackColor = &HFFFF00

ElseIf Label10.Caption = "3:00-3:50" Then
lblaft(0).BackColor = &HFFFF00
End If
End If
End If

If lblmorn(4).BackColor = &HFF& Then
    Call dbConnection
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
    RS.Open "select * from tblvalsched Where teacher='" & teacher & "'"
    With RS
        Do Until .EOF
        If .Fields("time") = "8:50-9:40" And .Fields(1) = "Morning" And .Fields("day") = day Then
If Label10.Caption = "8:50-9:40" Then
lblmorn(0).BackColor = &HFFFFC0
lblmorn(8).BackColor = &HFF&
Else
lblmorn(0).BackColor = &HFFFF00
lblmorn(8).BackColor = &H8080FF
End If
ElseIf .Fields("time") = "8:30-9:20" And .Fields(1) = "Morning" Then
If Label10.Caption = "8:30-9:20" And .Fields("day") = day Then
lblmorn(8).BackColor = &HFFFFC0
lblmorn(0).BackColor = &HFF&
Else
lblmorn(8).BackColor = &HFFFF00
lblmorn(0).BackColor = &H8080FF
End If
End If
.MoveNext
Loop
.Close
End With
End If

If lblaft(4).BackColor = &HFF& Then
Call dbConnection
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblvalsched Where teacher='" & teacher & "'"
With RS
Do Until .EOF
If .Fields("time") = "3:20-4:10" And .Fields(1) = "Afternoon" Then
If Label10.Caption = "3:20-4:10" Then
lblaft(0).BackColor = &HFFFFC0
lblaft(8).BackColor = &HFF&
Else
lblaft(0).BackColor = &HFFFF00
lblaft(8).BackColor = &H8080FF
End If
ElseIf .Fields("time") = "3:00-3:50" And .Fields(1) = "Afternoon" Then
If Label10.Caption = "3:00-3:50" Then
lblaft(8).BackColor = &HFFFFC0
lblaft(0).BackColor = &HFF&
Else
lblaft(8).BackColor = &HFFFF00
lblaft(0).BackColor = &H8080FF
End If
End If
.MoveNext
Loop
.Close
End With
End If
End Sub

Private Sub lvButtons_H1_Click()
If Label11.Caption = "" Then
MsgBox "Select Schedule First"
ElseIf Label11.Caption = "unavailable" Then
MsgBox "The Schedule that you have selected is not available. Please Select different schedule"
Else
For a = 1 To 2
If a <> indsched Then
If frmschedsetup.lblvalday(a) = day Then
MsgBox "You cannot select schedule with same day"
Else
frmschedsetup.lblvaltime(indsched).Caption = Label10.Caption
frmschedsetup.lblvalteacher(indsched).Caption = frmschedsetup.Combo1.Text
frmschedsetup.lblvalday(indsched).Caption = day
frmschedsetup.lblvalsubject(indsched).Caption = List1.Text
substeacher(indsched) = Label7
subssubject(indsched) = List1.Text
Unload Me
End If
End If
Next a

End If
End Sub

Private Sub optday_Click(Index As Integer)
If Label7.Caption = "" Then
MsgBox "Select Substitute Subject First"
List1.SetFocus
Else
day = optday(Index).Caption
 Call dayloader
End If
End Sub
