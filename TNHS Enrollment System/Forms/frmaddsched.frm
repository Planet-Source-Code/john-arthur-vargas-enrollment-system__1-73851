VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmaddsched 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   7515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   LinkTopic       =   "Form4"
   Moveable        =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   3165
      ItemData        =   "frmaddsched.frx":0000
      Left            =   240
      List            =   "frmaddsched.frx":0019
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin MSComctlLib.ListView list 
      Height          =   3255
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   11793649
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Teacher ID"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Teacher Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Availability"
         Object.Width           =   1940
      EndProperty
   End
   Begin TNHSES.lvButtons_H lvButtons_H1 
      Height          =   615
      Left            =   4200
      TabIndex        =   35
      Top             =   4560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "Select this teacher"
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
      Height          =   615
      Left            =   6480
      TabIndex        =   34
      Top             =   4560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
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
      cBhover         =   16777088
      LockHover       =   1
      cGradient       =   16777152
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   16776960
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   80
      Left            =   120
      Top             =   5280
      Width           =   8535
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   120
      Left            =   0
      Top             =   360
      Width           =   8535
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   5000
      Left            =   3240
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   120
      Left            =   240
      Top             =   -120
      Width           =   8535
   End
   Begin VB.Shape Shape6 
      Height          =   4830
      Left            =   120
      Top             =   480
      Width           =   8535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Teachers"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subjects"
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
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Picture         =   "frmaddsched.frx":005B
      Stretch         =   -1  'True
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
      Left            =   5880
      TabIndex        =   36
      Top             =   6960
      Width           =   2805
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   7365
      X2              =   7365
      Y1              =   6480
      Y2              =   6840
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   5580
      X2              =   5580
      Y1              =   6480
      Y2              =   6840
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   4695
      X2              =   4695
      Y1              =   5760
      Y2              =   6120
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2565
      X2              =   2565
      Y1              =   5760
      Y2              =   6120
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   3450
      X2              =   3450
      Y1              =   5760
      Y2              =   6120
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4350
      X2              =   4350
      Y1              =   5760
      Y2              =   6120
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   5235
      X2              =   5235
      Y1              =   5760
      Y2              =   6120
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   5580
      X2              =   5580
      Y1              =   5760
      Y2              =   6120
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   5760
      Y2              =   6120
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   7365
      X2              =   7365
      Y1              =   5760
      Y2              =   6120
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   2565
      X2              =   2565
      Y1              =   6480
      Y2              =   6840
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   3450
      X2              =   3450
      Y1              =   6480
      Y2              =   6840
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4350
      X2              =   4350
      Y1              =   6480
      Y2              =   6840
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   4695
      X2              =   4695
      Y1              =   6480
      Y2              =   6840
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   5235
      X2              =   5235
      Y1              =   6480
      Y2              =   6840
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   6480
      Y2              =   6840
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher's Name:"
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
      Left            =   120
      TabIndex        =   15
      Top             =   6960
      Width           =   5655
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
      Left            =   240
      TabIndex        =   14
      Top             =   4920
      Width           =   405
   End
   Begin VB.Label Label5 
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
      Left            =   720
      TabIndex        =   13
      Top             =   4920
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
      Left            =   1560
      TabIndex        =   12
      Top             =   4920
      Width           =   405
   End
   Begin VB.Label Label4 
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
      Left            =   2040
      TabIndex        =   11
      Top             =   4920
      Width           =   1815
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
      TabIndex        =   10
      Top             =   5760
      Width           =   1260
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
      TabIndex        =   9
      Top             =   5520
      Width           =   1260
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
      TabIndex        =   8
      Top             =   5520
      Width           =   7005
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
      TabIndex        =   7
      Top             =   6240
      Width           =   7005
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
      TabIndex        =   6
      Top             =   6240
      Width           =   1260
   End
   Begin VB.Label Label3 
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
      TabIndex        =   5
      Top             =   6480
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   7500
      Index           =   0
      Left            =   15
      Top             =   0
      Width           =   8835
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   8280
      Picture         =   "frmaddsched.frx":058B
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Schedule(Morning Class - 6:00-6:50)"
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
      Left            =   110
      TabIndex        =   0
      Top             =   110
      Width           =   6015
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   0
      Picture         =   "frmaddsched.frx":3538
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8835
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
      TabIndex        =   33
      Top             =   5760
      Width           =   865
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
      TabIndex        =   32
      Top             =   5760
      Width           =   865
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
      TabIndex        =   31
      Top             =   5760
      Width           =   880
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
      TabIndex        =   30
      Top             =   5760
      Width           =   865
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
      Top             =   5760
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
      Index           =   0
      Left            =   4350
      TabIndex        =   28
      Top             =   5760
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
      Index           =   8
      Left            =   5235
      TabIndex        =   27
      Top             =   5760
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
      Index           =   5
      Left            =   5580
      TabIndex        =   26
      Top             =   5760
      Width           =   880
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
      TabIndex        =   25
      Top             =   5760
      Width           =   520
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
      TabIndex        =   24
      Top             =   6480
      Width           =   865
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
      TabIndex        =   23
      Top             =   6480
      Width           =   865
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
      TabIndex        =   22
      Top             =   6480
      Width           =   880
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
      TabIndex        =   21
      Top             =   6480
      Width           =   865
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
      TabIndex        =   20
      Top             =   6480
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
      Index           =   0
      Left            =   4350
      TabIndex        =   19
      Top             =   6480
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
      Index           =   8
      Left            =   5235
      TabIndex        =   18
      Top             =   6480
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
      Index           =   5
      Left            =   5580
      TabIndex        =   17
      Top             =   6480
      Width           =   880
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
      TabIndex        =   16
      Top             =   6480
      Width           =   520
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00400000&
      Height          =   15
      Left            =   120
      Top             =   480
      Width           =   8520
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00400000&
      BorderWidth     =   4
      Height          =   1860
      Left            =   120
      Top             =   5400
      Width           =   8520
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      Height          =   4850
      Left            =   120
      Top             =   480
      Width           =   8535
   End
End
Attribute VB_Name = "frmaddsched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sessionstr As String, lipatclass As String, prof As String, lipattime As String
Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()


teacher = ""
If session = 1 Then
sessionstr = "Morning"
Else
sessionstr = "Afternoon"
End If
lipatclass = frmschedsetup.lblclass(indsched).Caption
Label20.Caption = "Subject Schedule(" & sessionstr & " Class - " & frmschedsetup.lbltime(indsched).Caption & ")"
For a = 0 To 7
If a <> indsched Then
If List1.List(6) = frmschedsetup.lblsubject(a).Caption Then List1.RemoveItem 6
If List1.List(5) = frmschedsetup.lblsubject(a).Caption Then List1.RemoveItem 5
If List1.List(4) = frmschedsetup.lblsubject(a).Caption Then List1.RemoveItem 4
If List1.List(3) = frmschedsetup.lblsubject(a).Caption Then List1.RemoveItem 3
If List1.List(2) = frmschedsetup.lblsubject(a).Caption Then List1.RemoveItem 2
If List1.List(1) = frmschedsetup.lblsubject(a).Caption Then List1.RemoveItem 1
If List1.List(0) = frmschedsetup.lblsubject(a).Caption Then List1.RemoveItem 0
End If
Next a
If sessionstr = "Morning" Then
If lblmorn(classint).BackColor = &H8080FF Then
lblmorn(classint).BackColor = &HFF&
Else
lblmorn(classint).BackColor = &HFFFF00
If frmschedsetup.lbltime(indsched).Caption = "8:50-9:40" Then
lblmorn(8).BackColor = &HFFFF00
ElseIf frmschedsetup.lbltime(indsched).Caption = "8:30-9:20" Then
lblmorn(0).BackColor = &HFFFF00
End If
End If
Else
If lblaft(classint).BackColor = &H8080FF Then
lblaft(classint).BackColor = &HFF&
Else
lblaft(classint).BackColor = &HFFFF00
If frmschedsetup.lbltime(indsched).Caption = "3:20-4:10" Then
lblaft(8).BackColor = &HFFFF00

ElseIf frmschedsetup.lbltime(indsched).Caption = "3:00-3:50" Then
lblaft(0).BackColor = &HFFFF00
End If
End If
End If
End Sub

Private Sub Image5_Click()
Unload Me
End Sub



Private Sub jcbutton1_Click()

End Sub

Private Sub list_Click()
On Error Resume Next
teacher = List.SelectedItem.ListSubItems(1)
Call displaysched
End Sub

Sub displaysched()
Label6.Caption = "Teacher's Name: " & teacher
For a = 0 To 8
lblmorn(a).BackColor = &HFFFFC0
lblaft(a).BackColor = &HFFFFC0
Next a
Call dbConnection
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblsched Where teacher='" & teacher & "'"
If RS1.EOF = False Then
Do Until RS1.EOF
If RS1.Fields(1) = "Morning" Then
Select Case RS1.Fields(4)
Case "1st": lblmorn(1).BackColor = &H8080FF
Case "2nd": lblmorn(2).BackColor = &H8080FF
Case "3rd": lblmorn(3).BackColor = &H8080FF
Case "4th":
lblmorn(4).BackColor = &H8080FF
If Right(RS1.Fields(5), 2) = "20" Then
lblmorn(0).BackColor = &H8080FF
Else
lblmorn(8).BackColor = &H8080FF
End If
Case "5th": lblmorn(5).BackColor = &H8080FF
Case "6th": lblmorn(6).BackColor = &H8080FF
Case "7th": lblmorn(7).BackColor = &H8080FF
End Select
Else
Select Case RS1.Fields(4)
Case "1st": lblaft(1).BackColor = &H8080FF
Case "2nd": lblaft(2).BackColor = &H8080FF
Case "3rd": lblaft(3).BackColor = &H8080FF
Case "4th": lblaft(4).BackColor = &H8080FF
If Right(RS1.Fields(5), 2) = "50" Then
lblaft(0).BackColor = &H8080FF
Else
lblaft(8).BackColor = &H8080FF
End If
Case "5th": lblaft(5).BackColor = &H8080FF
Case "6th": lblaft(6).BackColor = &H8080FF
Case "7th": lblaft(7).BackColor = &H8080FF
End Select
End If
RS1.MoveNext
Loop
End If
RS1.Close
If sessionstr = "Morning" Then
    If lblmorn(classint).BackColor = &H8080FF Then
    lblmorn(classint).BackColor = &HFF&
    Else
    lblmorn(classint).BackColor = &HFFFF00
        If frmschedsetup.lbltime(indsched).Caption = "8:50-9:40" Then
        lblmorn(8).BackColor = &HFFFF00
        ElseIf frmschedsetup.lbltime(indsched).Caption = "8:30-9:20" Then
        lblmorn(0).BackColor = &HFFFF00
        End If
    End If

If lblmorn(4).BackColor = &HFF& Then
Call dbConnection
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblsched Where teacher='" & teacher & "'"
With RS
Do Until .EOF
If .Fields("time") = "8:50-9:40" And .Fields(1) = "Morning" Then
If frmschedsetup.lbltime(indsched).Caption = "8:50-9:40" Then
lblmorn(0).BackColor = &HFFFFC0
lblmorn(8).BackColor = &HFF&
Else
lblmorn(0).BackColor = &HFFFF00
lblmorn(8).BackColor = &H8080FF
End If
ElseIf .Fields("time") = "8:30-9:20" And .Fields(1) = "Morning" Then
If frmschedsetup.lbltime(indsched).Caption = "8:30-9:20" Then
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

Else
If lblaft(classint).BackColor = &H8080FF Then
lblaft(classint).BackColor = &HFF&
Else
lblaft(classint).BackColor = &HFFFF00
If frmschedsetup.lbltime(indsched).Caption = "3:20-4:10" Then
lblaft(8).BackColor = &HFFFF00

ElseIf frmschedsetup.lbltime(indsched).Caption = "3:00-3:50" Then
lblaft(0).BackColor = &HFFFF00

End If
End If

If lblaft(4).BackColor = &HFF& Then
Call dbConnection
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblsched Where teacher='" & teacher & "'"
With RS
Do Until .EOF
If .Fields("time") = "3:20-4:10" And .Fields(1) = "Afternoon" Then
If frmschedsetup.lbltime(indsched).Caption = "3:20-4:10" Then
lblaft(0).BackColor = &HFFFFC0
lblaft(8).BackColor = &HFF&
Else
lblaft(0).BackColor = &HFFFF00
lblaft(8).BackColor = &H8080FF
End If
ElseIf .Fields("time") = "3:00-3:50" And .Fields(1) = "Afternoon" Then
If frmschedsetup.lbltime(indsched).Caption = "3:00-3:50" Then
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
End If

End Sub
Private Sub list_DblClick()
teacher = List.SelectedItem.Text
Call displaysched
End Sub
Private Sub lblaft_Click(Index As Integer)
If teacher = "" Then
MsgBox "no teacher has been selected"
Else
formboolean = 0
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
formboolean = 0
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
List.ListItems.Clear
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from users Where Subject='" & List1.Text & "'"
With RS
Do Until .EOF
List.ListItems.Add , , .Fields(0)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(4) & " " & .Fields(3)
prof = .Fields(4) & " " & .Fields(3)
Dim ab As Integer
ab = 0
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblsched Where teacher='" & prof & "'"
If RS1.EOF = False Then
Do Until RS1.EOF
If sessionstr = RS1.Fields(1) And lipatclass = RS1.Fields(4) Then
ab = 1
End If
RS1.MoveNext
Loop
End If
RS1.Close
If ab = 1 Then
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "unavailable"
Else
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "available"
End If
.MoveNext
Loop


.Close
End With

End Sub


Private Sub lvButtons_H1_Click()
If teacher = "" Then
MsgBox "No teacher has been selected"
Else
For a = 1 To 7
If List.SelectedItem.SubItems(2) = "unavailable" Then
MsgBox "The teacher you have selected is not available. Please select other teacher"
Exit Sub
Else
End If
If session = 1 Then
If lblaft(a).BackColor = &H8080FF Then
If MsgBox("The teacher you have selecter already has schedule in afternoon." & vbNewLine & "This may have conflict in the schedule of the selected teacher." & vbNewLine & "Are you sure you want to schedule this teacher in this time?", vbYesNo + vbQuestion) = vbYes Then
frmschedsetup.lblsubject(indsched).Caption = frmaddsched.List1.Text
frmschedsetup.lblteacher(indsched).Caption = teacher
teacherid(indsched) = List.SelectedItem.Text
Unload Me
Unload frmaddsched
Exit Sub
Else
Exit Sub
End If
End If
Else
If lblmorn(a).BackColor = &H8080FF Then
If MsgBox("The teacher you have selecter already has schedule in morning." & vbNewLine & "This may have conflict in the schedule of the selected teacher." & vbNewLine & "Are you sure you want to schedule this teacher in this time?", vbYesNo + vbQuestion) = vbYes Then
frmschedsetup.lblsubject(indsched).Caption = frmaddsched.List1.Text
frmschedsetup.lblteacher(indsched).Caption = teacher
teacherid(indsched) = List.SelectedItem.Text
Unload Me
Unload frmaddsched
Exit Sub
Else
Exit Sub
End If
End If
End If
Next a
If MsgBox("Are you sure you want to schedule this teacher in this time?", vbYesNo + vbQuestion) = vbYes Then
frmschedsetup.lblsubject(indsched).Caption = frmaddsched.List1.Text
frmschedsetup.lblteacher(indsched).Caption = teacher
teacherid(indsched) = List.SelectedItem.Text
Unload Me
Unload frmaddsched
Else
Exit Sub
End If
End If
End Sub
