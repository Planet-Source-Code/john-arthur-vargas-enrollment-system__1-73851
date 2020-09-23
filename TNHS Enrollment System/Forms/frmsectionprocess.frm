VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmschedsetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TNHS Enrollment System- Record Schedule Setup"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmsectionprocess.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picstep4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   7695
      TabIndex        =   76
      Top             =   720
      Width           =   7695
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2280
         TabIndex        =   110
         Text            =   "Select Teacher"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.CommandButton cmdvalues 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   6660
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   2520
         Width           =   300
      End
      Begin VB.CommandButton cmdvalues 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   6660
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   2880
         Width           =   300
      End
      Begin TNHSES.lvButtons_H lvButtons_H3 
         Height          =   405
         Left            =   3000
         TabIndex        =   121
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   3
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin TNHSES.lvButtons_H lvButtons_H5 
         Height          =   405
         Left            =   4560
         TabIndex        =   122
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "&Back"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   3
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin TNHSES.lvButtons_H lvButtons_H6 
         Height          =   405
         Left            =   6120
         TabIndex        =   123
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "&Next"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   3
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "is allowed to be a substitute In Values Educ. Subject"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   2880
         TabIndex        =   107
         Top             =   4320
         Width           =   3270
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Only Subject that is included in MAKABAYAN(TLE, MAPEH, Social Studies)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   2520
         TabIndex        =   106
         Top             =   4080
         Width           =   5160
      End
      Begin VB.Label lblvaltime 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Not Set"
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
         Height          =   300
         Index           =   1
         Left            =   2580
         TabIndex        =   102
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Time"
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
         Height          =   300
         Index           =   5
         Left            =   2580
         TabIndex        =   96
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Subs. Subject"
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
         Height          =   300
         Left            =   1200
         TabIndex        =   95
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   4
         Left            =   6660
         TabIndex        =   93
         Top             =   2160
         Width           =   300
      End
      Begin VB.Label lblvalteacher 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Not Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   4800
         TabIndex        =   105
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label lblvalteacher 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Not Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   2
         Left            =   4800
         TabIndex        =   104
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label lblvalday 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Not Set"
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
         Height          =   300
         Index           =   2
         Left            =   3660
         TabIndex        =   103
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblvaltime 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Not Set"
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
         Height          =   300
         Index           =   2
         Left            =   2580
         TabIndex        =   101
         Top             =   2880
         Width           =   1020
      End
      Begin VB.Label lblvalsubject 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Not Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   2
         Left            =   1200
         TabIndex        =   100
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblvalsubject 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Not Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   99
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Teacher"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   6
         Left            =   4800
         TabIndex        =   98
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3660
         TabIndex        =   97
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblvalday 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Not Set"
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
         Height          =   300
         Index           =   1
         Left            =   3660
         TabIndex        =   94
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Image Image10 
         Height          =   675
         Left            =   0
         Picture         =   "frmsectionprocess.frx":164A
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   7725
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teacher"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   960
         TabIndex        =   92
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[1]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Index           =   3
         Left            =   840
         TabIndex        =   83
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         TabIndex        =   82
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[2]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   1320
         TabIndex        =   81
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[3]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   1800
         TabIndex        =   80
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[4]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2280
         TabIndex        =   79
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Schedule for Values Educ."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1320
         TabIndex        =   78
         Top             =   600
         Width           =   4245
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[5]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   2760
         TabIndex        =   77
         Top             =   120
         Width           =   360
      End
      Begin VB.Line Line6 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         X1              =   -120
         X2              =   7680
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   1275
         Left            =   1080
         Shape           =   4  'Rounded Rectangle
         Top             =   2040
         Width           =   6015
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         Height          =   1515
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   6255
      End
      Begin VB.Image Image11 
         Height          =   4545
         Left            =   -120
         Picture         =   "frmsectionprocess.frx":477B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4485
      End
      Begin VB.Shape Shape7 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5295
         Left            =   2520
         Top             =   0
         Width           =   3495
      End
   End
   Begin VB.PictureBox picstep5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   7695
      TabIndex        =   60
      Top             =   720
      Width           =   7695
      Begin MSComctlLib.ListView list 
         Height          =   2895
         Left            =   2280
         TabIndex        =   111
         Top             =   1200
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   11793649
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Teacher Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Subject"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Availability"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "User ID"
            Object.Width           =   2
         EndProperty
      End
      Begin TNHSES.lvButtons_H lvButtons_H2 
         Height          =   405
         Left            =   2880
         TabIndex        =   124
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   3
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin TNHSES.lvButtons_H cmdback4 
         Height          =   405
         Left            =   4440
         TabIndex        =   125
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "&Back"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   3
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin TNHSES.lvButtons_H cmd4 
         Height          =   405
         Left            =   6000
         TabIndex        =   126
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "&Finish"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   3
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[5]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2760
         TabIndex        =   75
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Image8 
         Height          =   675
         Left            =   0
         Picture         =   "frmsectionprocess.frx":C31B
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   7725
      End
      Begin VB.Line Line5 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         X1              =   -120
         X2              =   7680
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Class Adviser"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1440
         TabIndex        =   66
         Top             =   720
         Width           =   2865
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[4]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   2280
         TabIndex        =   65
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[3]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   1800
         TabIndex        =   64
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[2]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   1320
         TabIndex        =   63
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[1]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Index           =   1
         Left            =   840
         TabIndex        =   61
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Image9 
         Height          =   4545
         Left            =   -120
         Picture         =   "frmsectionprocess.frx":F44C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4485
      End
      Begin VB.Shape Shape6 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5295
         Left            =   3480
         Top             =   1320
         Width           =   3495
      End
   End
   Begin VB.PictureBox picstep3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   7695
      TabIndex        =   17
      Top             =   720
      Width           =   7695
      Begin TNHSES.lvButtons_H lvButtons_H4 
         Height          =   405
         Left            =   3000
         TabIndex        =   113
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin TNHSES.lvButtons_H cmdback3 
         Height          =   405
         Left            =   4560
         TabIndex        =   114
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "&Back"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin TNHSES.lvButtons_H btn3 
         Height          =   405
         Left            =   6120
         TabIndex        =   115
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "&Next"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3435
         Left            =   720
         TabIndex        =   24
         Top             =   720
         Width           =   6375
         Begin VB.CommandButton cmdsubjsched 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   7
            Left            =   5940
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   3000
            Width           =   300
         End
         Begin VB.CommandButton cmdsubjsched 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   6
            Left            =   5940
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   2640
            Width           =   300
         End
         Begin VB.CommandButton cmdsubjsched 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   5940
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   2280
            Width           =   300
         End
         Begin VB.CommandButton cmdsubjsched 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   5940
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   1920
            Width           =   300
         End
         Begin VB.CommandButton cmdsubjsched 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   5940
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   1560
            Width           =   300
         End
         Begin VB.CommandButton cmdsubjsched 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   5940
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   1200
            Width           =   300
         End
         Begin VB.CommandButton cmdsubjsched 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   5940
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   840
            Width           =   300
         End
         Begin VB.CommandButton cmdsubjsched 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   5940
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   5940
            TabIndex        =   68
            Top             =   120
            Width           =   300
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
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
            Height          =   300
            Index           =   0
            Left            =   2460
            TabIndex        =   67
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Class"
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
            Height          =   300
            Left            =   120
            TabIndex        =   59
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Time"
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
            Height          =   300
            Index           =   18
            Left            =   1260
            TabIndex        =   58
            Top             =   120
            Width           =   1145
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Subject"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2460
            TabIndex        =   57
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Teacher"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   17
            Left            =   3960
            TabIndex        =   56
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lblclass 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "7th"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   7
            Left            =   120
            TabIndex        =   55
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label lblclass 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "5th"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   5
            Left            =   120
            TabIndex        =   54
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label lblclass 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Break Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   4
            Left            =   120
            TabIndex        =   53
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label lblclass 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "1st"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   52
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblclass 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "2nd"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblclass 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "3rd"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   50
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label lblclass 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "4th"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   49
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label lblclass 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "6th"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   6
            Left            =   120
            TabIndex        =   48
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "6:50-7:40"
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
            Height          =   300
            Index           =   1
            Left            =   1260
            TabIndex        =   47
            Top             =   840
            Width           =   1145
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "6:00-6:50"
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
            Height          =   300
            Index           =   0
            Left            =   1260
            TabIndex        =   46
            Top             =   480
            Width           =   1145
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "7:40-8:30"
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
            Height          =   300
            Index           =   2
            Left            =   1260
            TabIndex        =   45
            Top             =   1200
            Width           =   1145
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "8:30-9:20"
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
            Height          =   300
            Index           =   3
            Left            =   1260
            TabIndex        =   44
            Top             =   1560
            Width           =   1145
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "9:20-9:40"
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
            Height          =   300
            Index           =   4
            Left            =   1260
            TabIndex        =   43
            Top             =   1920
            Width           =   1145
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "9:40-10:30"
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
            Height          =   300
            Index           =   5
            Left            =   1260
            TabIndex        =   42
            Top             =   2280
            Width           =   1145
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "10:30-11:20"
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
            Height          =   300
            Index           =   6
            Left            =   1260
            TabIndex        =   41
            Top             =   2640
            Width           =   1145
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "11:20-12:10"
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
            Height          =   300
            Index           =   7
            Left            =   1260
            TabIndex        =   40
            Top             =   3000
            Width           =   1145
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
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
            Height          =   300
            Index           =   7
            Left            =   2460
            TabIndex        =   39
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
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
            Height          =   300
            Index           =   5
            Left            =   2460
            TabIndex        =   38
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00FFFFC0&
            Caption         =   "-------"
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
            Height          =   300
            Index           =   4
            Left            =   2460
            TabIndex        =   37
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
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
            Height          =   300
            Index           =   1
            Left            =   2460
            TabIndex        =   36
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
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
            Height          =   300
            Index           =   2
            Left            =   2460
            TabIndex        =   35
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
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
            Height          =   300
            Index           =   3
            Left            =   2460
            TabIndex        =   34
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
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
            Height          =   300
            Index           =   6
            Left            =   2460
            TabIndex        =   33
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label lblteacher 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   3960
            TabIndex        =   32
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label lblteacher 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   0
            Left            =   3960
            TabIndex        =   31
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lblteacher 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   3960
            TabIndex        =   30
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label lblteacher 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   3960
            TabIndex        =   29
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label lblteacher 
            BackColor       =   &H00FFFFC0&
            Caption         =   "-------"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   4
            Left            =   3960
            TabIndex        =   28
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label lblteacher 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   5
            Left            =   3960
            TabIndex        =   27
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label lblteacher 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   6
            Left            =   3960
            TabIndex        =   26
            Top             =   2640
            Width           =   1935
         End
         Begin VB.Label lblteacher 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Not Set"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   7
            Left            =   3960
            TabIndex        =   25
            Top             =   3000
            Width           =   1935
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   3315
            Left            =   60
            Top             =   60
            Width           =   6255
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00C0C000&
            BackStyle       =   1  'Opaque
            Height          =   3435
            Left            =   0
            Top             =   0
            Width           =   6375
         End
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[5]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   2760
         TabIndex        =   73
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Image7 
         Height          =   675
         Left            =   0
         Picture         =   "frmsectionprocess.frx":16FEC
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   7725
      End
      Begin VB.Line Line4 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         X1              =   -120
         X2              =   7680
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fill up Schedule"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   3480
         TabIndex        =   23
         Top             =   120
         Width           =   2265
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[4]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   2280
         TabIndex        =   22
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[3]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1800
         TabIndex        =   21
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[2]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   1320
         TabIndex        =   20
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[1]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   840
         TabIndex        =   18
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Image6 
         Height          =   4545
         Left            =   -120
         Picture         =   "frmsectionprocess.frx":1A11D
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4335
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   4575
         Left            =   4440
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.PictureBox picstep2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   7695
      TabIndex        =   10
      Top             =   720
      Width           =   7695
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   2400
         Width           =   1935
      End
      Begin TNHSES.lvButtons_H lvButtons_H1 
         Height          =   405
         Left            =   3000
         TabIndex        =   116
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   3
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin TNHSES.lvButtons_H cmd 
         Height          =   405
         Left            =   4560
         TabIndex        =   117
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "&Back"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   3
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin TNHSES.lvButtons_H cmd2 
         Height          =   405
         Left            =   6120
         TabIndex        =   118
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "&Next"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   3
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[5]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   2760
         TabIndex        =   72
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   600
         Left            =   6720
         Picture         =   "frmsectionprocess.frx":21CBD
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   615
      End
      Begin VB.Image Image4 
         Height          =   675
         Left            =   0
         Picture         =   "frmsectionprocess.frx":23C5B
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   7725
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[1]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Index           =   0
         Left            =   840
         TabIndex        =   16
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[2]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1320
         TabIndex        =   14
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[3]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   1800
         TabIndex        =   13
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[4]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   2280
         TabIndex        =   12
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Room"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   1800
      End
      Begin VB.Line Line3 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         X1              =   -120
         X2              =   7680
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Image Image5 
         Height          =   4545
         Left            =   0
         Picture         =   "frmsectionprocess.frx":26D8C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4245
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5295
         Left            =   4200
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.PictureBox picstep1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   7695
      TabIndex        =   1
      Top             =   720
      Width           =   7695
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Afternoon Class"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   9
         Top             =   2640
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Morning Class"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   8
         Top             =   1920
         Width           =   2775
      End
      Begin TNHSES.lvButtons_H cmdcancel 
         Height          =   405
         Left            =   4560
         TabIndex        =   119
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   3
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin TNHSES.lvButtons_H cmd1 
         Height          =   405
         Left            =   6120
         TabIndex        =   120
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "Next"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   3
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[5]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   2760
         TabIndex        =   74
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(12:30pm-6:40pm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4680
         TabIndex        =   70
         Top             =   3120
         Width           =   1830
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(6:00am-12:10pm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4680
         TabIndex        =   69
         Top             =   2400
         Width           =   1830
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Class Session"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1800
         TabIndex        =   7
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[4]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   2280
         TabIndex        =   6
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[3]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   1800
         TabIndex        =   5
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[2]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[1]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   840
         TabIndex        =   2
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   675
         Left            =   0
         Picture         =   "frmsectionprocess.frx":2E92C
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   7725
      End
      Begin VB.Line Line2 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         X1              =   -120
         X2              =   7680
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Image imgL1Bg 
         Height          =   4545
         Left            =   0
         Picture         =   "frmsectionprocess.frx":31A5D
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4485
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5295
         Left            =   4200
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3360
      TabIndex        =   112
      Top             =   120
      Width           =   2130
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   0
      X2              =   7680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record Schedule For:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   3420
   End
   Begin VB.Image Image3 
      Height          =   675
      Left            =   0
      Picture         =   "frmsectionprocess.frx":395FD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7725
   End
End
Attribute VB_Name = "frmschedsetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adviser As String, tich() As String
Private Sub btn3_Click()
For a = 0 To 7
If lblsubject(a).Caption = "Not Set" Or lblteacher(a).Caption = "Not Set" Then
MsgBox "complete schedule first before you proceed to the next step"
Exit Sub
End If
Next a
picstep1.Visible = False
picstep2.Visible = False
picstep3.Visible = False
picstep4.Visible = True
   For a = 1 To 2
    lblvalsubject(a).Caption = "Not Set"
    lblvalteacher(a).Caption = "Not Set"
    lblvaltime(a).Caption = "Not Set"
    lblvalday(a).Caption = "Not Set"
Next a
End Sub

Private Sub cmd_Click()
picstep2.Visible = False
picstep3.Visible = False
picstep4.Visible = False
picstep1.Visible = True
End Sub

Private Sub cmd1_Click()
If Option1.Value = True Or Option2.Value = True Then
picstep1.Visible = False
picstep3.Visible = False
picstep4.Visible = False
picstep2.Visible = True
cmd2.SetFocus
Text1.Text = Clear
Else
MsgBox "Select Class Session First"
End If
End Sub

Private Sub cmd4_Click()
Dim usrid As String

If List.SelectedItem.ListSubItems(2) = "available" Then
If MsgBox("Are you sure you want to select this teacher?", vbQuestion + vbYesNo) = vbYes Then
Dim sessionstr As String
If Option1.Value = True Then
sessionstr = "Morning"
Else
sessionstr = "Afternoon"
End If
Call dbConnection
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblsubjsched"
With RS1
RS1.Find ("section='" & Label33.Caption & "'")
If .EOF = False Then
If Option1.Value = True Then
.Fields(1) = sessionstr
Else
.Fields(1) = sessionstr
End If
.Fields(2) = "Finished"
.Fields(3) = Text1.Text
.Fields(4) = List.SelectedItem.Text
.Fields(5) = List.SelectedItem.SubItems(3)
End If
.Update
.Close
End With
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from roommap Where roomname='" & Text1.Text & "'"
If RS.EOF = False Then
If Option1.Value = True Then
RS.Fields(2) = frmsubsched.Text2.Text
Else
RS.Fields(3) = frmsubsched.Text2.Text
End If
RS.Update
End If
RS.Close
For a = 0 To 7
If cmdsubjsched(a).Caption <> "B" Then
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblsched"
RS.AddNew
RS.Fields(0) = frmsubsched.Text2.Text
RS.Fields(1) = sessionstr
RS.Fields(2) = "Finished"
RS.Fields(3) = Text1.Text
RS.Fields(4) = lblclass(a).Caption
RS.Fields(5) = lbltime(a).Caption
RS.Fields(6) = lblsubject(a).Caption
RS.Fields(7) = lblteacher(a).Caption
RS.Fields(8) = teacherid(a)
RS.Update
RS.Close
End If
Next a

For a = 1 To 2
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblvalsched"
RS.AddNew
RS.Fields(0) = frmsubsched.Text2.Text
RS.Fields(1) = sessionstr
RS.Fields(2) = "Finished"
RS.Fields(3) = Text1.Text
RS.Fields(4) = classes(a)
RS.Fields(5) = lblvaltime(a).Caption
RS.Fields(6) = subssubject(a)
RS.Fields(7) = lblvalteacher(a).Caption
RS.Fields(8) = lblvalday(a).Caption
RS.Fields(9) = substeacher(a)
RS.Fields(11) = valteacherid(a)
RS.Fields(12) = valsubteacherid(a)
Dim X As Integer
Select Case lblvalday(a).Caption
Case "Monday": X = 1
Case "Tuesday": X = 2
Case "Wednesday": X = 3
Case "Thursday": X = 4
Case "Friday": X = 5
End Select
RS.Fields(10) = X
RS.Update
RS.Close
Next a
usrlog ("Set the schedule of section " & frmsubsched.Text2.Text)
MsgBox "Setup Has Successfully Finished!"
Call loadyrsec
Unload Me
End If
Else
MsgBox "The teacher you have selected already has an advisory class"
End If
End Sub

Private Sub cmdback3_Click()
picstep1.Visible = False
picstep4.Visible = False
picstep3.Visible = False
picstep2.Visible = True
End Sub

Private Sub cmdback4_Click()
picstep1.Visible = False
picstep2.Visible = False
picstep5.Visible = False
picstep4.Visible = True
End Sub

Private Sub cmdCancel_Click()
Call cancelselect
End Sub


Private Sub cmdsubjsched_Click(Index As Integer)
If cmdsubjsched(Index).Caption <> "B" Then
If lblteacher(Index).Caption <> "Not Set" Then
    If MsgBox("The schedule you selected is already set." & vbNewLine & " Do you want to unset this schedule?", vbYesNo + vbQuestion) = vbYes Then
    lblsubject(Index).Caption = "Not Set"
    lblteacher(Index).Caption = "Not Set"
    Else
indsched = Index
If Mid(lblclass(Index).Caption, 1, 3) = "1st" Then classint = 1
If Mid(lblclass(Index).Caption, 1, 3) = "2nd" Then classint = 2
If Mid(lblclass(Index).Caption, 1, 3) = "3rd" Then classint = 3
If Mid(lblclass(Index).Caption, 1, 3) = "4th" Then classint = 4
If Mid(lblclass(Index).Caption, 1, 3) = "5th" Then classint = 5
If Mid(lblclass(Index).Caption, 1, 3) = "6th" Then classint = 6
If Mid(lblclass(Index).Caption, 1, 3) = "7th" Then classint = 7
    frmaddsched.Show 1
    End If
Else
indsched = Index
If Mid(lblclass(Index).Caption, 1, 3) = "1st" Then classint = 1
If Mid(lblclass(Index).Caption, 1, 3) = "2nd" Then classint = 2
If Mid(lblclass(Index).Caption, 1, 3) = "3rd" Then classint = 3
If Mid(lblclass(Index).Caption, 1, 3) = "4th" Then classint = 4
If Mid(lblclass(Index).Caption, 1, 3) = "5th" Then classint = 5
If Mid(lblclass(Index).Caption, 1, 3) = "6th" Then classint = 6
If Mid(lblclass(Index).Caption, 1, 3) = "7th" Then classint = 7
frmaddsched.Show 1
End If
End If
End Sub

Private Sub cmdvalues_Click(Index As Integer)
If Combo1.Text = "Select Teacher" Then
MsgBox "Select Teacher first"
Else
If lblvalsubject(Index).Caption <> "Not Set" Then
    If MsgBox("The schedule you selected is already set." & vbNewLine & " Do you want to unset this schedule?", vbYesNo + vbQuestion) = vbYes Then
    lblvalsubject(Index).Caption = "Not Set"
    lblvalteacher(Index).Caption = "Not Set"
    lblvaltime(Index).Caption = "Not Set"
    lblvalday(Index).Caption = "Not Set"
    End If
    Else
indsched = Index
frmaddvalsched.Show 1
End If

End If
End Sub
Private Sub Combo1_Click()
For a = 1 To 2
lblvalsubject(a).Caption = "Not Set"
lblvaltime(a).Caption = "Not Set"
lblvalday(a).Caption = "Not Set"
lblvalteacher(a).Caption = Combo1.Text
valteacherid(a) = tich(Combo1.ListIndex)
Next a
End Sub

Private Sub Form_Load()
Call dbConnection
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from users Where Subject='Values Educ.'"
With RS
ReDim tich(RS.RecordCount) As String
Dim ab As Integer
b = 0
Do Until .EOF
Combo1.AddItem RS.Fields(4) & " " & RS.Fields(3)
tich(ab) = RS.Fields(0)
ab = ab + 1
.MoveNext
Loop
.Close
End With
Label5.Caption = "Record Schedule For:"
Label33.Caption = frmsubsched.Text2.Text
picstep1.Visible = True
picstep2.Visible = False
picstep3.Visible = False
picstep4.Visible = False
picstep5.Visible = False
End Sub

Private Sub Image12_Click()
frmvalteachersched.Show 1
End Sub

Private Sub Image2_Click()
If Option1.Value = True Then
session = 1
Else
session = 2
End If
frmmappop.Show 1
End Sub



Private Sub list_Click()
On Error Resume Next
adviser = List.SelectedItem.Text
End Sub

Private Sub lvButtons_H1_Click()
Call cancelselect
End Sub

Private Sub lvButtons_H2_Click()
Call cancelselect
End Sub

Private Sub lvButtons_H3_Click()
Call cancelselect
End Sub

Private Sub lvButtons_H4_Click()
Call cancelselect
End Sub
Sub cancelselect()
If MsgBox("Are you sure you want to cancel Setup Schedule", vbYesNo + vbCritical) = vbYes Then
MsgBox "Setup Cancelled"
Unload Me
End If
End Sub

Private Sub lvButtons_H7_Click()
MsgBox "Setup Finished. Schedule for " & frmsubsched.Text2.Text & "is set."
End Sub

Private Sub lvButtons_H5_Click()
picstep4.Visible = False
picstep3.Visible = True
End Sub

Private Sub lvButtons_H6_Click()
If lblvalsubject(1).Caption = "Not Set" Or lblvalsubject(2).Caption = "Not Set" Then
MsgBox "complete schedule first before you proceed to the next step"
Else
picstep5.Visible = True
picstep4.Visible = False
cmd4.SetFocus
For a = 0 To 7
    If cmdsubjsched(a).Caption <> "B" Then
    List.ListItems.Add , , Me.lblteacher(a).Caption
    List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , Me.lblsubject(a).Caption
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
    RS.Open "select * from tblsubjsched Where adviser='" & lblteacher(a).Caption & "'"
If RS.EOF = False Then
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "unavailable"
Else
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "available"
End If
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , teacherid(a)
RS.Close
End If
Next a
End If
End Sub

Private Sub cmd2_Click()
If Len(Text1.Text) = 0 Then
MsgBox "Select Room First"
Else
picstep1.Visible = False
picstep2.Visible = False
picstep4.Visible = False
picstep3.Visible = True
If Option1.Value = True Then
For a = 0 To 7
lblsubject(a).Caption = "Not Set"
lblteacher(a).Caption = "Not Set"
Next a
lbltime(0).Caption = "6:00-6:50"
lblclass(0).Caption = "1st"
cmdsubjsched(0).Caption = 1
lbltime(1).Caption = "6:50-7:40"
lblclass(1).Caption = "2nd"
cmdsubjsched(1).Caption = 2
lbltime(2).Caption = "7:40-8:30"
lblclass(2).Caption = "3rd"
cmdsubjsched(2).Caption = 3
lbltime(5).Caption = "9:40-10:30"
lblclass(5).Caption = "5th"
cmdsubjsched(5).Caption = 5
lbltime(6).Caption = "10:30-11:20"
lblclass(6).Caption = "6th"
cmdsubjsched(6).Caption = 6
lbltime(7).Caption = "11:20-12:10"
lblclass(7).Caption = "7th"
cmdsubjsched(7).Caption = 7
If Mid(frmsubsched.Text2.Text, 1, 1) = "1" Or Mid(frmsubsched.Text2.Text, 1, 1) = "2" Then
lbltime(3).Caption = "8:30-8:50"
lblclass(3).Caption = "Break Time"
cmdsubjsched(3).Caption = "B"
lblteacher(3).Caption = "-------"
lblsubject(3).Caption = "-------"
lbltime(4).Caption = "8:50-9:40"
lblclass(4).Caption = "4th"
cmdsubjsched(4).Caption = "4"
lblteacher(4).Caption = "Not Set"
lblsubject(4).Caption = "Not Set"
Else
lbltime(3).Caption = "8:30-9:20"
lblclass(3).Caption = "4th"
lblteacher(3).Caption = "Not Set"
lblsubject(3).Caption = "Not Set"
cmdsubjsched(4).Caption = "4"
lbltime(4).Caption = "9:20-9:40"
lblclass(4).Caption = "Break Time"
cmdsubjsched(4).Caption = "B"
lblteacher(4).Caption = "-------"
lblsubject(4).Caption = "-------"
End If
Else
lbltime(0).Caption = "12:30-1:20"
lblclass(0).Caption = "1st"
cmdsubjsched(0).Caption = 1
lbltime(1).Caption = "1:20-2:10"
lblclass(1).Caption = "2nd"
cmdsubjsched(1).Caption = 2
lbltime(2).Caption = "2:10-3:00"
lblclass(2).Caption = "3rd"
cmdsubjsched(2).Caption = 3
lbltime(5).Caption = "4:10-5:00"
lblclass(5).Caption = "5th"
cmdsubjsched(5).Caption = 5
lbltime(6).Caption = "5:00-5:50"
lblclass(6).Caption = "6th"
cmdsubjsched(6).Caption = 6
lbltime(7).Caption = "5:50-6:40"
lblclass(7).Caption = "7th"
cmdsubjsched(7).Caption = 7
If Mid(frmsubsched.Text2.Text, 1, 1) = "1" Or Mid(frmsubsched.Text2.Text, 1, 1) = "2" Then
lbltime(3).Caption = "3:00-3:20"
lblclass(3).Caption = "Break Time"
cmdsubjsched(3).Caption = "B"
lblteacher(3).Caption = "-------"
lblsubject(3).Caption = "-------"
lbltime(4).Caption = "3:20-4:10"
lblclass(4).Caption = "4th"
cmdsubjsched(4).Caption = "4"
lblteacher(4).Caption = "Not Set"
lblsubject(4).Caption = "Not Set"
Else
lbltime(3).Caption = "3:00-3:50"
lblclass(3).Caption = "4th"
lblteacher(3).Caption = "Not Set"
lblsubject(3).Caption = "Not Set"
cmdsubjsched(4).Caption = "4"
lbltime(4).Caption = "3:50-4:10"
lblclass(4).Caption = "Break Time"
cmdsubjsched(4).Caption = "B"
lblteacher(4).Caption = "-------"
lblsubject(4).Caption = "-------"
End If
End If
End If
End Sub

Private Sub Picture1_Click()
frmvalteachersched.Show 1
End Sub

