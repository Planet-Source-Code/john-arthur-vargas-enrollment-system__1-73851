VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmroomsched 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   10125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14040
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   14040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   1
      Left            =   8745
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   52
      ImageHeight     =   53
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":2D16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":59BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":8424
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":B380
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":E4E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":101BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":11CC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":13980
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":1565A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":173F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":19FFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":1BDFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":1E89F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":205CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":24852
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":284E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":2A3DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":2C0A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":2E998
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":31632
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   9360
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   52
      ImageHeight     =   53
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":332FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":36255
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":38F41
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":3BA95
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":3EB17
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":41C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":43A96
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":455E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":4728B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":49028
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":4ADC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":4DB3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":4F9B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":52550
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":54377
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":586A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":5C399
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":5E2E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":60119
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":62A51
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":64839
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   2
      Left            =   9960
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   91
      ImageHeight     =   86
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":66599
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":6A851
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":6EAEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubsched.frx":72A91
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7575
      Left            =   10740
      TabIndex        =   76
      Top             =   480
      Width           =   3135
      Begin VB.Frame Frame7 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   255
         Left            =   -240
         TabIndex        =   77
         Top             =   9120
         Width           =   3615
      End
      Begin VB.Line Line62 
         X1              =   120
         X2              =   1440
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Line Line61 
         X1              =   120
         X2              =   1440
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Line60 
         X1              =   120
         X2              =   1440
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line59 
         X1              =   120
         X2              =   1440
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line58 
         X1              =   120
         X2              =   1440
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line57 
         X1              =   120
         X2              =   1440
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line56 
         X1              =   120
         X2              =   1440
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line55 
         X1              =   1440
         X2              =   3120
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Line Line54 
         X1              =   1440
         X2              =   3120
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Line53 
         X1              =   1440
         X2              =   3120
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line52 
         X1              =   1440
         X2              =   3120
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line51 
         X1              =   1440
         X2              =   3120
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line50 
         X1              =   1440
         X2              =   3120
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line49 
         X1              =   1440
         X2              =   3120
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line48 
         X1              =   1440
         X2              =   3120
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line47 
         X1              =   1440
         X2              =   3120
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line46 
         X1              =   1440
         X2              =   3120
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line45 
         X1              =   1440
         X2              =   3120
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line44 
         X1              =   1320
         X2              =   3120
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line43 
         X1              =   1440
         X2              =   3120
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line42 
         X1              =   1440
         X2              =   3120
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line41 
         X1              =   120
         X2              =   1440
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line40 
         X1              =   120
         X2              =   1440
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line39 
         X1              =   120
         X2              =   1440
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line38 
         X1              =   120
         X2              =   1440
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line37 
         X1              =   120
         X2              =   1320
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line36 
         X1              =   120
         X2              =   1320
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line35 
         X1              =   120
         X2              =   1440
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Afternoon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   0
         TabIndex        =   116
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label lblaftsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   115
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label lblaftsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   7
         Left            =   1440
         TabIndex        =   114
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label lblaftsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   113
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label lblaftsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   112
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label lblaftsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   111
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label lblaftsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   110
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label lblaftsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   109
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label lblaftsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   108
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label lblafttime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "5:00-5:50"
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
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   107
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label lblafttime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "5:50-6:40"
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
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   106
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Label lblafttime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "4:10-5:00"
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
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   105
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label lblafttime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "3:20-4:10"
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
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   104
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label lblafttime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "3:00-3:20"
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
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   103
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label lblafttime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "2:10-3:00"
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
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   102
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label lblafttime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "12:30-1:20"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   101
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label lblafttime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "1:20-2:10"
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   100
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Section to Teach"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   1320
         TabIndex        =   99
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   120
         TabIndex        =   98
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label lblmornsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   97
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblmornsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   7
         Left            =   1440
         TabIndex        =   96
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label lblmornsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   95
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label lblmornsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   94
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblmornsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   93
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblmornsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   92
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblmornsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   91
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblmornsec 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Class"
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
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   90
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblmorntime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "10:30-11:20"
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
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   89
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label lblmorntime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "11:20-12:10"
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
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   88
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblmorntime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "9:40-10:30"
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
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   87
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblmorntime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "8:50-9:40"
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
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   86
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblmorntime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "8:30-8:50"
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
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   85
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblmorntime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "7:40-8:30"
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
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   84
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblmorntime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "6:00-6:50"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   83
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblmorntime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "6:50-7:40"
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   82
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Section to Teach"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   1320
         TabIndex        =   81
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   120
         TabIndex        =   80
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Morning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   120
         TabIndex        =   79
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Left            =   0
         TabIndex        =   78
         Top             =   0
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00400000&
         Height          =   7530
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   3135
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   0
         Picture         =   "frmsubsched.frx":769CA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7335
      End
   End
   Begin VB.OptionButton optsday 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Friday"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   7920
      TabIndex        =   46
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton optsday 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Thursday"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6720
      TabIndex        =   45
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton optsday 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Wednesday"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   44
      Top             =   480
      Width           =   1335
   End
   Begin VB.OptionButton optsday 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Monday"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   43
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton optsday 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tuesday"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   42
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9420
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3135
      Begin VB.TextBox Text6 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   127
         Top             =   4120
         Width           =   1815
      End
      Begin VB.TextBox Text5 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   125
         Top             =   3680
         Width           =   1815
      End
      Begin VB.TextBox Text4 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   123
         Top             =   5400
         Width           =   1815
      End
      Begin VB.TextBox Text3 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   121
         Top             =   4980
         Width           =   1815
      End
      Begin VB.TextBox Text2 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   119
         Top             =   4560
         Width           =   1815
      End
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   117
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   255
         Left            =   -240
         TabIndex        =   75
         Top             =   9240
         Width           =   3615
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2760
         Width           =   1815
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf1 
         Height          =   2655
         Left            =   0
         TabIndex        =   5
         Top             =   6600
         Width           =   3135
         _cx             =   5530
         _cy             =   4683
         FlashVars       =   ""
         Movie           =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\teachersched.swf"
         Src             =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\teachersched.swf"
         WMode           =   "Window"
         Play            =   0   'False
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         AllowScriptAccess=   ""
         Scale           =   "NoBorder"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   -1  'True
         Profile         =   0   'False
         ProfileAddress  =   ""
         ProfilePort     =   0
      End
      Begin TNHSES.lvButtons_H Command4 
         Height          =   615
         Left            =   120
         TabIndex        =   129
         Top             =   5880
         Width           =   1455
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   "Print Schedule"
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
         Left            =   1680
         TabIndex        =   130
         Top             =   5880
         Width           =   1335
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
      Begin VB.Image Image7 
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   600
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   3360
         Y1              =   6580
         Y2              =   6580
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No."
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
         Left            =   120
         TabIndex        =   128
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   120
         TabIndex        =   126
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Room"
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
         Left            =   120
         TabIndex        =   124
         Top             =   5400
         Width           =   855
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Advisory Class"
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
         Left            =   120
         TabIndex        =   122
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
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
         Left            =   120
         TabIndex        =   120
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   120
         TabIndex        =   118
         Top             =   3240
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00400000&
         Height          =   6570
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   " Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Left            =   120
         TabIndex        =   3
         Top             =   2880
         Width           =   495
      End
      Begin VB.Image Image6 
         Height          =   495
         Left            =   0
         Picture         =   "frmsubsched.frx":76EFA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7335
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   3360
      TabIndex        =   6
      Top             =   840
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "1st Floor"
      TabPicture(0)   =   "frmsubsched.frx":7742A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "2nd Floor"
      TabPicture(1)   =   "frmsubsched.frx":77446
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optday(4)"
      Tab(1).Control(1)=   "optday(3)"
      Tab(1).Control(2)=   "optday(2)"
      Tab(1).Control(3)=   "optday(0)"
      Tab(1).Control(4)=   "optday(1)"
      Tab(1).Control(5)=   "Frame2(1)"
      Tab(1).Control(6)=   "Shape2"
      Tab(1).Control(7)=   "Label37"
      Tab(1).Control(8)=   "Label13(18)"
      Tab(1).Control(9)=   "Label13(0)"
      Tab(1).Control(10)=   "Label8"
      Tab(1).Control(11)=   "lblmorn(1)"
      Tab(1).Control(12)=   "lblmorn(2)"
      Tab(1).Control(13)=   "lblmorn(3)"
      Tab(1).Control(14)=   "lblmorn(6)"
      Tab(1).Control(15)=   "lblmorn(7)"
      Tab(1).Control(16)=   "lblmorn(8)"
      Tab(1).Control(17)=   "lblmorn(5)"
      Tab(1).Control(18)=   "lblmorn(4)"
      Tab(1).Control(19)=   "lblaft(4)"
      Tab(1).Control(20)=   "lblaft(5)"
      Tab(1).Control(21)=   "lblaft(8)"
      Tab(1).Control(22)=   "lblaft(0)"
      Tab(1).Control(23)=   "lblaft(7)"
      Tab(1).Control(24)=   "lblaft(6)"
      Tab(1).Control(25)=   "lblaft(3)"
      Tab(1).Control(26)=   "lblaft(2)"
      Tab(1).Control(27)=   "lblaft(1)"
      Tab(1).Control(28)=   "lblclass(7)"
      Tab(1).Control(29)=   "lblclass(0)"
      Tab(1).Control(30)=   "Label2"
      Tab(1).Control(31)=   "lblaft(1000)"
      Tab(1).Control(32)=   "Label9"
      Tab(1).Control(33)=   "lblaft(100)"
      Tab(1).Control(34)=   "Line18"
      Tab(1).Control(35)=   "Line17"
      Tab(1).Control(36)=   "Line16"
      Tab(1).Control(37)=   "Line15"
      Tab(1).Control(38)=   "Line13"
      Tab(1).Control(39)=   "Line12"
      Tab(1).Control(40)=   "Line11"
      Tab(1).Control(41)=   "Line10"
      Tab(1).Control(42)=   "Line9"
      Tab(1).Control(43)=   "Line3"
      Tab(1).Control(44)=   "Line4"
      Tab(1).Control(45)=   "Line5"
      Tab(1).Control(46)=   "Line7"
      Tab(1).Control(47)=   "Line8"
      Tab(1).Control(48)=   "Line6"
      Tab(1).Control(49)=   "Line14"
      Tab(1).ControlCount=   50
      Begin VB.OptionButton optday 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Index           =   4
         Left            =   -70200
         TabIndex        =   14
         Top             =   7320
         Width           =   855
      End
      Begin VB.OptionButton optday 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Index           =   3
         Left            =   -71400
         TabIndex        =   13
         Top             =   7320
         Width           =   1095
      End
      Begin VB.OptionButton optday 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Index           =   2
         Left            =   -72840
         TabIndex        =   12
         Top             =   7320
         Width           =   1335
      End
      Begin VB.OptionButton optday 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Index           =   0
         Left            =   -75000
         TabIndex        =   11
         Top             =   7320
         Width           =   975
      End
      Begin VB.OptionButton optday 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Index           =   1
         Left            =   -73920
         TabIndex        =   10
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6900
         Index           =   1
         Left            =   -75000
         TabIndex        =   9
         Top             =   360
         Width           =   7335
         Begin VB.Image Img3 
            Height          =   930
            Index           =   12
            Left            =   5790
            Picture         =   "frmsubsched.frx":77462
            ToolTipText     =   "TRS1-6"
            Top             =   4680
            Width           =   795
         End
         Begin VB.Image Img3 
            Height          =   990
            Index           =   13
            Left            =   5820
            Picture         =   "frmsubsched.frx":77B6C
            ToolTipText     =   "TRS1-5"
            Top             =   3720
            Width           =   750
         End
         Begin VB.Image Img3 
            Height          =   990
            Index           =   14
            Left            =   5790
            Picture         =   "frmsubsched.frx":781E8
            ToolTipText     =   "TRS1-4"
            Top             =   2760
            Width           =   780
         End
         Begin VB.Image Img3 
            Height          =   1290
            Index           =   15
            Left            =   2480
            Picture         =   "frmsubsched.frx":7880E
            ToolTipText     =   "TRS2-1"
            Top             =   3080
            Width           =   1365
         End
         Begin VB.Image Img3 
            Height          =   1305
            Index           =   16
            Left            =   3350
            Picture         =   "frmsubsched.frx":794B8
            ToolTipText     =   "TRS2-2"
            Top             =   3820
            Width           =   1305
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H8000000D&
            BorderWidth     =   3
            Height          =   6900
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   7260
         End
         Begin VB.Image Img3 
            Height          =   930
            Index           =   17
            Left            =   810
            Picture         =   "frmsubsched.frx":79F92
            ToolTipText     =   "SEDP-7"
            Top             =   375
            Width           =   750
         End
         Begin VB.Image Img3 
            Height          =   945
            Index           =   18
            Left            =   1530
            Picture         =   "frmsubsched.frx":7A6EB
            ToolTipText     =   "SEDP-8"
            Top             =   315
            Width           =   765
         End
         Begin VB.Image Img3 
            Height          =   900
            Index           =   19
            Left            =   2970
            Picture         =   "frmsubsched.frx":7ADFE
            ToolTipText     =   "SEDP-9"
            Top             =   370
            Width           =   615
         End
         Begin VB.Image Img3 
            Height          =   885
            Index           =   20
            Left            =   3570
            Picture         =   "frmsubsched.frx":7B39B
            ToolTipText     =   "SEDP-10"
            Top             =   360
            Width           =   975
         End
         Begin VB.Image Img3 
            Height          =   945
            Index           =   21
            Left            =   5720
            Picture         =   "frmsubsched.frx":7BA37
            ToolTipText     =   "SEDP-11"
            Top             =   360
            Width           =   840
         End
         Begin VB.Image Image8 
            Height          =   7560
            Left            =   -240
            Picture         =   "frmsubsched.frx":7C0DF
            Top             =   -600
            Width           =   7560
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6900
         Left            =   -75000
         TabIndex        =   8
         Top             =   320
         Width           =   7335
         Begin VB.Image Image5 
            Height          =   7560
            Left            =   -240
            Picture         =   "frmsubsched.frx":97B40
            Top             =   -600
            Width           =   7560
         End
         Begin VB.Image Image3 
            Height          =   930
            Index           =   12
            Left            =   5790
            Picture         =   "frmsubsched.frx":A4EE7
            Top             =   4680
            Width           =   795
         End
         Begin VB.Image Image3 
            Height          =   990
            Index           =   13
            Left            =   5820
            Picture         =   "frmsubsched.frx":A55F1
            Top             =   3720
            Width           =   750
         End
         Begin VB.Image Image3 
            Height          =   990
            Index           =   14
            Left            =   5790
            Picture         =   "frmsubsched.frx":A5C6D
            Top             =   2760
            Width           =   780
         End
         Begin VB.Image Image3 
            Height          =   1290
            Index           =   15
            Left            =   2480
            Picture         =   "frmsubsched.frx":A6293
            Top             =   3080
            Width           =   1365
         End
         Begin VB.Image Image3 
            Height          =   1305
            Index           =   16
            Left            =   3350
            Picture         =   "frmsubsched.frx":A6F3D
            Top             =   3820
            Width           =   1305
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H8000000D&
            BorderWidth     =   3
            Height          =   6900
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   7260
         End
         Begin VB.Image Image3 
            Height          =   930
            Index           =   17
            Left            =   840
            Picture         =   "frmsubsched.frx":A7A17
            Top             =   375
            Width           =   750
         End
         Begin VB.Image Image3 
            Height          =   945
            Index           =   18
            Left            =   1530
            Picture         =   "frmsubsched.frx":A8170
            Top             =   320
            Width           =   765
         End
         Begin VB.Image Image3 
            Height          =   900
            Index           =   19
            Left            =   2970
            Picture         =   "frmsubsched.frx":A8883
            Top             =   370
            Width           =   615
         End
         Begin VB.Image Image3 
            Height          =   885
            Index           =   20
            Left            =   3570
            Picture         =   "frmsubsched.frx":A8E20
            Top             =   360
            Width           =   975
         End
         Begin VB.Image Image3 
            Height          =   945
            Index           =   21
            Left            =   5720
            Picture         =   "frmsubsched.frx":A94BC
            Top             =   360
            Width           =   840
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   7900
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   320
         Width           =   7455
         Begin VB.Image Img3 
            Height          =   795
            Index           =   1
            Left            =   5835
            ToolTipText     =   "TRS1-E2"
            Top             =   5550
            Width           =   780
         End
         Begin VB.Image Img3 
            Height          =   915
            Index           =   2
            Left            =   5805
            Picture         =   "frmsubsched.frx":A9B64
            ToolTipText     =   "TRS1-3"
            Top             =   4680
            Width           =   780
         End
         Begin VB.Image Img3 
            Height          =   990
            Index           =   3
            Left            =   5835
            Picture         =   "frmsubsched.frx":AA1F5
            ToolTipText     =   "TRS1-2"
            Top             =   3720
            Width           =   735
         End
         Begin VB.Image Img3 
            Height          =   1035
            Index           =   4
            Left            =   5790
            Picture         =   "frmsubsched.frx":AA828
            ToolTipText     =   "TRS1-1"
            Top             =   2760
            Width           =   825
         End
         Begin VB.Image Img3 
            Height          =   915
            Index           =   5
            Left            =   780
            Picture         =   "frmsubsched.frx":AAF3A
            ToolTipText     =   "SEDP-1"
            Top             =   405
            Width           =   795
         End
         Begin VB.Image Img3 
            Height          =   915
            Index           =   7
            Left            =   2970
            ToolTipText     =   "SEDP-3"
            Top             =   360
            Width           =   660
         End
         Begin VB.Image Img3 
            Height          =   900
            Index           =   8
            Left            =   3600
            Picture         =   "frmsubsched.frx":AB753
            ToolTipText     =   "SEDP-4"
            Top             =   360
            Width           =   975
         End
         Begin VB.Image Img3 
            Height          =   900
            Index           =   9
            Left            =   4515
            Picture         =   "frmsubsched.frx":ABE43
            ToolTipText     =   "SEDP-5"
            Top             =   360
            Width           =   930
         End
         Begin VB.Image Img3 
            Height          =   945
            Index           =   10
            Left            =   5730
            Picture         =   "frmsubsched.frx":AC536
            ToolTipText     =   "SEDP-6"
            Top             =   360
            Width           =   855
         End
         Begin VB.Image Img3 
            Height          =   570
            Index           =   11
            Left            =   5520
            Picture         =   "frmsubsched.frx":ACBFA
            ToolTipText     =   "TRS1-E1"
            Top             =   1800
            Width           =   1065
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H8000000D&
            BorderWidth     =   3
            Height          =   6900
            Left            =   0
            Top             =   0
            Width           =   7260
         End
         Begin VB.Image Img3 
            Height          =   915
            Index           =   6
            Left            =   1545
            Picture         =   "frmsubsched.frx":AD1D2
            ToolTipText     =   "SEDP-2"
            Top             =   375
            Width           =   780
         End
         Begin VB.Image Image4 
            Height          =   7560
            Left            =   -240
            Picture         =   "frmsubsched.frx":AED94
            Top             =   -600
            Width           =   7560
         End
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C000&
         BorderWidth     =   4
         Height          =   2100
         Left            =   -75000
         Top             =   7680
         Width           =   8400
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
         Left            =   -74880
         TabIndex        =   41
         Top             =   8280
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
         Left            =   -74880
         TabIndex        =   40
         Top             =   8040
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
         Left            =   -74880
         TabIndex        =   39
         Top             =   8760
         Width           =   1260
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
         Left            =   -74880
         TabIndex        =   38
         Top             =   9000
         Width           =   1260
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
         Left            =   -73440
         TabIndex        =   37
         Top             =   8280
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
         Index           =   2
         Left            =   -72555
         TabIndex        =   36
         Top             =   8280
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
         Index           =   3
         Left            =   -71670
         TabIndex        =   35
         Top             =   8280
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
         Left            =   -68640
         TabIndex        =   34
         Top             =   8280
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
         Index           =   7
         Left            =   -67755
         TabIndex        =   33
         Top             =   8280
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
         Index           =   8
         Left            =   -69885
         TabIndex        =   32
         Top             =   8280
         Width           =   360
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
         Left            =   -69540
         TabIndex        =   31
         Top             =   8280
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
         Index           =   4
         Left            =   -70425
         TabIndex        =   30
         Top             =   8280
         Width           =   540
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
         Left            =   -70425
         TabIndex        =   29
         Top             =   9000
         Width           =   540
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
         Left            =   -69540
         TabIndex        =   28
         Top             =   9000
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
         Index           =   8
         Left            =   -69885
         TabIndex        =   27
         Top             =   9000
         Width           =   360
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
         Left            =   -70770
         TabIndex        =   26
         Top             =   9000
         Width           =   360
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
         Left            =   -67755
         TabIndex        =   25
         Top             =   9000
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
         Left            =   -68640
         TabIndex        =   24
         Top             =   9000
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
         Index           =   3
         Left            =   -71670
         TabIndex        =   23
         Top             =   9000
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
         Index           =   2
         Left            =   -72555
         TabIndex        =   22
         Top             =   9000
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
         Index           =   1
         Left            =   -73440
         TabIndex        =   21
         Top             =   9000
         Width           =   900
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
         Left            =   -73560
         TabIndex        =   20
         Top             =   8040
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
         Left            =   -73560
         TabIndex        =   19
         Top             =   8760
         Width           =   7005
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -67560
         TabIndex        =   18
         Top             =   7320
         Width           =   1815
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
         Left            =   -68040
         TabIndex        =   17
         Top             =   7320
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68760
         TabIndex        =   16
         Top             =   7320
         Width           =   615
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
         Left            =   -69240
         TabIndex        =   15
         Top             =   7320
         Width           =   405
      End
      Begin VB.Line Line18 
         BorderWidth     =   2
         X1              =   -67755
         X2              =   -67755
         Y1              =   9000
         Y2              =   9360
      End
      Begin VB.Line Line17 
         BorderWidth     =   2
         X1              =   -68640
         X2              =   -68640
         Y1              =   9000
         Y2              =   9360
      End
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   -69885
         X2              =   -69885
         Y1              =   9000
         Y2              =   9360
      End
      Begin VB.Line Line15 
         BorderWidth     =   2
         X1              =   -70425
         X2              =   -70425
         Y1              =   9000
         Y2              =   9360
      End
      Begin VB.Line Line13 
         BorderWidth     =   2
         X1              =   -70770
         X2              =   -70770
         Y1              =   9000
         Y2              =   9360
      End
      Begin VB.Line Line12 
         BorderWidth     =   2
         X1              =   -71670
         X2              =   -71670
         Y1              =   9000
         Y2              =   9360
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   -72555
         X2              =   -72555
         Y1              =   9000
         Y2              =   9360
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   -72555
         X2              =   -72555
         Y1              =   8280
         Y2              =   8640
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   -71670
         X2              =   -71670
         Y1              =   8280
         Y2              =   8640
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   -70770
         X2              =   -70770
         Y1              =   8280
         Y2              =   8640
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   -70425
         X2              =   -70425
         Y1              =   8280
         Y2              =   8640
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   -69885
         X2              =   -69885
         Y1              =   8280
         Y2              =   8640
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   -68640
         X2              =   -68640
         Y1              =   8280
         Y2              =   8640
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   -67755
         X2              =   -67755
         Y1              =   8280
         Y2              =   8640
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   -69540
         X2              =   -69540
         Y1              =   8280
         Y2              =   8640
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   -69540
         X2              =   -69540
         Y1              =   9000
         Y2              =   9360
      End
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
      Index           =   3
      Left            =   11160
      TabIndex        =   131
      Top             =   9600
      Width           =   2805
   End
   Begin VB.Line Line34 
      BorderWidth     =   2
      X1              =   10605
      X2              =   10605
      Y1              =   9240
      Y2              =   9600
   End
   Begin VB.Line Line33 
      BorderWidth     =   2
      X1              =   9720
      X2              =   9720
      Y1              =   9240
      Y2              =   9600
   End
   Begin VB.Line Line32 
      BorderWidth     =   2
      X1              =   8475
      X2              =   8475
      Y1              =   9240
      Y2              =   9600
   End
   Begin VB.Line Line31 
      BorderWidth     =   2
      X1              =   7935
      X2              =   7935
      Y1              =   9240
      Y2              =   9600
   End
   Begin VB.Line Line30 
      BorderWidth     =   2
      X1              =   7590
      X2              =   7590
      Y1              =   9240
      Y2              =   9600
   End
   Begin VB.Line Line29 
      BorderWidth     =   2
      X1              =   6690
      X2              =   6690
      Y1              =   9240
      Y2              =   9600
   End
   Begin VB.Line Line28 
      BorderWidth     =   2
      X1              =   5805
      X2              =   5805
      Y1              =   9240
      Y2              =   9600
   End
   Begin VB.Line Line27 
      BorderWidth     =   2
      X1              =   5805
      X2              =   5805
      Y1              =   8520
      Y2              =   8880
   End
   Begin VB.Line Line26 
      BorderWidth     =   2
      X1              =   6690
      X2              =   6690
      Y1              =   8520
      Y2              =   8880
   End
   Begin VB.Line Line25 
      BorderWidth     =   2
      X1              =   7590
      X2              =   7590
      Y1              =   8520
      Y2              =   8880
   End
   Begin VB.Line Line24 
      BorderWidth     =   2
      X1              =   7935
      X2              =   7935
      Y1              =   8520
      Y2              =   8880
   End
   Begin VB.Line Line23 
      BorderWidth     =   2
      X1              =   8475
      X2              =   8475
      Y1              =   8520
      Y2              =   8880
   End
   Begin VB.Line Line22 
      BorderWidth     =   2
      X1              =   9720
      X2              =   9720
      Y1              =   8520
      Y2              =   8880
   End
   Begin VB.Line Line21 
      BorderWidth     =   2
      X1              =   10605
      X2              =   10605
      Y1              =   8520
      Y2              =   8880
   End
   Begin VB.Line Line20 
      BorderWidth     =   2
      X1              =   8820
      X2              =   8820
      Y1              =   8520
      Y2              =   8880
   End
   Begin VB.Line Line19 
      BorderWidth     =   2
      X1              =   8820
      X2              =   8820
      Y1              =   9240
      Y2              =   9600
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderWidth     =   2
      Height          =   10110
      Index           =   3
      Left            =   10
      Top             =   10
      Width           =   14025
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
      Index           =   10
      Left            =   11760
      TabIndex        =   74
      Top             =   9000
      Width           =   405
   End
   Begin VB.Label Label3 
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
      Left            =   12240
      TabIndex        =   73
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      Left            =   3480
      TabIndex        =   72
      Top             =   8520
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
      Index           =   2
      Left            =   3480
      TabIndex        =   71
      Top             =   8280
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
      Index           =   1
      Left            =   3480
      TabIndex        =   70
      Top             =   9000
      Width           =   1260
   End
   Begin VB.Label Label4 
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
      Left            =   3480
      TabIndex        =   69
      Top             =   9240
      Width           =   1260
   End
   Begin VB.Label lblmorns 
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
      Left            =   4920
      TabIndex        =   68
      ToolTipText     =   "1st"
      Top             =   8520
      Width           =   865
   End
   Begin VB.Label lblmorns 
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
      Left            =   5805
      TabIndex        =   67
      ToolTipText     =   "2nd"
      Top             =   8520
      Width           =   865
   End
   Begin VB.Label lblmorns 
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
      Left            =   6690
      TabIndex        =   66
      ToolTipText     =   "3rd"
      Top             =   8520
      Width           =   880
   End
   Begin VB.Label lblmorns 
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
      Left            =   9720
      TabIndex        =   65
      ToolTipText     =   "6th"
      Top             =   8520
      Width           =   870
   End
   Begin VB.Label lblmorns 
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
      Left            =   10605
      TabIndex        =   64
      ToolTipText     =   "7th"
      Top             =   8520
      Width           =   900
   End
   Begin VB.Label lblmorns 
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
      Left            =   7590
      TabIndex        =   63
      ToolTipText     =   "4th"
      Top             =   8520
      Width           =   330
   End
   Begin VB.Label lblmorns 
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
      Left            =   8475
      TabIndex        =   62
      ToolTipText     =   "4th"
      Top             =   8520
      Width           =   330
   End
   Begin VB.Label lblmorns 
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
      Left            =   8820
      TabIndex        =   61
      ToolTipText     =   "5th"
      Top             =   8520
      Width           =   885
   End
   Begin VB.Label lblmorns 
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
      Left            =   7935
      TabIndex        =   60
      ToolTipText     =   "4th"
      Top             =   8520
      Width           =   525
   End
   Begin VB.Label lblafts 
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
      Left            =   7935
      TabIndex        =   59
      ToolTipText     =   "4th"
      Top             =   9240
      Width           =   525
   End
   Begin VB.Label lblafts 
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
      Left            =   8820
      TabIndex        =   58
      ToolTipText     =   "5th"
      Top             =   9240
      Width           =   885
   End
   Begin VB.Label lblafts 
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
      Left            =   8475
      TabIndex        =   57
      ToolTipText     =   "4th"
      Top             =   9240
      Width           =   330
   End
   Begin VB.Label lblafts 
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
      Left            =   7590
      TabIndex        =   56
      ToolTipText     =   "4th"
      Top             =   9240
      Width           =   330
   End
   Begin VB.Label lblafts 
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
      Left            =   10605
      TabIndex        =   55
      ToolTipText     =   "7th"
      Top             =   9240
      Width           =   900
   End
   Begin VB.Label lblafts 
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
      Left            =   9720
      TabIndex        =   54
      ToolTipText     =   "6th"
      Top             =   9240
      Width           =   870
   End
   Begin VB.Label lblafts 
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
      Left            =   6690
      TabIndex        =   53
      ToolTipText     =   "3rd"
      Top             =   9240
      Width           =   880
   End
   Begin VB.Label lblafts 
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
      Left            =   5805
      TabIndex        =   52
      ToolTipText     =   "2nd"
      Top             =   9240
      Width           =   865
   End
   Begin VB.Label lblafts 
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
      Left            =   4920
      TabIndex        =   51
      ToolTipText     =   "1st"
      Top             =   9240
      Width           =   865
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
      Index           =   2
      Left            =   4800
      TabIndex        =   50
      Top             =   8280
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
      Index           =   1
      Left            =   4800
      TabIndex        =   49
      Top             =   9000
      Width           =   7005
   End
   Begin VB.Label Label1 
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
      Left            =   12240
      TabIndex        =   48
      Top             =   8640
      Width           =   615
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
      Index           =   9
      Left            =   11760
      TabIndex        =   47
      Top             =   8640
      Width           =   405
   End
   Begin VB.Image Image3 
      Height          =   225
      Index           =   0
      Left            =   13560
      Picture         =   "frmsubsched.frx":CDB92
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "TNHS Enrollment System- Teacher's Schedule"
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
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmsubsched.frx":D0B3F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14115
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      BorderWidth     =   4
      Height          =   1740
      Left            =   3360
      Top             =   8160
      Width           =   10560
   End
End
Attribute VB_Name = "frmroomsched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If usrsubject = "Values Educ." Then
Call dbConnection
Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = 3
    RS.ActiveConnection = Con
    RS.LockType = 2
        RS.Open "select * from tblvalsched Where teacherid='" & teacherids & "' Order by intday"
Set DataReport5.DataSource = RS.DataSource
For Each obj In DataReport5.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = RS.DataMember
    End If
Next
DataReport5.Sections("Section2").Controls("label19").Caption = txt1.Text
DataReport5.Sections("Section2").Controls("label24").Caption = Text5.Text
DataReport5.Sections("Section2").Controls("label21").Caption = Text6.Text
DataReport5.Sections("Section2").Controls("label13").Caption = Text3.Text
DataReport5.Sections("Section2").Controls("label26").Caption = Text4.Text
DataReport5.Sections("Section1").Controls("Text1").DataField = "class"
DataReport5.Sections("Section1").Controls("Text2").DataField = "session"
DataReport5.Sections("Section1").Controls("Text3").DataField = "yr-section"
DataReport5.Sections("Section1").Controls("Text4").DataField = "room"
DataReport5.Sections("Section1").Controls("Text5").DataField = "time"
DataReport5.Sections("Section1").Controls("Text6").DataField = "day"
DataReport5.Refresh
DataReport5.Show 1
Else
Call dbConnection
Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = 3
    RS.ActiveConnection = Con
    RS.LockType = 2
        RS.Open "select * from tblsched Where teacherid='" & teacherids & "' Order by class"
Set DataReport4.DataSource = RS.DataSource
For Each obj In DataReport4.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = RS.DataMember
    End If
Next
DataReport4.Sections("Section2").Controls("label19").Caption = txt1.Text
DataReport4.Sections("Section2").Controls("label24").Caption = Text5.Text
DataReport4.Sections("Section2").Controls("label21").Caption = Text6.Text
DataReport4.Sections("Section2").Controls("label13").Caption = Text3.Text
DataReport4.Sections("Section2").Controls("label26").Caption = Text4.Text
DataReport4.Sections("Section1").Controls("Text1").DataField = "class"
DataReport4.Sections("Section1").Controls("Text2").DataField = "session"
DataReport4.Sections("Section1").Controls("Text3").DataField = "yr-section"
DataReport4.Sections("Section1").Controls("Text4").DataField = "room"
DataReport4.Sections("Section1").Controls("Text5").DataField = "time"
DataReport4.Refresh
DataReport4.Show 1
End If
End Sub

Private Sub Form_Load()

swf1.Movie = App.Path & "\flash\teachersched.swf"
swf1.Play
Me.Width = 40
If userlevel = "Admin" Then
Else
'Picture1.Enabled = False
txt1.Text = userid
Text1.Text = names
Text2.Text = usrsubject
Text5.Text = address
Text6.Text = number
teacher = names
teacherids = userid
Image7.Picture = LoadPicture(App.Path & "/usrpic/" & txt1.Text & ".jpg")
If usrsubject = "Social Studies" Or usrsubject = "MAPEH" Or usrsubject = "TLE" Or usrsubject = "Values Educ." Then
optsday(1).Enabled = True
optsday(2).Enabled = True
optsday(3).Enabled = True
optsday(4).Enabled = True
optsday(5).Enabled = True
optsday(1).Value = True
days = "Monday"
End If
Call loadall
End If



End Sub

Private Sub Image3_Click(Index As Integer)
Unload Me
End Sub



Private Sub img3_Click(Index As Integer)
boolpop = 1
mapid = Index
pid = Img3(Index).ToolTipText
frmviewschedpop.Show 1

End Sub

Private Sub lblafts_Click(Index As Integer)
boolpop = 2
mapid = Index
pid = lblmorns(Index).ToolTipText
sessions = "Afternoon"
frmviewschedpop.Show 1
End Sub

Private Sub lblmorns_Click(Index As Integer)
boolpop = 2
mapid = Index
pid = lblmorns(Index).ToolTipText
sessions = "Morning"
frmviewschedpop.Show 1
End Sub

Private Sub optsday_Click(Index As Integer)
days = optsday(Index).Caption
Call loadall
End Sub

Private Sub Timer1_Timer()
If Me.Width = 14040 Then
Timer1.Enabled = False
Else
Me.Left = Me.Left - 250
Me.Width = Me.Width + 500
End If
End Sub
Sub loadall()
For a = 0 To 7
lblmornsec(a).Caption = "No Class"
lblaftsec(a).Caption = "No Class"
Next a
For a = 1 To 21
Img3(a).Picture = ImageList1(1).ListImages(a).Picture
Next a
If usrsubject <> "Values Educ." Then
Call dbConnection
  Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "Select * From tblsubjsched Where adviserid='" & teacherids & "'"
If RS.EOF = False Then
Text3.Text = RS.Fields(0)
Text4.Text = RS.Fields(3)
Else
Text3.Text = "No Advisory Class"
End If
For a = 0 To 8
lblmorns(a).BackColor = &HFFFFC0
lblafts(a).BackColor = &HFFFFC0
Next a
Call dbConnection
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblsched Where teacherid='" & teacherids & "'"
If RS1.EOF = False Then
Do Until RS1.EOF
If RS1.Fields(1) = "Morning" Then
Select Case RS1.Fields(5)
Case "6:00-6:50": lblmornsec(0).Caption = RS1.Fields(0)
Case "6:50-7:40": lblmornsec(1).Caption = RS1.Fields(0)
Case "7:40-8:30": lblmornsec(2).Caption = RS1.Fields(0)
Case "9:40-10:30": lblmornsec(5).Caption = RS1.Fields(0)
Case "10:30-11:20": lblmornsec(6).Caption = RS1.Fields(0)
Case "11:20-12:10": lblmornsec(7).Caption = RS1.Fields(0)
Case "8:50-9:40":
lblmorntime(3).Caption = "8:30-8:50"
lblmornsec(3).Caption = "Break Time"
lblmorntime(4).Caption = "8:50-9:40"
lblmornsec(4).Caption = RS1.Fields(0)
Case "8:30-9:20":
lblmorntime(4).Caption = "9:20-9:40"
lblmornsec(4).Caption = "Break Time"
lblmorntime(3).Caption = "8:30-9:20"
lblmornsec(3).Caption = RS1.Fields(0)
End Select
Select Case RS1.Fields(4)
Case "1st": lblmorns(1).BackColor = &H8080FF
Case "2nd": lblmorns(2).BackColor = &H8080FF
Case "3rd": lblmorns(3).BackColor = &H8080FF
Case "4th":
lblmorns(4).BackColor = &H8080FF
If Right(RS1.Fields(5), 2) = "20" Then
lblmorns(0).BackColor = &H8080FF
Else
lblmorns(8).BackColor = &H8080FF
End If
Case "5th": lblmorns(5).BackColor = &H8080FF
Case "6th": lblmorns(6).BackColor = &H8080FF
Case "7th": lblmorns(7).BackColor = &H8080FF
End Select
Else
Select Case RS1.Fields(5)
Case "12:30-1:20": lblaftsec(0).Caption = RS1.Fields(0)
Case "1:20-2:10": lblaftsec(1).Caption = RS1.Fields(0)
Case "2:10-3:00": lblaftsec(2).Caption = RS1.Fields(0)
Case "4:10-5:00": lblaftsec(5).Caption = RS1.Fields(0)
Case "5:00-5:50": lblaftsec(6).Caption = RS1.Fields(0)
Case "5:50-6:40": lblaftsec(7).Caption = RS1.Fields(0)
Case "3:20-4:10":
lblafttime(3).Caption = "3:00-3:20"
lblaftsec(3).Caption = "Break Time"
lblafttime(4).Caption = "3:20-4:10"
lblaftsec(4).Caption = RS1.Fields(0)
Case "3:00-3:50":
lblafttime(4).Caption = "3:50-4:10"
lblaftsec(4).Caption = "Break Time"
lblafttime(3).Caption = "3:00-3:50"
lblaftsec(3).Caption = RS1.Fields(0)
End Select
Select Case RS1.Fields(4)
Case "1st": lblafts(1).BackColor = &H8080FF
Case "2nd": lblafts(2).BackColor = &H8080FF
Case "3rd": lblafts(3).BackColor = &H8080FF
Case "4th": lblafts(4).BackColor = &H8080FF
If Right(RS1.Fields(5), 2) = "50" Then
lblafts(0).BackColor = &H8080FF
Else
lblafts(8).BackColor = &H8080FF
End If
Case "5th": lblafts(5).BackColor = &H8080FF
Case "6th": lblafts(6).BackColor = &H8080FF
Case "7th": lblafts(7).BackColor = &H8080FF
End Select
End If
RS1.MoveNext
Loop
End If
RS1.Close
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblsched Where teacherid='" & teacherids & "'"
If RS.EOF = False Then
Do Until RS.EOF

Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblsubjsched"
If RS1.EOF = False Then
Do Until RS1.EOF
If RS1.Fields(0) = RS.Fields(0) Then
Select Case RS1.Fields(3)
Case "TRS1-E2": Img3(1).Picture = ImageList1(0).ListImages(1).Picture
Case "TRS1-3": Img3(2).Picture = ImageList1(0).ListImages(2).Picture
Case "TRS1-2": Img3(3).Picture = ImageList1(0).ListImages(3).Picture
Case "TRS1-1": Img3(4).Picture = ImageList1(0).ListImages(4).Picture
Case "SEDP-1": Img3(5).Picture = ImageList1(0).ListImages(5).Picture
Case "SEDP-2": Img3(6).Picture = ImageList1(0).ListImages(6).Picture
Case "SEDP-3": Img3(7).Picture = ImageList1(0).ListImages(7).Picture
Case "SEDP-4": Img3(8).Picture = ImageList1(0).ListImages(8).Picture
Case "SEDP-5": Img3(9).Picture = ImageList1(0).ListImages(9).Picture
Case "SEDP-6": Img3(10).Picture = ImageList1(0).ListImages(10).Picture
Case "TRS1-E1": Img3(11).Picture = ImageList1(0).ListImages(11).Picture
Case "TRS1-6": Img3(12).Picture = ImageList1(0).ListImages(12).Picture
Case "TRS1-5": Img3(13).Picture = ImageList1(0).ListImages(13).Picture
Case "TRS1-4": Img3(14).Picture = ImageList1(0).ListImages(14).Picture
Case "TRS2-1": Img3(15).Picture = ImageList1(0).ListImages(15).Picture
Case "TRS2-2": Img3(16).Picture = ImageList1(0).ListImages(16).Picture
Case "SEDP-7": Img3(17).Picture = ImageList1(0).ListImages(17).Picture
Case "SEDP-8": Img3(18).Picture = ImageList1(0).ListImages(18).Picture
Case "SEDP-9": Img3(19).Picture = ImageList1(0).ListImages(19).Picture
Case "SEDP-10": Img3(20).Picture = ImageList1(0).ListImages(20).Picture
Case "SEDP-11": Img3(21).Picture = ImageList1(0).ListImages(21).Picture
End Select
End If
RS1.MoveNext
Loop
RS1.Close
End If
RS.MoveNext
Loop
RS.Close
Else
End If
If usrsubject = "Social Studies" Or usrsubject = "MAPEH" Or usrsubject = "TLE" Then
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblvalsched Where substeacherid='" & teacherids & "'"
If RS.EOF = False Then
Do Until RS.EOF
If days = RS.Fields("day") Then
Select Case RS.Fields(3)
Case "TRS1-E2": Img3(1).Picture = ImageList1(1).ListImages(1).Picture
Case "TRS1-3": Img3(2).Picture = ImageList1(1).ListImages(2).Picture
Case "TRS1-2": Img3(3).Picture = ImageList1(1).ListImages(3).Picture
Case "TRS1-1": Img3(4).Picture = ImageList1(1).ListImages(4).Picture
Case "SEDP-1": Img3(5).Picture = ImageList1(1).ListImages(5).Picture
Case "SEDP-2": Img3(6).Picture = ImageList1(1).ListImages(6).Picture
Case "SEDP-3": Img3(7).Picture = ImageList1(1).ListImages(7).Picture
Case "SEDP-4": Img3(8).Picture = ImageList1(1).ListImages(8).Picture
Case "SEDP-5": Img3(9).Picture = ImageList1(1).ListImages(9).Picture
Case "SEDP-6": Img3(10).Picture = ImageList1(1).ListImages(10).Picture
Case "TRS1-E1": Img3(11).Picture = ImageList1(1).ListImages(11).Picture
Case "TRS1-6": Img3(12).Picture = ImageList1(1).ListImages(12).Picture
Case "TRS1-5": Img3(13).Picture = ImageList1(1).ListImages(13).Picture
Case "TRS1-4": Img3(14).Picture = ImageList1(1).ListImages(14).Picture
Case "TRS2-1": Img3(15).Picture = ImageList1(1).ListImages(15).Picture
Case "TRS2-2": Img3(16).Picture = ImageList1(1).ListImages(16).Picture
Case "SEDP-7": Img3(17).Picture = ImageList1(1).ListImages(17).Picture
Case "SEDP-8": Img3(18).Picture = ImageList1(1).ListImages(18).Picture
Case "SEDP-9": Img3(19).Picture = ImageList1(1).ListImages(19).Picture
Case "SEDP-10": Img3(20).Picture = ImageList1(1).ListImages(20).Picture
Case "SEDP-11": Img3(21).Picture = ImageList1(1).ListImages(21).Picture
End Select
If RS.Fields(1) = "Morning" Then
Select Case RS.Fields(5)
Case "6:00-6:50": lblmornsec(0).Caption = "No Class"
Case "6:50-7:40": lblmornsec(1).Caption = "No Class"
Case "7:40-8:30": lblmornsec(2).Caption = "No Class"
Case "9:40-10:30": lblmornsec(5).Caption = "No Class"
Case "10:30-11:20": lblmornsec(6).Caption = "No Class"
Case "11:20-12:10": lblmornsec(7).Caption = "No Class"
Case "8:50-9:40":
lblmorntime(3).Caption = "8:30-8:50"
lblmornsec(3).Caption = "Break Time"
lblmorntime(4).Caption = "8:50-9:40"
lblmornsec(4).Caption = "No Class"
Case "8:30-9:20":
lblmorntime(4).Caption = "9:20-9:40"
lblmornsec(4).Caption = "Break Time"
lblmorntime(3).Caption = "8:30-9:20"
lblmornsec(3).Caption = "No Class"
End Select
Select Case RS.Fields(4)
Case "1st": lblmorns(1).BackColor = &HFFFFC0
Case "2nd": lblmorns(2).BackColor = &HFFFFC0
Case "3rd": lblmorns(3).BackColor = &HFFFFC0
Case "4th":
lblmorns(4).BackColor = &HFFFFC0
If Right(RS1.Fields(5), 2) = "20" Then
lblmorns(0).BackColor = &HFFFFC0
Else
lblmorns(8).BackColor = &HFFFFC0
End If
Case "5th": lblmorns(5).BackColor = &HFFFFC0
Case "6th": lblmorns(6).BackColor = &HFFFFC0
Case "7th": lblmorns(7).BackColor = &HFFFFC0
End Select
Else
Select Case RS.Fields(5)
Case "12:30-1:20": lblaftsec(0).Caption = "No Class"
Case "1:20-2:10": lblaftsec(1).Caption = "No Class"
Case "2:10-3:00": lblaftsec(2).Caption = "No Class"
Case "4:10-5:00": lblaftsec(5).Caption = "No Class"
Case "5:00-5:50": lblaftsec(6).Caption = "No Class"
Case "5:50-6:40": lblaftsec(7).Caption = "No Class"
Case "3:20-4:10":
lblafttime(3).Caption = "3:00-3:20"
lblaftsec(3).Caption = "Break Time"
lblafttime(4).Caption = "3:20-4:10"
lblaftsec(4).Caption = "No Class"
Case "3:00-3:50":
lblafttime(4).Caption = "3:50-4:10"
lblaftsec(4).Caption = "Break Time"
lblafttime(3).Caption = "3:00-3:50"
lblaftsec(3).Caption = "No Class"
End Select
Select Case RS.Fields(4)
Case "1st": lblafts(1).BackColor = &HFFFFC0
Case "2nd": lblafts(2).BackColor = &HFFFFC0
Case "3rd": lblafts(3).BackColor = &HFFFFC0
Case "4th":
lblafts(4).BackColor = &HFFFFC0
If Right(RS1.Fields(5), 2) = "20" Then
lblafts(0).BackColor = &HFFFFC0
Else
lblafts(8).BackColor = &HFFFFC0
End If
Case "5th": lblafts(5).BackColor = &HFFFFC0
Case "6th": lblafts(6).BackColor = &HFFFFC0
Case "7th": lblafts(7).BackColor = &HFFFFC0
End Select
End If
End If
RS.MoveNext
Loop
End If
RS.Close
End If
Else
Call loadval
End If

If Img3(15).Picture = ImageList1(0).ListImages(15).Picture And Img3(16).Picture = ImageList1(0).ListImages(16).Picture Then
Img3(15).Picture = ImageList1(2).ListImages(4).Picture
ElseIf Img3(15).Picture = ImageList1(1).ListImages(15).Picture And Img3(16).Picture = ImageList1(0).ListImages(16).Picture Then
Img3(15).Picture = ImageList1(2).ListImages(2).Picture
ElseIf Img3(15).Picture = ImageList1(0).ListImages(15).Picture And Img3(16).Picture = ImageList1(1).ListImages(16).Picture Then
Img3(15).Picture = ImageList1(2).ListImages(3).Picture
ElseIf Img3(15).Picture = ImageList1(1).ListImages(15).Picture And Img3(16).Picture = ImageList1(1).ListImages(16).Picture Then
Img3(15).Picture = ImageList1(2).ListImages(1).Picture
End If
End Sub
Sub loadval()
For a = 1 To 21
Img3(a).Picture = ImageList1(1).ListImages(a).Picture
Next a

Text3.Text = "No Advisory Class"
For a = 0 To 8
lblmorns(a).BackColor = &HFFFFC0
lblafts(a).BackColor = &HFFFFC0
Next a
Call dbConnection
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblvalsched Where teacherid='" & teacherids & "' And day='" & days & "'"
If RS1.EOF = False Then
Do Until RS1.EOF
If RS1.Fields(1) = "Morning" Then
Select Case RS1.Fields(5)
Case "6:00-6:50": lblmornsec(0).Caption = RS1.Fields(0)
Case "6:50-7:40": lblmornsec(1).Caption = RS1.Fields(0)
Case "7:40-8:30": lblmornsec(2).Caption = RS1.Fields(0)
Case "9:40-10:30": lblmornsec(5).Caption = RS1.Fields(0)
Case "10:30-11:20": lblmornsec(6).Caption = RS1.Fields(0)
Case "11:20-12:10": lblmornsec(7).Caption = RS1.Fields(0)
Case "8:50-9:40":
lblmorntime(3).Caption = "8:30-8:50"
lblmornsec(3).Caption = "Break Time"
lblmorntime(4).Caption = "8:50-9:40"
lblmornsec(4).Caption = RS1.Fields(0)
Case "8:30-9:20":
lblmorntime(4).Caption = "9:20-9:40"
lblmornsec(4).Caption = "Break Time"
lblmorntime(3).Caption = "8:30-9:20"
lblmornsec(3).Caption = RS1.Fields(0)
End Select
Select Case RS1.Fields(4)
Case "1st": lblmorns(1).BackColor = &H8080FF
Case "2nd": lblmorns(2).BackColor = &H8080FF
Case "3rd": lblmorns(3).BackColor = &H8080FF
Case "4th":
lblmorns(4).BackColor = &H8080FF
If Right(RS1.Fields(5), 2) = "20" Then
lblmorns(0).BackColor = &H8080FF
Else
lblmorns(8).BackColor = &H8080FF
End If
Case "5th": lblmorns(5).BackColor = &H8080FF
Case "6th": lblmorns(6).BackColor = &H8080FF
Case "7th": lblmorns(7).BackColor = &H8080FF
End Select
Else
Select Case RS1.Fields(5)
Case "12:30-1:20": lblaftsec(0).Caption = RS1.Fields(0)
Case "1:20-2:10": lblaftsec(1).Caption = RS1.Fields(0)
Case "2:10-3:00": lblaftsec(2).Caption = RS1.Fields(0)
Case "4:10-5:00": lblaftsec(5).Caption = RS1.Fields(0)
Case "5:00-5:50": lblaftsec(6).Caption = RS1.Fields(0)
Case "5:50-6:40": lblaftsec(7).Caption = RS1.Fields(0)
Case "3:20-4:10":
lblafttime(3).Caption = "3:00-3:20"
lblaftsec(3).Caption = "Break Time"
lblafttime(4).Caption = "3:20-4:10"
lblaftsec(4).Caption = RS1.Fields(0)
Case "3:00-3:50":
lblafttime(4).Caption = "3:50-4:10"
lblaftsec(4).Caption = "Break Time"
lblafttime(3).Caption = "3:00-3:50"
lblaftsec(3).Caption = RS1.Fields(0)
End Select
Select Case RS1.Fields(4)
Case "1st": lblafts(1).BackColor = &H8080FF
Case "2nd": lblafts(2).BackColor = &H8080FF
Case "3rd": lblafts(3).BackColor = &H8080FF
Case "4th": lblafts(4).BackColor = &H8080FF
If Right(RS1.Fields(5), 2) = "50" Then
lblafts(0).BackColor = &H8080FF
Else
lblafts(8).BackColor = &H8080FF
End If
Case "5th": lblafts(5).BackColor = &H8080FF
Case "6th": lblafts(6).BackColor = &H8080FF
Case "7th": lblafts(7).BackColor = &H8080FF
End Select
End If
RS1.MoveNext
Loop
End If
RS1.Close
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblvalsched Where teacherid='" & teacherids & "' and day='" & days & "'"
If RS.EOF = False Then
Do Until RS.EOF

Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblsubjsched"
If RS1.EOF = False Then
Do Until RS1.EOF
If RS1.Fields(0) = RS.Fields(0) Then
Select Case RS1.Fields(3)
Case "TRS1-E2": Img3(1).Picture = ImageList1(0).ListImages(1).Picture
Case "TRS1-3": Img3(2).Picture = ImageList1(0).ListImages(2).Picture
Case "TRS1-2": Img3(3).Picture = ImageList1(0).ListImages(3).Picture
Case "TRS1-1": Img3(4).Picture = ImageList1(0).ListImages(4).Picture
Case "SEDP-1": Img3(5).Picture = ImageList1(0).ListImages(5).Picture
Case "SEDP-2": Img3(6).Picture = ImageList1(0).ListImages(6).Picture
Case "SEDP-3": Img3(7).Picture = ImageList1(0).ListImages(7).Picture
Case "SEDP-4": Img3(8).Picture = ImageList1(0).ListImages(8).Picture
Case "SEDP-5": Img3(9).Picture = ImageList1(0).ListImages(9).Picture
Case "SEDP-6": Img3(10).Picture = ImageList1(0).ListImages(10).Picture
Case "TRS1-E1": Img3(11).Picture = ImageList1(0).ListImages(11).Picture
Case "TRS1-6": Img3(12).Picture = ImageList1(0).ListImages(12).Picture
Case "TRS1-5": Img3(13).Picture = ImageList1(0).ListImages(13).Picture
Case "TRS1-4": Img3(14).Picture = ImageList1(0).ListImages(14).Picture
Case "TRS2-1": Img3(15).Picture = ImageList1(0).ListImages(15).Picture
Case "TRS2-2": Img3(16).Picture = ImageList1(0).ListImages(16).Picture
Case "SEDP-7": Img3(17).Picture = ImageList1(0).ListImages(17).Picture
Case "SEDP-8": Img3(18).Picture = ImageList1(0).ListImages(18).Picture
Case "SEDP-9": Img3(19).Picture = ImageList1(0).ListImages(19).Picture
Case "SEDP-10": Img3(20).Picture = ImageList1(0).ListImages(20).Picture
Case "SEDP-11": Img3(21).Picture = ImageList1(0).ListImages(21).Picture
End Select
End If
RS1.MoveNext
Loop
RS1.Close
End If
RS.MoveNext
Loop
RS.Close
Else
End If
End Sub
