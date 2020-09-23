VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmstudentinfo 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Information"
   ClientHeight    =   10740
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   15270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmstudendinfo.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10740
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3600
      Top             =   1320
   End
   Begin VB.PictureBox pic21 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   7845
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   3855
      TabIndex        =   63
      Top             =   2520
      Width           =   3885
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Enabled         =   0   'False
         Height          =   7215
         Left            =   0
         TabIndex        =   85
         Top             =   600
         Width           =   4215
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
            Height          =   360
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   98
            Top             =   285
            Width           =   1335
         End
         Begin VB.TextBox txt5 
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
            Left            =   1440
            TabIndex        =   97
            Top             =   3405
            Width           =   2175
         End
         Begin VB.TextBox txt7 
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
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   96
            Top             =   3885
            Width           =   495
         End
         Begin VB.TextBox txt9 
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
            Left            =   1440
            TabIndex        =   95
            Top             =   5325
            Width           =   2175
         End
         Begin VB.TextBox txt8 
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
            Left            =   1440
            TabIndex        =   94
            Top             =   4845
            Width           =   2175
         End
         Begin VB.TextBox txt10 
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
            Left            =   1440
            TabIndex        =   93
            Top             =   5805
            Width           =   2175
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
            Left            =   1440
            TabIndex        =   92
            Top             =   1965
            Width           =   2175
         End
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
            Left            =   1440
            TabIndex        =   91
            Top             =   2445
            Width           =   2175
         End
         Begin VB.TextBox txt4 
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
            Left            =   1440
            TabIndex        =   90
            Top             =   2925
            Width           =   2175
         End
         Begin VB.ComboBox Combo1 
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
            ItemData        =   "frmstudendinfo.frx":164A
            Left            =   1440
            List            =   "frmstudendinfo.frx":1654
            TabIndex        =   89
            Text            =   "Select Gender.."
            Top             =   4365
            Width           =   2175
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
            Left            =   1800
            TabIndex        =   88
            Top             =   6765
            Width           =   1815
         End
         Begin VB.TextBox txtyr 
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
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   87
            Top             =   885
            Width           =   735
         End
         Begin VB.TextBox txtsec 
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
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   1365
            Width           =   735
         End
         Begin MSMask.MaskEdBox txt6 
            Height          =   375
            Left            =   1440
            TabIndex        =   99
            Top             =   3885
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777088
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt11 
            Height          =   375
            Left            =   1800
            TabIndex        =   100
            Top             =   6285
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777088
            MaxLength       =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "#########"
            PromptChar      =   "_"
         End
         Begin VB.Label Label21 
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
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   116
            Top             =   45
            Width           =   1095
         End
         Begin VB.Label Label20 
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
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   115
            Top             =   6405
            Width           =   1215
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Birthdate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   114
            Top             =   3885
            Width           =   1095
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   2640
            TabIndex        =   113
            Top             =   3885
            Width           =   1455
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Relation"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   112
            Top             =   5400
            Width           =   1095
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Guardian"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   111
            Top             =   4845
            Width           =   1815
         End
         Begin VB.Label Label12 
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
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   110
            Top             =   5925
            Width           =   1455
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Surname"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   109
            Top             =   1965
            Width           =   1095
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "FirstName"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   108
            Top             =   2445
            Width           =   1455
         End
         Begin VB.Label Label9 
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
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   107
            Top             =   3405
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Middle Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   106
            Top             =   2925
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "+639"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   495
            Index           =   6
            Left            =   1365
            TabIndex        =   105
            Top             =   6315
            Width           =   1215
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   104
            Top             =   4365
            Width           =   1095
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Previous School"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   103
            Top             =   6885
            Width           =   1695
         End
         Begin VB.Label Label32 
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
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   102
            Top             =   885
            Width           =   1095
         End
         Begin VB.Label Label39 
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
            ForeColor       =   &H00400000&
            Height          =   495
            Left            =   120
            TabIndex        =   101
            Top             =   1365
            Width           =   1095
         End
         Begin VB.Shape Shape1 
            Height          =   1815
            Left            =   1800
            Top             =   0
            Width           =   1815
         End
         Begin VB.Image Image5 
            BorderStyle     =   1  'Fixed Single
            Height          =   1830
            Left            =   1800
            Picture         =   "frmstudendinfo.frx":1666
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1830
         End
      End
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   4
         Left            =   60
         ScaleHeight     =   690
         ScaleWidth      =   2535
         TabIndex        =   67
         Top             =   9360
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   4
            Left            =   315
            Picture         =   "frmstudendinfo.frx":1D38
            ScaleHeight     =   315
            ScaleWidth      =   1905
            TabIndex        =   68
            Top             =   240
            Width           =   1935
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Exit"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   4
               Left            =   840
               TabIndex        =   69
               Top             =   0
               Visible         =   0   'False
               Width           =   285
            End
         End
      End
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   840
         Index           =   7
         Left            =   60
         ScaleHeight     =   810
         ScaleWidth      =   2535
         TabIndex        =   64
         Top             =   8880
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   5
            Left            =   435
            Picture         =   "frmstudendinfo.frx":69C60
            ScaleHeight     =   315
            ScaleWidth      =   1905
            TabIndex        =   65
            Top             =   240
            Width           =   1935
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Lock this system"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   5
               Left            =   240
               TabIndex        =   66
               Top             =   0
               Visible         =   0   'False
               Width           =   1440
            End
         End
         Begin VB.Image Image8 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.Label lbllogout 
         BackStyle       =   0  'Transparent
         Caption         =   "Log-Out"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   3000
         TabIndex        =   71
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   2
         Left            =   -15
         Picture         =   "frmstudendinfo.frx":D1B88
         Stretch         =   -1  'True
         Top             =   11175
         Width           =   2730
      End
      Begin VB.Label lbl21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Student Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   70
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image Image9 
         Height          =   495
         Left            =   0
         Picture         =   "frmstudendinfo.frx":D4CBF
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4875
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3000
      Top             =   1320
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   7845
      Left            =   11040
      ScaleHeight     =   7815
      ScaleWidth      =   4365
      TabIndex        =   7
      Top             =   2520
      Width           =   4395
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
         Height          =   360
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   6200
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1575
         Left            =   0
         TabIndex        =   50
         Top             =   6600
         Width           =   4575
         Begin VB.Label lblvaltime 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
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
            Left            =   90
            TabIndex        =   59
            Top             =   870
            Width           =   1000
         End
         Begin VB.Label lblvalsubject 
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
            Left            =   1140
            TabIndex        =   58
            Top             =   870
            Width           =   1395
         End
         Begin VB.Label lblvalday 
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
            Left            =   2580
            TabIndex        =   57
            Top             =   870
            Width           =   1695
         End
         Begin VB.Label lblvalsubject 
            BackColor       =   &H00FFFFC0&
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
            Left            =   1140
            TabIndex        =   56
            Top             =   510
            Width           =   1395
         End
         Begin VB.Label labelvaltime 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
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
            Left            =   90
            TabIndex        =   55
            Top             =   150
            Width           =   1000
         End
         Begin VB.Label labelvalsubject 
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
            Left            =   1140
            TabIndex        =   54
            Top             =   150
            Width           =   1395
         End
         Begin VB.Label Labelvalday 
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
            Left            =   2580
            TabIndex        =   53
            Top             =   150
            Width           =   1695
         End
         Begin VB.Label lblvaltime 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
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
            Left            =   90
            TabIndex        =   52
            Top             =   510
            Width           =   1000
         End
         Begin VB.Label lblvalday 
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
            Index           =   1
            Left            =   2580
            TabIndex        =   51
            Top             =   510
            Width           =   1695
         End
         Begin VB.Image Image13 
            Height          =   105
            Left            =   0
            Picture         =   "frmstudendinfo.frx":D7F94
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4995
         End
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
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1920
         Width           =   3135
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1440
         Width           =   1215
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
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1440
         Width           =   1215
      End
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   1
         Left            =   60
         ScaleHeight     =   690
         ScaleWidth      =   2535
         TabIndex        =   11
         Top             =   9360
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   315
            Picture         =   "frmstudendinfo.frx":DB269
            ScaleHeight     =   315
            ScaleWidth      =   1905
            TabIndex        =   12
            Top             =   240
            Width           =   1935
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Exit"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   1
               Left            =   840
               TabIndex        =   13
               Top             =   0
               Visible         =   0   'False
               Width           =   285
            End
         End
      End
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   840
         Index           =   0
         Left            =   60
         ScaleHeight     =   810
         ScaleWidth      =   2535
         TabIndex        =   8
         Top             =   8880
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   435
            Picture         =   "frmstudendinfo.frx":143191
            ScaleHeight     =   315
            ScaleWidth      =   1905
            TabIndex        =   9
            Top             =   240
            Width           =   1935
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Lock this system"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   0
               Left            =   240
               TabIndex        =   10
               Top             =   0
               Visible         =   0   'False
               Width           =   1440
            End
         End
         Begin VB.Image Image2 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3400
         Left            =   0
         TabIndex        =   14
         Top             =   2520
         Width           =   4575
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Index           =   8
            Left            =   90
            TabIndex        =   119
            Top             =   3030
            Width           =   1000
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00C0FFFF&
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
            Index           =   8
            Left            =   1140
            TabIndex        =   118
            Top             =   3030
            Width           =   1395
         End
         Begin VB.Label lblteacher 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Index           =   8
            Left            =   2580
            TabIndex        =   117
            Top             =   3030
            Width           =   1695
         End
         Begin VB.Image Image11 
            Height          =   105
            Left            =   0
            Picture         =   "frmstudendinfo.frx":1AB0B9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4995
         End
         Begin VB.Line Line2 
            BorderStyle     =   5  'Dash-Dot-Dot
            BorderWidth     =   5
            X1              =   0
            X2              =   3500
            Y1              =   45
            Y2              =   45
         End
         Begin VB.Label lblteacher 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   2580
            TabIndex        =   38
            Top             =   510
            Width           =   1695
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   90
            TabIndex        =   37
            Top             =   510
            Width           =   1000
         End
         Begin VB.Label labelteacher 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   2580
            TabIndex        =   36
            Top             =   150
            Width           =   1695
         End
         Begin VB.Label labelsubject 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   1140
            TabIndex        =   35
            Top             =   150
            Width           =   1395
         End
         Begin VB.Label labeltime 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   90
            TabIndex        =   34
            Top             =   150
            Width           =   1000
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00C0FFFF&
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
            Left            =   1140
            TabIndex        =   33
            Top             =   510
            Width           =   1395
         End
         Begin VB.Label lblteacher 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   2580
            TabIndex        =   32
            Top             =   1230
            Width           =   1695
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   90
            TabIndex        =   31
            Top             =   1230
            Width           =   1000
         End
         Begin VB.Label lblteacher 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   2580
            TabIndex        =   30
            Top             =   870
            Width           =   1695
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00C0FFFF&
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
            Left            =   1140
            TabIndex        =   29
            Top             =   870
            Width           =   1395
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   90
            TabIndex        =   28
            Top             =   870
            Width           =   1000
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00C0FFFF&
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
            Left            =   1140
            TabIndex        =   27
            Top             =   1230
            Width           =   1395
         End
         Begin VB.Label lblteacher 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   2580
            TabIndex        =   26
            Top             =   1950
            Width           =   1695
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   90
            TabIndex        =   25
            Top             =   1950
            Width           =   1000
         End
         Begin VB.Label lblteacher 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   2580
            TabIndex        =   24
            Top             =   1590
            Width           =   1695
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00C0FFFF&
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
            Left            =   1140
            TabIndex        =   23
            Top             =   1590
            Width           =   1395
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   90
            TabIndex        =   22
            Top             =   1590
            Width           =   1000
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00C0FFFF&
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
            Left            =   1140
            TabIndex        =   21
            Top             =   1950
            Width           =   1395
         End
         Begin VB.Label lblteacher 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   2580
            TabIndex        =   20
            Top             =   2670
            Width           =   1695
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   90
            TabIndex        =   19
            Top             =   2670
            Width           =   1000
         End
         Begin VB.Label lblteacher 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   2580
            TabIndex        =   18
            Top             =   2310
            Width           =   1695
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00C0FFFF&
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
            Left            =   1140
            TabIndex        =   17
            Top             =   2310
            Width           =   1395
         End
         Begin VB.Label lbltime 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   90
            TabIndex        =   16
            Top             =   2310
            Width           =   1000
         End
         Begin VB.Label lblsubject 
            BackColor       =   &H00C0FFFF&
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
            Left            =   1140
            TabIndex        =   15
            Top             =   2670
            Width           =   1395
         End
      End
      Begin TNHSES.lvButtons_H lvButtons_H1 
         Height          =   375
         Left            =   2640
         TabIndex        =   72
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Print Schedule"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648384
         cGradient       =   12648384
         Gradient        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777152
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Values Teacher:"
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
         Left            =   0
         TabIndex        =   61
         Top             =   6340
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Values Class Schedule"
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
         Left            =   0
         TabIndex        =   60
         Top             =   5950
         Width           =   2175
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Section Schedule "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   49
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Adviser"
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
         TabIndex        =   48
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label27 
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
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2280
         TabIndex        =   46
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Class Schedule"
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
         Left            =   0
         TabIndex        =   45
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Image Image17 
         Height          =   105
         Left            =   0
         Picture         =   "frmstudendinfo.frx":1AE38E
         Stretch         =   -1  'True
         Top             =   1250
         Width           =   4995
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2280
         TabIndex        =   43
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "School Year:2010-2011"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Session"
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
         TabIndex        =   40
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   0
         Left            =   -15
         Picture         =   "frmstudendinfo.frx":1B1663
         Stretch         =   -1  'True
         Top             =   11175
         Width           =   2730
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   0
         Picture         =   "frmstudendinfo.frx":1B479A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7155
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6960
      Top             =   5160
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   7845
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   15735
      TabIndex        =   0
      Top             =   2520
      Width           =   15765
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Enabled         =   0   'False
         Height          =   495
         Left            =   0
         TabIndex        =   80
         Top             =   0
         Width           =   11655
         Begin VB.Label lblfloor23 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   9720
            TabIndex        =   84
            Top             =   120
            Width           =   135
         End
         Begin VB.Label lblfloor2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "2nd Floor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   9840
            TabIndex        =   83
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblfloor1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "1st Floor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   8640
            TabIndex        =   82
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label34 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Map"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3960
            TabIndex        =   81
            Top             =   120
            Width           =   3015
         End
         Begin VB.Image Image10 
            Height          =   495
            Left            =   1560
            Picture         =   "frmstudendinfo.frx":1B7A6F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   21315
         End
      End
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   3
         Left            =   60
         ScaleHeight     =   690
         ScaleWidth      =   2535
         TabIndex        =   4
         Top             =   9360
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   3
            Left            =   315
            Picture         =   "frmstudendinfo.frx":1BAD44
            ScaleHeight     =   315
            ScaleWidth      =   1905
            TabIndex        =   5
            Top             =   240
            Width           =   1935
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Exit"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   3
               Left            =   840
               TabIndex        =   6
               Top             =   0
               Visible         =   0   'False
               Width           =   285
            End
         End
      End
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   840
         Index           =   2
         Left            =   60
         ScaleHeight     =   810
         ScaleWidth      =   2535
         TabIndex        =   1
         Top             =   8880
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   2
            Left            =   435
            Picture         =   "frmstudendinfo.frx":222C6C
            ScaleHeight     =   315
            ScaleWidth      =   1905
            TabIndex        =   2
            Top             =   240
            Width           =   1935
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Lock this system"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   2
               Left            =   240
               TabIndex        =   3
               Top             =   0
               Visible         =   0   'False
               Width           =   1440
            End
         End
         Begin VB.Image Image6 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.Frame frm1st 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   7575
         Left            =   1560
         TabIndex        =   120
         Top             =   240
         Width           =   14655
         Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf13 
            Height          =   3855
            Left            =   2280
            TabIndex        =   123
            Top             =   3960
            Width           =   3615
            _cx             =   6376
            _cy             =   6800
            FlashVars       =   ""
            Movie           =   "Chfgh:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\flash2.swf"
            Src             =   "Chfgh:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\flash2.swf"
            WMode           =   "Window"
            Play            =   -1  'True
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
         Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf11 
            Height          =   3855
            Left            =   2280
            TabIndex        =   124
            Top             =   240
            Width           =   3615
            _cx             =   6376
            _cy             =   6800
            FlashVars       =   ""
            Movie           =   "C:\Documents and Settings\Vargas\My Documents\TNHSgfdhf Enrollment System\flash\flash2.swf"
            Src             =   "C:\Documents and Settings\Vargas\My Documents\TNHSgfdhf Enrollment System\flash\flash2.swf"
            WMode           =   "Window"
            Play            =   -1  'True
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
         Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf14 
            Height          =   3855
            Left            =   5880
            TabIndex        =   125
            Top             =   3960
            Width           =   3615
            _cx             =   6376
            _cy             =   6800
            FlashVars       =   ""
            Movie           =   "hfghfg"
            Src             =   "hfghfg"
            WMode           =   "Window"
            Play            =   -1  'True
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
         Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf12 
            Height          =   3855
            Left            =   5895
            TabIndex        =   126
            Top             =   240
            Width           =   3615
            _cx             =   6376
            _cy             =   6800
            FlashVars       =   ""
            Movie           =   "C:\hfghDocuments and Settings\Vargas\My Documents\TNHS Enrollment System\flash\flash2.swf"
            Src             =   "C:\hfghDocuments and Settings\Vargas\My Documents\TNHS Enrollment System\flash\flash2.swf"
            WMode           =   "Window"
            Play            =   -1  'True
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
         Begin VB.Frame frm1stfloor 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   0  'None
            Height          =   7665
            Left            =   1080
            TabIndex        =   121
            Top             =   240
            Width           =   9735
            Begin VB.Image Image3 
               Height          =   1200
               Index           =   1
               Left            =   7000
               Picture         =   "frmstudendinfo.frx":28AB94
               ToolTipText     =   "TRS1-E2"
               Top             =   5670
               Width           =   1155
            End
            Begin VB.Image Image3 
               Height          =   1260
               Index           =   2
               Left            =   6825
               Picture         =   "frmstudendinfo.frx":28B47B
               ToolTipText     =   "TRS1-3"
               Top             =   4440
               Width           =   1110
            End
            Begin VB.Image Image3 
               Height          =   1170
               Index           =   3
               Left            =   6595
               Picture         =   "frmstudendinfo.frx":28BCF9
               ToolTipText     =   "TRS1-2"
               Top             =   3300
               Width           =   1080
            End
            Begin VB.Image Image3 
               Height          =   1005
               Index           =   4
               Left            =   6435
               Picture         =   "frmstudendinfo.frx":28C566
               ToolTipText     =   "TRS1-1"
               Top             =   2310
               Width           =   975
            End
            Begin VB.Image Image3 
               Height          =   690
               Index           =   7
               Left            =   3610
               Picture         =   "frmstudendinfo.frx":28CCCF
               ToolTipText     =   "SEDP-3"
               Top             =   285
               Width           =   540
            End
            Begin VB.Image Image3 
               Height          =   735
               Index           =   10
               Left            =   6010
               Picture         =   "frmstudendinfo.frx":28D1B8
               ToolTipText     =   "SEDP-6"
               Top             =   285
               Width           =   915
            End
            Begin VB.Image Image3 
               Height          =   495
               Index           =   11
               Left            =   6015
               Picture         =   "frmstudendinfo.frx":28D89B
               ToolTipText     =   "TRS1-E1"
               Top             =   1410
               Width           =   1095
            End
            Begin VB.Image Image3 
               Height          =   705
               Index           =   6
               Left            =   2280
               Picture         =   "frmstudendinfo.frx":28DE58
               ToolTipText     =   "SEDP-2"
               Top             =   285
               Width           =   735
            End
            Begin VB.Image Image3 
               Height          =   705
               Index           =   5
               Left            =   1740
               ToolTipText     =   "SEDP-1"
               Top             =   285
               Width           =   645
            End
            Begin VB.Image Image3 
               Height          =   705
               Index           =   9
               Left            =   4910
               Picture         =   "frmstudendinfo.frx":28E3FF
               ToolTipText     =   "SEDP-5"
               Top             =   270
               Width           =   960
            End
            Begin VB.Image Image3 
               Height          =   705
               Index           =   8
               Left            =   4135
               ToolTipText     =   "SEDP-4"
               Top             =   270
               Width           =   900
            End
            Begin VB.Image Image7 
               Height          =   10500
               Left            =   -720
               Picture         =   "frmstudendinfo.frx":28EA51
               Top             =   0
               Width           =   10500
            End
         End
         Begin VB.Frame frm2ndfloor 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   7740
            Left            =   1080
            TabIndex        =   122
            Top             =   240
            Width           =   9015
            Begin VB.Image Image3 
               Height          =   1245
               Index           =   12
               Left            =   6810
               Picture         =   "frmstudendinfo.frx":2ABF58
               Top             =   4470
               Width           =   1095
            End
            Begin VB.Image Image3 
               Height          =   1170
               Index           =   13
               Left            =   6630
               Picture         =   "frmstudendinfo.frx":2AC7ED
               Top             =   3310
               Width           =   1035
            End
            Begin VB.Image Image3 
               Height          =   1020
               Index           =   14
               Left            =   6430
               Picture         =   "frmstudendinfo.frx":2AD02C
               Top             =   2310
               Width           =   1110
            End
            Begin VB.Image Image3 
               Height          =   735
               Index           =   21
               Left            =   6050
               Picture         =   "frmstudendinfo.frx":2AD7C3
               Top             =   270
               Width           =   915
            End
            Begin VB.Image Image3 
               Height          =   1650
               Index           =   16
               Left            =   3840
               Picture         =   "frmstudendinfo.frx":2ADE20
               Top             =   3450
               Width           =   1530
            End
            Begin VB.Image Image3 
               Height          =   1455
               Index           =   15
               Left            =   3050
               Picture         =   "frmstudendinfo.frx":2AEB06
               Top             =   2640
               Width           =   1380
            End
            Begin VB.Image Image3 
               Height          =   735
               Index           =   17
               Left            =   1670
               Picture         =   "frmstudendinfo.frx":2AF7DB
               Top             =   290
               Width           =   675
            End
            Begin VB.Image Image3 
               Height          =   735
               Index           =   18
               Left            =   2210
               Picture         =   "frmstudendinfo.frx":2AFDED
               Top             =   290
               Width           =   780
            End
            Begin VB.Image Image3 
               Height          =   720
               Index           =   20
               Left            =   4110
               Picture         =   "frmstudendinfo.frx":2B0410
               Top             =   290
               Width           =   930
            End
            Begin VB.Image Image3 
               Height          =   735
               Index           =   19
               Left            =   3590
               Picture         =   "frmstudendinfo.frx":2B09CC
               Top             =   270
               Width           =   585
            End
            Begin VB.Image Image14 
               Height          =   10500
               Left            =   -720
               Picture         =   "frmstudendinfo.frx":2B0EFF
               Top             =   0
               Width           =   10500
            End
         End
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf1 
         Height          =   8055
         Left            =   3840
         TabIndex        =   73
         Top             =   0
         Width           =   7455
         _cx             =   13150
         _cy             =   14208
         FlashVars       =   ""
         Movie           =   "C:\TNHS Enrollment System\flash\flash.swf"
         Src             =   "C:\TNHS Enrollment System\flash\flash.swf"
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
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   79
         Top             =   0
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   1
         Left            =   -15
         Picture         =   "frmstudendinfo.frx":2CC20E
         Stretch         =   -1  'True
         Top             =   11175
         Width           =   2730
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1575
      Left            =   0
      TabIndex        =   74
      Top             =   1920
      Width           =   16215
      Begin VB.TextBox txtstudno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   75
         Top             =   120
         Width           =   2175
      End
      Begin TNHSES.lvButtons_H command1 
         Height          =   375
         Left            =   4080
         TabIndex        =   76
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Ok"
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Information"
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
         TabIndex        =   77
         Top             =   120
         Width           =   2055
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   1
      Left            =   240
      Top             =   240
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
            Picture         =   "frmstudendinfo.frx":2CF345
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2D3449
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2D70F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2DAFE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2DE84E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2E1991
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2E4873
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2E748D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2E8F4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2EAADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2EDF71
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2F1179
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2F53AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2F918C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":2FCC57
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":301F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":30762A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":30A90D
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":30DB87
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":310875
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":313931
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   615
      Left            =   0
      TabIndex        =   127
      Top             =   1920
      Width           =   15615
      _cx             =   27543
      _cy             =   1085
      FlashVars       =   ""
      Movie           =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\studinfo-header.swf"
      Src             =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\studinfo-header.swf"
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
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   1080
      Top             =   240
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
            Picture         =   "frmstudendinfo.frx":316D43
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31763A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":317EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":318745
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":318EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31947E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":319A35
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":319F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31A53C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31AB9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31B291
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31B85E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31C103
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31C952
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31D0F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31DDDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31EAD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31F0F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31F729
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":31FC6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstudendinfo.frx":320238
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash2 
      Height          =   495
      Left            =   -2280
      TabIndex        =   128
      Top             =   10320
      Width           =   19455
      _cx             =   34316
      _cy             =   873
      FlashVars       =   ""
      Movie           =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\studinfo-footer.swf"
      Src             =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\studinfo-footer.swf"
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
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
   Begin VB.Image Image19 
      Height          =   4020
      Left            =   0
      Picture         =   "frmstudendinfo.frx":3208A5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15330
   End
   Begin VB.Menu admin 
      Caption         =   "Administrator"
      Begin VB.Menu lock 
         Caption         =   "Lock This Form"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit to This Form"
      End
   End
End
Attribute VB_Name = "frmstudentinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yrsect As String, floor As String, login As Date
Private Sub Command1_Click()
Call dbConnection
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblstudents Where studno='" & txtstudno.Text & "'"
If RS.EOF = False Then
If RS.Fields("status") = "Enrolled" Then
MsgBox "Login Success!"
With RS
txt1.Text = .Fields(0)
txt2.Text = .Fields(1)
txt3.Text = .Fields(2)
txt4.Text = .Fields(3)
txt5.Text = .Fields(4)
txt6.Text = Format(.Fields(5), "MM/DD/YYYY")
txt7.Text = Year(Now) - Year(.Fields(5))
Combo1.Text = .Fields("gender")
txt8.Text = .Fields(7)
txt9.Text = .Fields(8)
txt10.Text = .Fields(9)
txt11.Text = Mid(.Fields(12), 5, 9)
Text4.Text = .Fields(18)
txtyr.Text = .Fields("years")
txtsec.Text = .Fields("section")
yrsect = .Fields("yrsec")
Image5.Picture = LoadPicture(App.Path & "/students/" & .Fields(0) & ".jpg")
End With
frm1st.Visible = True
Call viewsched
Label29.Caption = "Section Schedule (" & yrsect & ")"
ShockwaveFlash1.Visible = True
ShockwaveFlash1.Movie = App.Path & "\flash\studinfo-header.swf"
Timer2.Enabled = True
Else
MsgBox "Student Not Enrolled!"
txtstudno.Text = Clear
txtstudno.SetFocus
End If
Else
MsgBox "Student Number does not Exist!"
txtstudno.Text = Clear
txtstudno.SetFocus
End If
RS.Close
End Sub
Sub viewsched()
On Error GoTo erd
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblsubjsched"
RS1.Find ("section='" & yrsect & "'")
If RS1.EOF = False Then
Text2.Text = RS1.Fields(1)
Text3.Text = RS1.Fields(3)
Text5.Text = RS1.Fields(4)
End If
RS1.Close


Dim ch As Integer
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblsched"
Do Until RS1.EOF
If RS1.Fields(0) = yrsect Then
Select Case RS1.Fields("class")
Case "1st": ch = 1
Case "2nd": ch = 2
Case "3rd": ch = 3
Case "4th":
If Mid(yrsect, 1, 1) = 1 Or Mid(yrsect, 1, 1) = 2 Then
ch = 5
    If Text2.Text = "Morning" Then
    lbltime(4).Caption = "8:30-8:50"
    lblteacher(4).Caption = "-------"
    lblsubject(4).Caption = "Break Time"
    Else
    lbltime(4).Caption = "3:00-3:20"
    lblteacher(4).Caption = "-------"
    lblsubject(4).Caption = "Break Time"
    End If
Else
ch = 4
    If Text2.Text = "Afternoon" Then
    lbltime(5).Caption = "9:20-9:40"
    lblteacher(5).Caption = "-------"
    lblsubject(5).Caption = "Break Time"
    Else
    lbltime(5).Caption = "3:50-4:10"
    lblteacher(5).Caption = "-------"
    lblsubject(5).Caption = "Break Time"
    End If
End If
Case "5th": ch = 6
Case "6th": ch = 7
Case "7th": ch = 8
End Select


'MsgBox ch & " " & RS1.Fields("time")
labeltime.Caption = "Time"
labelsubject.Caption = "Subject"
labelteacher.Caption = "Teacher"

    lbltime(ch).Caption = RS1.Fields("time")
    lblteacher(ch).Caption = RS1.Fields("teacher")
    lblsubject(ch).Caption = RS1.Fields("subject")
    
    
    labelvaltime.Caption = "Time"
    labelvalsubject.Caption = "Subs. Subject"
    Labelvalday.Caption = "Day"
    End If
RS1.MoveNext

Loop
ch = 1
RS1.Close

Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblvalsched"
Do Until RS1.EOF
If RS1.Fields(0) = yrsect Then
Text1.Text = RS1.Fields("teacher")
lblvaltime(ch).Caption = RS1.Fields(5)
lblvalsubject(ch).Caption = RS1.Fields(6)
lblvalday(ch).Caption = RS1.Fields(8)
ch = ch + 1
End If
RS1.MoveNext
Loop
RS1.Close
For a = 1 To 21
Image3(a).Picture = imagelist1(0).ListImages(a).Picture
Next a

Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from roommap Where roomname='" & Text3.Text & "'"
If RS1.EOF = False Then
For a = 1 To 21
Image3(a).Visible = False
Next
Image3(Int(RS1.Fields(0))).Visible = True
Image3(Int(RS1.Fields(0))).Picture = imagelist1(1).ListImages(Int(RS1.Fields(0))).Picture

floor = RS1.Fields(4)
Label34.Caption = "Map (" & Text3.Text & ")"
End If
RS1.Close

If floor = "1st" Then
lblfloor1.ForeColor = &HC00000
Else
lblfloor2.ForeColor = &HC00000
End If
Exit Sub
erd:
    Text2.Text = "Not Set"
    Text3.Text = "Not Set"
    Text5.Text = "Not Set"
End Sub
Private Sub exit_Click()

MsgBox "Enter User Password"
adminbool = False
frmadminpass.Show 1
End Sub
Private Sub Form_Load()
frm1st.Visible = False
ShockwaveFlash1.Visible = False
swf11.Movie = App.Path & "\flash\flash2.swf"
swf12.Movie = App.Path & "\flash\flash2.swf"
swf13.Movie = App.Path & "\flash\flash2.swf"
swf14.Movie = App.Path & "\flash\flash2.swf"
swf1.Movie = App.Path & "\flash\flash2.swf"
ShockwaveFlash1.Movie = App.Path & "\flash\studinfo-header.swf"
ShockwaveFlash2.Movie = App.Path & "\flash\studinfo-footer.swf"
ShockwaveFlash1.Play
swf1.Play
swf12.Play
swf13.Play
swf14.Play
txtstudno.TabIndex = 0
End Sub


Private Sub frm1stfloor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call strike
End Sub

Private Sub frm2ndfloor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call strike
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call strike
End Sub
Sub strike()
lbllogout.FontStrikethru = False
lblfloor1.FontStrikethru = False
lblfloor2.FontStrikethru = False
End Sub


Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call strike
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call strike
End Sub

Private Sub imgstud_Click()
viewstudrec = "studinfo"
frmviewstudrec.Show 1
End Sub

Private Sub Label32_Click()
Call Reset
End Sub



Private Sub lbl21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call strike
End Sub

Private Sub lblfloor1_Click()
frm1stfloor.Visible = True
frm2ndfloor.Visible = False
End Sub

Private Sub lblfloor1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call strike
lblfloor1.FontStrikethru = True
End Sub

Private Sub lblfloor2_Click()
frm2ndfloor.Visible = True
frm1stfloor.Visible = False
End Sub

Private Sub lblfloor2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call strike
lblfloor2.FontStrikethru = True
End Sub

Private Sub lbllogout_Click()
If txt1.Text <> Clear Then
If MsgBox("Log out your account?", vbYesNo + vbQuestion) = vbYes Then
Call dbConnection
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from studlog"
        With RS
.AddNew
.Fields(0) = txt1.Text
.Fields(1) = txt3.Text & " " & Mid(txt4.Text, 1, 1) & ". " & txt2.Text
.Fields(2) = Date
.Fields(3) = login
.Fields(4) = Time
.Update
End With
RS.Close
lbllogout.Enabled = False
i_Clear.cLearMe Me
Image5.Picture = frmmain.imagelist1.ListImages(51).Picture
labeltime.Caption = Clear
labelsubject.Caption = Clear
labelteacher.Caption = Clear
labelvaltime.Caption = Clear
labelvalsubject.Caption = Clear
Labelvalday.Caption = Clear
For a = 1 To 8
lbltime(a).Caption = Clear
lblsubject(a).Caption = Clear
lblteacher(a).Caption = Clear
If a < 3 Then
lblvaltime(a).Caption = Clear
lblvalday(a).Caption = Clear
lblvalday(a).Caption = Clear
End If
Next a
Label29.Caption = "Section Schedule"
Timer3.Enabled = True
Label34.Caption = "Map"
lblfloor1.ForeColor = vbBlack
lblfloor2.ForeColor = vbBlack

End If
End If
End Sub

Private Sub lbllogout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call strike
lbllogout.FontStrikethru = True
End Sub

Private Sub lock_Click()
adminbool = True
frmadminpass.Show 1
End Sub



Private Sub lvButtons_H1_Click()
If txt1.Text <> Clear Then
If Text2.Text = "Not Set" Then
MsgBox "Section Schedule not Set."
Else
Dim rs7 As New ADODB.Recordset
rs7.Open "SELECT * From tblsched Where room='" & Text3.Text & "'", Con
Set DataReport7.DataSource = rs7.DataSource
For Each obj In DataReport7.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rs7.DataMember
    End If
Next
DataReport7.Sections("Section2").Controls("lblyrsec").Caption = txtyr.Text & "-" & txtsec.Text
DataReport7.Sections("Section2").Controls("lblsession").Caption = Text2.Text
DataReport7.Sections("Section2").Controls("Label25").Caption = txt1.Text
DataReport7.Sections("Section2").Controls("lblroom").Caption = Text3.Text
DataReport7.Sections("Section2").Controls("lbladviser").Caption = Text5.Text
DataReport7.Sections("Section2").Controls("lbl1").Caption = txt3.Text & " " & Mid(txt4.Text, 1, 1) & ". " & txt2.Text
DataReport7.Sections("Section2").Controls("lbl2").Caption = txt5.Text
DataReport7.Sections("Section2").Controls("lbl3").Caption = txt6.Text
DataReport7.Sections("Section2").Controls("lbl4").Caption = Combo1.Text
DataReport7.Sections("Section2").Controls("lbl5").Caption = txt8.Text
DataReport7.Sections("Section2").Controls("lbl6").Caption = txt9.Text
DataReport7.Sections("Section2").Controls("lbl7").Caption = "+639" & txt11.Text
DataReport7.Sections("Section2").Controls("lbl8").Caption = Text4.Text
DataReport7.Sections("Section1").Controls("Text1").DataField = "class"
DataReport7.Sections("Section1").Controls("Text2").DataField = "time"
DataReport7.Sections("Section1").Controls("Text3").DataField = "subject"
DataReport7.Sections("Section1").Controls("Text4").DataField = "teacher"
Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from tblvalsched Where room='" & Text3.Text & "'"
        DataReport7.Sections("Section3").Controls("lblday1").Caption = RS.Fields("day")
        DataReport7.Sections("Section3").Controls("lblclass1").Caption = RS.Fields(4)
        DataReport7.Sections("Section3").Controls("lbltime1").Caption = RS.Fields(5)
        DataReport7.Sections("Section3").Controls("lblsubs1").Caption = RS.Fields("subject")
        DataReport7.Sections("Section3").Controls("lbltsub1").Caption = RS.Fields("substeacher")
        DataReport7.Sections("Section3").Controls("lblteacher1").Caption = RS.Fields("teacher")
        RS.MoveNext
        DataReport7.Sections("Section3").Controls("lblday2").Caption = RS.Fields("day")
        DataReport7.Sections("Section3").Controls("lblclass2").Caption = RS.Fields(4)
        DataReport7.Sections("Section3").Controls("lbltime2").Caption = RS.Fields(5)
        DataReport7.Sections("Section3").Controls("lblsubs2").Caption = RS.Fields("subject")
        DataReport7.Sections("Section3").Controls("lbltsub2").Caption = RS.Fields("substeacher")
      RS.Close
DataReport7.Refresh
DataReport7.Show 1
Set rs7 = Nothing
End If
End If
End Sub

Private Sub pic21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call strike
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call strike
End Sub

Private Sub Timer1_Timer()
Label24.Caption = "Time: " & Time
Label25.Caption = "Date: " & Date
End Sub
Private Sub Timer2_Timer()
If Frame3.Left > 16500 Then
Frame4.Enabled = True
Timer2.Enabled = False
login = Time
lbllogout.Enabled = True
Else
Frame3.Left = Frame3.Left + 500
swf11.Top = swf11.Top - 200
swf11.Left = swf11.Left - 200
swf12.Top = swf12.Top - 200
swf12.Left = swf12.Left + 200
swf13.Left = swf13.Left - 200
swf13.Top = swf13.Top + 200
swf14.Left = swf14.Left + 200
swf14.Top = swf14.Top + 200
End If
End Sub

Private Sub Timer3_Timer()
If Frame3.Left <= 0 Then
ShockwaveFlash1.Visible = False
Frame4.Enabled = False
Timer3.Enabled = False
frm1st.Visible = False
txtstudno.SetFocus
Else
Frame3.Left = Frame3.Left - 500
swf11.Top = swf11.Top + 200
swf11.Left = swf11.Left + 200
swf12.Top = swf12.Top + 200
swf12.Left = swf12.Left - 200
swf13.Left = swf13.Left + 200
swf13.Top = swf13.Top - 200
swf14.Left = swf14.Left - 200
swf14.Top = swf14.Top - 200
End If
End Sub

Private Sub txtstudno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
Exit Sub
End If
If Len(txtstudno.Text) = 4 Then
If KeyAscii <> 45 And Chr$(KeyAscii) <> vbBack Then KeyAscii = 0
Else
If (Not IsNumeric(Chr$(KeyAscii)) And Chr$(KeyAscii) <> vbBack) Then KeyAscii = 0
End If
End Sub
