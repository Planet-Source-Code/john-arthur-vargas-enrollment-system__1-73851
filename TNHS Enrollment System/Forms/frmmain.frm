VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tugatog National High School- Enrollment System"
   ClientHeight    =   10755
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   15270
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10755
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgbtn_unmask 
      Left            =   5520
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   150
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":164A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B879
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":10448
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":15104
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1A1CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1F966
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":254DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2B2E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":309E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":34FC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3A5EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3FA7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":451B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4A06F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4CD1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":523FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgbtn_mask 
      Left            =   3000
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   150
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":579E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5C93D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":61491
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":65F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6AB73
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6FA19
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":74EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7A61B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":800B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":85070
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8955D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8E7A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":93967
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":98DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9DB63
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":A0711
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":A5ABD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   15
      Left            =   1920
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   7005
      Left            =   2760
      ScaleHeight     =   6975
      ScaleWidth      =   9570
      TabIndex        =   22
      Top             =   3120
      Width           =   9600
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf1 
         Height          =   6975
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   9615
         _cx             =   16960
         _cy             =   12303
         FlashVars       =   ""
         Movie           =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\enrollment-header.swf"
         Src             =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\enrollment-header.swf"
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
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   3
         Left            =   60
         ScaleHeight     =   690
         ScaleWidth      =   2535
         TabIndex        =   26
         Top             =   9360
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   3
            Left            =   315
            Picture         =   "frmmain.frx":AAD71
            ScaleHeight     =   315
            ScaleWidth      =   1905
            TabIndex        =   27
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
               TabIndex        =   28
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
         TabIndex        =   23
         Top             =   8880
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   2
            Left            =   435
            Picture         =   "frmmain.frx":112C99
            ScaleHeight     =   315
            ScaleWidth      =   1905
            TabIndex        =   24
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
               TabIndex        =   25
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
      Begin VB.Image Image7 
         Height          =   500
         Left            =   0
         Picture         =   "frmmain.frx":17ABC1
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9795
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   1
         Left            =   -15
         Picture         =   "frmmain.frx":17DE96
         Stretch         =   -1  'True
         Top             =   11175
         Width           =   2730
      End
   End
   Begin VB.PictureBox pic21 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   7000
      Left            =   12360
      ScaleHeight     =   6975
      ScaleWidth      =   2970
      TabIndex        =   15
      Top             =   3120
      Width           =   3000
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   840
         Index           =   7
         Left            =   60
         ScaleHeight     =   810
         ScaleWidth      =   2535
         TabIndex        =   19
         Top             =   8880
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   5
            Left            =   435
            Picture         =   "frmmain.frx":180FCD
            ScaleHeight     =   315
            ScaleWidth      =   1905
            TabIndex        =   20
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
               TabIndex        =   21
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
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   4
         Left            =   60
         ScaleHeight     =   690
         ScaleWidth      =   2535
         TabIndex        =   16
         Top             =   9360
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   4
            Left            =   315
            Picture         =   "frmmain.frx":1E8EF5
            ScaleHeight     =   315
            ScaleWidth      =   1905
            TabIndex        =   17
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
               TabIndex        =   18
               Top             =   0
               Visible         =   0   'False
               Width           =   285
            End
         End
      End
      Begin VB.PictureBox pic24 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   6795
         Left            =   0
         ScaleHeight     =   6795
         ScaleWidth      =   3090
         TabIndex        =   74
         Top             =   5400
         Width           =   3090
         Begin VB.PictureBox PicMnu 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   840
            Index           =   18
            Left            =   60
            ScaleHeight     =   810
            ScaleWidth      =   2535
            TabIndex        =   80
            Top             =   8880
            Width           =   2565
            Begin VB.PictureBox PicBar 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   345
               Index           =   18
               Left            =   435
               Picture         =   "frmmain.frx":250E1D
               ScaleHeight     =   315
               ScaleWidth      =   1905
               TabIndex        =   81
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
                  Index           =   18
                  Left            =   240
                  TabIndex        =   82
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1440
               End
            End
            Begin VB.Image Image18 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   345
               Left            =   120
               Stretch         =   -1  'True
               Top             =   240
               Width           =   330
            End
         End
         Begin VB.PictureBox PicMnu 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   19
            Left            =   60
            ScaleHeight     =   690
            ScaleWidth      =   2535
            TabIndex        =   77
            Top             =   9360
            Width           =   2565
            Begin VB.PictureBox PicBar 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   345
               Index           =   19
               Left            =   315
               Picture         =   "frmmain.frx":2B8D45
               ScaleHeight     =   315
               ScaleWidth      =   1905
               TabIndex        =   78
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
                  Index           =   19
                  Left            =   840
                  TabIndex        =   79
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   285
               End
            End
         End
         Begin VB.PictureBox piclock 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   360
            Picture         =   "frmmain.frx":320C6D
            ScaleHeight     =   450
            ScaleWidth      =   2250
            TabIndex        =   76
            Top             =   240
            Width           =   2280
         End
         Begin VB.PictureBox piclogout 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   360
            Picture         =   "frmmain.frx":325B15
            ScaleHeight     =   450
            ScaleWidth      =   2250
            TabIndex        =   75
            Top             =   960
            Width           =   2280
         End
         Begin VB.Line Line2 
            BorderWidth     =   5
            X1              =   0
            X2              =   3500
            Y1              =   45
            Y2              =   45
         End
         Begin VB.Image Image1 
            Height          =   360
            Index           =   9
            Left            =   -15
            Picture         =   "frmmain.frx":3287B5
            Stretch         =   -1  'True
            Top             =   11175
            Width           =   2730
         End
         Begin VB.Image Image19 
            Height          =   105
            Left            =   -120
            Picture         =   "frmmain.frx":32B8EC
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4035
         End
      End
      Begin VB.Line Line1 
         BorderStyle     =   5  'Dash-Dot-Dot
         BorderWidth     =   5
         X1              =   0
         X2              =   3500
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
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
         TabIndex        =   100
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "usertype"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   99
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Image Image17 
         Height          =   105
         Left            =   0
         Picture         =   "frmmain.frx":32EBC1
         Stretch         =   -1  'True
         Top             =   1630
         Width           =   3435
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Edit Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   72
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Height          =   255
         Left            =   0
         TabIndex        =   70
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No:"
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
         TabIndex        =   69
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   68
         Top             =   4320
         Width           =   2655
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "usertype"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   5040
         Width           =   2295
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SY:"
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
         Height          =   255
         Left            =   0
         TabIndex        =   66
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "usertype"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "dasd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   63
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User Type:"
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
         Left            =   0
         TabIndex        =   62
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name:"
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
         Height          =   255
         Left            =   0
         TabIndex        =   61
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
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
         Height          =   255
         Left            =   0
         TabIndex        =   60
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID:"
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
         Height          =   255
         Left            =   0
         TabIndex        =   59
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
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
         Height          =   255
         Left            =   0
         TabIndex        =   58
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
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
         Height          =   255
         Left            =   0
         TabIndex        =   57
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lbl21 
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
         TabIndex        =   54
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   2
         Left            =   -15
         Picture         =   "frmmain.frx":331E96
         Stretch         =   -1  'True
         Top             =   11175
         Width           =   2730
      End
      Begin VB.Image Image9 
         Height          =   495
         Left            =   0
         Picture         =   "frmmain.frx":334FCD
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3315
      End
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   7005
      Left            =   0
      ScaleHeight     =   6975
      ScaleWidth      =   2730
      TabIndex        =   0
      Top             =   3120
      Width           =   2760
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
         Height          =   5655
         Left            =   2760
         TabIndex        =   55
         Top             =   0
         Width           =   30
         _cx             =   53
         _cy             =   9975
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
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
      Begin VB.PictureBox pic2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   6800
         Left            =   0
         ScaleHeight     =   6795
         ScaleWidth      =   2730
         TabIndex        =   7
         Top             =   4000
         Width           =   2730
         Begin VB.PictureBox pic3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   6800
            Left            =   0
            ScaleHeight     =   6795
            ScaleWidth      =   2730
            TabIndex        =   29
            Top             =   500
            Width           =   2730
            Begin VB.PictureBox pic4 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               FillStyle       =   0  'Solid
               ForeColor       =   &H80000008&
               Height          =   6800
               Left            =   0
               ScaleHeight     =   6795
               ScaleWidth      =   2730
               TabIndex        =   38
               Top             =   500
               Width           =   2730
               Begin VB.PictureBox pic5 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  FillStyle       =   0  'Solid
                  ForeColor       =   &H80000008&
                  Height          =   6800
                  Left            =   0
                  ScaleHeight     =   6795
                  ScaleWidth      =   2730
                  TabIndex        =   46
                  Top             =   500
                  Width           =   2730
                  Begin VB.PictureBox pic23 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BorderStyle     =   0  'None
                     FillStyle       =   0  'Solid
                     ForeColor       =   &H80000008&
                     Height          =   6800
                     Left            =   0
                     ScaleHeight     =   6795
                     ScaleWidth      =   3090
                     TabIndex        =   83
                     Top             =   500
                     Width           =   3090
                     Begin VB.PictureBox pic22 
                        Appearance      =   0  'Flat
                        BackColor       =   &H00FFFFFF&
                        BorderStyle     =   0  'None
                        FillStyle       =   0  'Solid
                        ForeColor       =   &H80000008&
                        Height          =   6675
                        Left            =   0
                        ScaleHeight     =   6675
                        ScaleWidth      =   2970
                        TabIndex        =   91
                        Top             =   500
                        Width           =   2970
                        Begin VB.PictureBox picuserlog 
                           Appearance      =   0  'Flat
                           BackColor       =   &H80000005&
                           ForeColor       =   &H80000008&
                           Height          =   480
                           Left            =   240
                           Picture         =   "frmmain.frx":3382A2
                           ScaleHeight     =   450
                           ScaleWidth      =   2250
                           TabIndex        =   105
                           Top             =   1440
                           Width           =   2280
                        End
                        Begin VB.PictureBox picmanageuser 
                           Appearance      =   0  'Flat
                           BackColor       =   &H80000005&
                           ForeColor       =   &H80000008&
                           Height          =   480
                           Left            =   240
                           Picture         =   "frmmain.frx":33D0BB
                           ScaleHeight     =   450
                           ScaleWidth      =   2250
                           TabIndex        =   104
                           Top             =   720
                           Width           =   2280
                        End
                        Begin VB.PictureBox PicMnu 
                           Appearance      =   0  'Flat
                           BackColor       =   &H00E0E0E0&
                           ForeColor       =   &H80000008&
                           Height          =   720
                           Index           =   17
                           Left            =   60
                           ScaleHeight     =   690
                           ScaleWidth      =   2535
                           TabIndex        =   95
                           Top             =   9360
                           Width           =   2565
                           Begin VB.PictureBox PicBar 
                              Appearance      =   0  'Flat
                              BackColor       =   &H80000005&
                              ForeColor       =   &H80000008&
                              Height          =   345
                              Index           =   17
                              Left            =   315
                              Picture         =   "frmmain.frx":3424B1
                              ScaleHeight     =   315
                              ScaleWidth      =   1905
                              TabIndex        =   96
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
                                 Index           =   17
                                 Left            =   840
                                 TabIndex        =   97
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
                           Index           =   16
                           Left            =   60
                           ScaleHeight     =   810
                           ScaleWidth      =   2535
                           TabIndex        =   92
                           Top             =   8880
                           Width           =   2565
                           Begin VB.PictureBox PicBar 
                              Appearance      =   0  'Flat
                              BackColor       =   &H80000005&
                              ForeColor       =   &H80000008&
                              Height          =   345
                              Index           =   16
                              Left            =   435
                              Picture         =   "frmmain.frx":3AA3D9
                              ScaleHeight     =   315
                              ScaleWidth      =   1905
                              TabIndex        =   93
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
                                 Index           =   16
                                 Left            =   240
                                 TabIndex        =   94
                                 Top             =   0
                                 Visible         =   0   'False
                                 Width           =   1440
                              End
                           End
                           Begin VB.Image Image16 
                              Appearance      =   0  'Flat
                              BorderStyle     =   1  'Fixed Single
                              Height          =   345
                              Left            =   120
                              Stretch         =   -1  'True
                              Top             =   240
                              Width           =   330
                           End
                        End
                        Begin VB.Image Image1 
                           Height          =   360
                           Index           =   8
                           Left            =   -15
                           Picture         =   "frmmain.frx":412301
                           Stretch         =   -1  'True
                           Top             =   11175
                           Width           =   2730
                        End
                        Begin VB.Label lbl22 
                           Appearance      =   0  'Flat
                           BackColor       =   &H80000005&
                           BackStyle       =   0  'Transparent
                           Caption         =   "User"
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
                           TabIndex        =   98
                           Top             =   120
                           Width           =   1095
                        End
                        Begin VB.Image img1 
                           Height          =   495
                           Index           =   6
                           Left            =   -2040
                           Picture         =   "frmmain.frx":415438
                           Stretch         =   -1  'True
                           Top             =   0
                           Width           =   4875
                        End
                     End
                     Begin VB.PictureBox PicMnu 
                        Appearance      =   0  'Flat
                        BackColor       =   &H00E0E0E0&
                        ForeColor       =   &H80000008&
                        Height          =   840
                        Index           =   15
                        Left            =   60
                        ScaleHeight     =   810
                        ScaleWidth      =   2535
                        TabIndex        =   87
                        Top             =   8880
                        Width           =   2565
                        Begin VB.PictureBox PicBar 
                           Appearance      =   0  'Flat
                           BackColor       =   &H80000005&
                           ForeColor       =   &H80000008&
                           Height          =   345
                           Index           =   15
                           Left            =   435
                           Picture         =   "frmmain.frx":41870D
                           ScaleHeight     =   315
                           ScaleWidth      =   1905
                           TabIndex        =   88
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
                              Index           =   15
                              Left            =   240
                              TabIndex        =   89
                              Top             =   0
                              Visible         =   0   'False
                              Width           =   1440
                           End
                        End
                        Begin VB.Image Image12 
                           Appearance      =   0  'Flat
                           BorderStyle     =   1  'Fixed Single
                           Height          =   345
                           Left            =   120
                           Stretch         =   -1  'True
                           Top             =   240
                           Width           =   330
                        End
                     End
                     Begin VB.PictureBox PicMnu 
                        Appearance      =   0  'Flat
                        BackColor       =   &H00E0E0E0&
                        ForeColor       =   &H80000008&
                        Height          =   720
                        Index           =   14
                        Left            =   60
                        ScaleHeight     =   690
                        ScaleWidth      =   2535
                        TabIndex        =   84
                        Top             =   9360
                        Width           =   2565
                        Begin VB.PictureBox PicBar 
                           Appearance      =   0  'Flat
                           BackColor       =   &H80000005&
                           ForeColor       =   &H80000008&
                           Height          =   345
                           Index           =   14
                           Left            =   315
                           Picture         =   "frmmain.frx":480635
                           ScaleHeight     =   315
                           ScaleWidth      =   1905
                           TabIndex        =   85
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
                              Index           =   14
                              Left            =   840
                              TabIndex        =   86
                              Top             =   0
                              Visible         =   0   'False
                              Width           =   285
                           End
                        End
                     End
                     Begin VB.PictureBox picabout 
                        Appearance      =   0  'Flat
                        BackColor       =   &H80000005&
                        ForeColor       =   &H80000008&
                        Height          =   480
                        Left            =   240
                        Picture         =   "frmmain.frx":4E855D
                        ScaleHeight     =   450
                        ScaleWidth      =   2250
                        TabIndex        =   106
                        Top             =   840
                        Width           =   2280
                     End
                     Begin VB.PictureBox picsy 
                        Appearance      =   0  'Flat
                        BackColor       =   &H80000005&
                        ForeColor       =   &H80000008&
                        Height          =   480
                        Left            =   240
                        Picture         =   "frmmain.frx":4EDB75
                        ScaleHeight     =   450
                        ScaleWidth      =   2250
                        TabIndex        =   116
                        Top             =   1680
                        Width           =   2280
                     End
                     Begin VB.Label lbl23 
                        Appearance      =   0  'Flat
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "About"
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
                        TabIndex        =   90
                        Top             =   120
                        Width           =   1695
                     End
                     Begin VB.Image Image1 
                        Height          =   360
                        Index           =   6
                        Left            =   -15
                        Picture         =   "frmmain.frx":4F3153
                        Stretch         =   -1  'True
                        Top             =   11175
                        Width           =   2730
                     End
                     Begin VB.Image img1 
                        Height          =   495
                        Index           =   5
                        Left            =   -240
                        Picture         =   "frmmain.frx":4F628A
                        Stretch         =   -1  'True
                        Top             =   0
                        Width           =   3075
                     End
                  End
                  Begin VB.PictureBox PicMnu 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00E0E0E0&
                     ForeColor       =   &H80000008&
                     Height          =   720
                     Index           =   13
                     Left            =   60
                     ScaleHeight     =   690
                     ScaleWidth      =   2535
                     TabIndex        =   50
                     Top             =   9360
                     Width           =   2565
                     Begin VB.PictureBox PicBar 
                        Appearance      =   0  'Flat
                        BackColor       =   &H80000005&
                        ForeColor       =   &H80000008&
                        Height          =   345
                        Index           =   13
                        Left            =   315
                        Picture         =   "frmmain.frx":4F955F
                        ScaleHeight     =   315
                        ScaleWidth      =   1905
                        TabIndex        =   51
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
                           Index           =   13
                           Left            =   840
                           TabIndex        =   52
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
                     Index           =   12
                     Left            =   60
                     ScaleHeight     =   810
                     ScaleWidth      =   2535
                     TabIndex        =   47
                     Top             =   8880
                     Width           =   2565
                     Begin VB.PictureBox PicBar 
                        Appearance      =   0  'Flat
                        BackColor       =   &H80000005&
                        ForeColor       =   &H80000008&
                        Height          =   345
                        Index           =   12
                        Left            =   435
                        Picture         =   "frmmain.frx":561487
                        ScaleHeight     =   315
                        ScaleWidth      =   1905
                        TabIndex        =   48
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
                           Index           =   12
                           Left            =   240
                           TabIndex        =   49
                           Top             =   0
                           Visible         =   0   'False
                           Width           =   1440
                        End
                     End
                     Begin VB.Image Image15 
                        Appearance      =   0  'Flat
                        BorderStyle     =   1  'Fixed Single
                        Height          =   345
                        Left            =   120
                        Stretch         =   -1  'True
                        Top             =   240
                        Width           =   330
                     End
                  End
                  Begin VB.PictureBox Picture3 
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     ForeColor       =   &H80000008&
                     Height          =   480
                     Left            =   240
                     Picture         =   "frmmain.frx":5C93AF
                     ScaleHeight     =   450
                     ScaleWidth      =   2250
                     TabIndex        =   108
                     Top             =   1560
                     Width           =   2280
                  End
                  Begin VB.PictureBox picrptstudents 
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     ForeColor       =   &H80000008&
                     Height          =   480
                     Left            =   240
                     Picture         =   "frmmain.frx":5CEADA
                     ScaleHeight     =   450
                     ScaleWidth      =   2250
                     TabIndex        =   107
                     Top             =   840
                     Width           =   2280
                  End
                  Begin VB.PictureBox Picture6 
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     ForeColor       =   &H80000008&
                     Height          =   480
                     Left            =   240
                     Picture         =   "frmmain.frx":5D3F5C
                     ScaleHeight     =   450
                     ScaleWidth      =   2250
                     TabIndex        =   109
                     Top             =   2280
                     Width           =   2280
                  End
                  Begin VB.Image Image1 
                     Height          =   360
                     Index           =   5
                     Left            =   -15
                     Picture         =   "frmmain.frx":5D9627
                     Stretch         =   -1  'True
                     Top             =   11175
                     Width           =   2730
                  End
                  Begin VB.Label lbl5 
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Reports"
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
                     TabIndex        =   53
                     Top             =   120
                     Width           =   1335
                  End
                  Begin VB.Image img1 
                     Height          =   495
                     Index           =   4
                     Left            =   0
                     Picture         =   "frmmain.frx":5DC75E
                     Stretch         =   -1  'True
                     Top             =   0
                     Width           =   2775
                  End
               End
               Begin VB.PictureBox PicMnu 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H80000008&
                  Height          =   840
                  Index           =   11
                  Left            =   60
                  ScaleHeight     =   810
                  ScaleWidth      =   2535
                  TabIndex        =   42
                  Top             =   8880
                  Width           =   2565
                  Begin VB.PictureBox PicBar 
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     ForeColor       =   &H80000008&
                     Height          =   345
                     Index           =   11
                     Left            =   435
                     Picture         =   "frmmain.frx":5DFA33
                     ScaleHeight     =   315
                     ScaleWidth      =   1905
                     TabIndex        =   43
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
                        Index           =   11
                        Left            =   240
                        TabIndex        =   44
                        Top             =   0
                        Visible         =   0   'False
                        Width           =   1440
                     End
                  End
                  Begin VB.Image Image13 
                     Appearance      =   0  'Flat
                     BorderStyle     =   1  'Fixed Single
                     Height          =   345
                     Left            =   120
                     Stretch         =   -1  'True
                     Top             =   240
                     Width           =   330
                  End
               End
               Begin VB.PictureBox PicMnu 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H80000008&
                  Height          =   720
                  Index           =   10
                  Left            =   60
                  ScaleHeight     =   690
                  ScaleWidth      =   2535
                  TabIndex        =   39
                  Top             =   9360
                  Width           =   2565
                  Begin VB.PictureBox PicBar 
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     ForeColor       =   &H80000008&
                     Height          =   345
                     Index           =   10
                     Left            =   315
                     Picture         =   "frmmain.frx":64795B
                     ScaleHeight     =   315
                     ScaleWidth      =   1905
                     TabIndex        =   40
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
                        Index           =   9
                        Left            =   840
                        TabIndex        =   41
                        Top             =   0
                        Visible         =   0   'False
                        Width           =   285
                     End
                  End
               End
               Begin VB.PictureBox picsms 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   480
                  Left            =   240
                  Picture         =   "frmmain.frx":6AF883
                  ScaleHeight     =   450
                  ScaleWidth      =   2250
                  TabIndex        =   110
                  Top             =   840
                  Width           =   2280
               End
               Begin VB.Label lbl4 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "SMS"
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
                  TabIndex        =   45
                  Top             =   120
                  Width           =   1095
               End
               Begin VB.Image Image1 
                  Height          =   360
                  Index           =   4
                  Left            =   -15
                  Picture         =   "frmmain.frx":6B3E55
                  Stretch         =   -1  'True
                  Top             =   11175
                  Width           =   2730
               End
               Begin VB.Image img1 
                  Height          =   495
                  Index           =   3
                  Left            =   0
                  Picture         =   "frmmain.frx":6B6F8C
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   2835
               End
            End
            Begin VB.PictureBox PicMnu 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   720
               Index           =   9
               Left            =   60
               ScaleHeight     =   690
               ScaleWidth      =   2535
               TabIndex        =   33
               Top             =   9360
               Width           =   2565
               Begin VB.PictureBox PicBar 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   345
                  Index           =   8
                  Left            =   315
                  Picture         =   "frmmain.frx":6BA261
                  ScaleHeight     =   315
                  ScaleWidth      =   1905
                  TabIndex        =   34
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
                     Index           =   7
                     Left            =   840
                     TabIndex        =   35
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
               Index           =   8
               Left            =   60
               ScaleHeight     =   810
               ScaleWidth      =   2535
               TabIndex        =   30
               Top             =   8880
               Width           =   2565
               Begin VB.PictureBox PicBar 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   345
                  Index           =   6
                  Left            =   435
                  Picture         =   "frmmain.frx":722189
                  ScaleHeight     =   315
                  ScaleWidth      =   1905
                  TabIndex        =   31
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
                     Index           =   6
                     Left            =   240
                     TabIndex        =   32
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   1440
                  End
               End
               Begin VB.Image Image10 
                  Appearance      =   0  'Flat
                  BorderStyle     =   1  'Fixed Single
                  Height          =   345
                  Left            =   120
                  Stretch         =   -1  'True
                  Top             =   240
                  Width           =   330
               End
            End
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   240
               Picture         =   "frmmain.frx":78A0B1
               ScaleHeight     =   450
               ScaleWidth      =   2250
               TabIndex        =   111
               Top             =   840
               Width           =   2280
            End
            Begin VB.Image Image1 
               Height          =   360
               Index           =   3
               Left            =   -15
               Picture         =   "frmmain.frx":78FEA8
               Stretch         =   -1  'True
               Top             =   11175
               Width           =   2730
            End
            Begin VB.Label lbl3 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Students"
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
               TabIndex        =   36
               Top             =   120
               Width           =   1815
            End
            Begin VB.Image img1 
               Height          =   495
               Index           =   2
               Left            =   0
               Picture         =   "frmmain.frx":792FDF
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2835
            End
         End
         Begin VB.PictureBox PicMnu 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   840
            Index           =   1
            Left            =   60
            ScaleHeight     =   810
            ScaleWidth      =   2535
            TabIndex        =   11
            Top             =   8880
            Width           =   2565
            Begin VB.PictureBox PicBar 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   345
               Index           =   1
               Left            =   435
               Picture         =   "frmmain.frx":7962B4
               ScaleHeight     =   315
               ScaleWidth      =   1905
               TabIndex        =   12
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
                  Index           =   1
                  Left            =   240
                  TabIndex        =   13
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1440
               End
            End
            Begin VB.Image Image4 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   345
               Left            =   120
               Stretch         =   -1  'True
               Top             =   240
               Width           =   330
            End
         End
         Begin VB.PictureBox PicMnu 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   0
            Left            =   60
            ScaleHeight     =   690
            ScaleWidth      =   2535
            TabIndex        =   8
            Top             =   9360
            Width           =   2565
            Begin VB.PictureBox PicBar 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   345
               Index           =   0
               Left            =   315
               Picture         =   "frmmain.frx":7FE1DC
               ScaleHeight     =   315
               ScaleWidth      =   1905
               TabIndex        =   9
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
                  Index           =   0
                  Left            =   840
                  TabIndex        =   10
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   285
               End
            End
         End
         Begin VB.PictureBox picsubsked 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   240
            Picture         =   "frmmain.frx":866104
            ScaleHeight     =   450
            ScaleWidth      =   2250
            TabIndex        =   112
            Top             =   960
            Width           =   2280
         End
         Begin VB.PictureBox picroommap 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   240
            Picture         =   "frmmain.frx":86B88D
            ScaleHeight     =   450
            ScaleWidth      =   2250
            TabIndex        =   113
            Top             =   1920
            Width           =   2280
         End
         Begin VB.Label lbl2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule"
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
            Left            =   0
            TabIndex        =   14
            Top             =   120
            Width           =   1455
         End
         Begin VB.Image Image1 
            Height          =   360
            Index           =   0
            Left            =   -15
            Picture         =   "frmmain.frx":8713F2
            Stretch         =   -1  'True
            Top             =   11175
            Width           =   2730
         End
         Begin VB.Image img1 
            Height          =   495
            Index           =   0
            Left            =   -120
            Picture         =   "frmmain.frx":874529
            Stretch         =   -1  'True
            Top             =   0
            Width           =   3195
         End
      End
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   5
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
            Index           =   7
            Left            =   315
            Picture         =   "frmmain.frx":8777FE
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
               Index           =   8
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
         Index           =   6
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
            Index           =   9
            Left            =   435
            Picture         =   "frmmain.frx":8DF726
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
               Index           =   10
               Left            =   240
               TabIndex        =   3
               Top             =   0
               Visible         =   0   'False
               Width           =   1440
            End
         End
         Begin VB.Image Image11 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   240
         Picture         =   "frmmain.frx":94764E
         ScaleHeight     =   450
         ScaleWidth      =   2250
         TabIndex        =   101
         Top             =   2160
         Width           =   2280
      End
      Begin VB.PictureBox picregular 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   240
         Picture         =   "frmmain.frx":94C20D
         ScaleHeight     =   450
         ScaleWidth      =   2250
         TabIndex        =   102
         Top             =   2760
         Width           =   2280
      End
      Begin VB.PictureBox pictransferee 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   240
         Picture         =   "frmmain.frx":950EB9
         ScaleHeight     =   450
         ScaleWidth      =   2250
         TabIndex        =   103
         Top             =   3360
         Width           =   2280
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   240
         Picture         =   "frmmain.frx":955F72
         ScaleHeight     =   450
         ScaleWidth      =   2250
         TabIndex        =   114
         Top             =   720
         Width           =   2280
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Enrollment"
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
         Left            =   0
         TabIndex        =   37
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image img1 
         Height          =   495
         Index           =   1
         Left            =   0
         Picture         =   "frmmain.frx":95B660
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2835
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   7
         Left            =   -15
         Picture         =   "frmmain.frx":95E935
         Stretch         =   -1  'True
         Top             =   11175
         Width           =   2730
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Add Enrollee"
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
         Left            =   0
         TabIndex        =   115
         Top             =   1620
         Width           =   2295
      End
      Begin VB.Line Line3 
         BorderStyle     =   5  'Dash-Dot-Dot
         BorderWidth     =   5
         X1              =   0
         X2              =   3500
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Image Image14 
         Height          =   105
         Left            =   0
         Picture         =   "frmmain.frx":961A6C
         Stretch         =   -1  'True
         Top             =   1875
         Width           =   3435
      End
   End
   Begin MSComctlLib.ImageList imagelist1 
      Left            =   4320
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   51
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":964D41
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9666D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9673AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":968D41
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":96A6D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":96C065
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":96D9F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":96E6D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":96F3AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":96FC87
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":970963
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97163F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":971F23
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":972BFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9734DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9741B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":975B4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9774DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":977DBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":978695
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":978F6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":979849
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97A123
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97A6BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97A9D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97ACF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97B5CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97BEA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97C77F
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97D459
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97D773
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97DBC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97EA17
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97EE69
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97F743
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":98001D
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":98580F
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9860E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":986403
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9870DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9879B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":988291
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":988B6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":989445
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":989D1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":98A5F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":98AED3
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9997B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":99B97D
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9A6544
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9A6D33
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash2 
      Height          =   3135
      Left            =   -360
      TabIndex        =   71
      Top             =   120
      Width           =   16335
      _cx             =   28813
      _cy             =   5530
      FlashVars       =   ""
      Movie           =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\title.swf"
      Src             =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\title.swf"
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
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   15500
      TabIndex        =   73
      Top             =   10320
      Width           =   49995
   End
   Begin VB.Image Image5 
      Height          =   810
      Left            =   1440
      Picture         =   "frmmain.frx":9A7415
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   945
   End
   Begin VB.Image Image3 
      Height          =   720
      Left            =   -240
      Picture         =   "frmmain.frx":9AA8A6
      Stretch         =   -1  'True
      Top             =   10080
      Width           =   15720
   End
   Begin VB.Image Image2 
      Height          =   4500
      Left            =   2280
      Picture         =   "frmmain.frx":9AC49E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   15360
   End
   Begin VB.Menu Enrollment 
      Caption         =   "Enrollment"
      Begin VB.Menu fmen 
         Caption         =   "Freshmen"
      End
      Begin VB.Menu reg 
         Caption         =   "regular"
      End
      Begin VB.Menu transferee 
         Caption         =   "Transferee"
      End
   End
   Begin VB.Menu Schedule 
      Caption         =   "Schedule"
      Begin VB.Menu subsched 
         Caption         =   "Subject Schedule"
      End
      Begin VB.Menu teachsched 
         Caption         =   "Teacher's Schedule"
      End
   End
   Begin VB.Menu stud 
      Caption         =   "Students"
      Begin VB.Menu studinfo 
         Caption         =   "Student's Information"
      End
   End
   Begin VB.Menu abt 
      Caption         =   "About"
      Begin VB.Menu about 
         Caption         =   "About TNHS"
      End
      Begin VB.Menu sysd 
         Caption         =   "School Year"
      End
   End
   Begin VB.Menu report 
      Caption         =   "Reports"
      Begin VB.Menu liststud 
         Caption         =   "List of Students Report"
      End
      Begin VB.Menu sched 
         Caption         =   "Subject Schedule"
      End
      Begin VB.Menu rmassignment 
         Caption         =   "Room Assignment"
      End
   End
   Begin VB.Menu Users 
      Caption         =   "Users"
      Begin VB.Menu muser 
         Caption         =   "Manage Users"
      End
      Begin VB.Menu ulog 
         Caption         =   "User's Log"
      End
      Begin VB.Menu editinfo 
         Caption         =   "Edit Info"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
      Begin VB.Menu lock 
         Caption         =   "Lock this System"
      End
      Begin VB.Menu lout 
         Caption         =   "Log Out"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim barclick As Integer, prevbar As String, mouses As Integer
Dim sys As String, aw As Integer
Private Sub about_Click()
Call picabout_Click
End Sub

Private Sub editinfo_Click()
frmeditinfo.Show 1
End Sub

Private Sub fmen_Click()
Call Picture1_Click
End Sub

Private Sub Form_Load()
Call dbConnection
Call displaysy
barclick = 1
ShockwaveFlash2.Movie = App.Path & "/flash/title.swf"
swf1.Movie = App.Path & "/flash/enrollment-header.swf"
swf1.Play
mouses = 0
Call changer
End Sub
Private Sub Form_Unload(Cancel As Integer)
      usrlog ("Logged out to the system")
Load frmsecurity
frmsecurity.Show
End Sub

Private Sub img1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
End Sub

Private Sub Label7_Click()
frmeditinfo.Show 1
End Sub

Private Sub lbl1_Click()
prevbar = barclick
barclick = 1
If prevbar = barclick Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
swf1.Movie = App.Path & "/flash/enrollment-header.swf"
mouses = 1
End If
End Sub

Private Sub lbl1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
lbl1.FontUnderline = True
lbl1.ForeColor = &H8000000D
End Sub
Private Sub lbl2_Click()
prevbar = barclick
barclick = 2
If prevbar = barclick Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
swf1.Movie = App.Path & "/flash/schedule-header.swf"
mouses = 1
End If
End Sub

Private Sub lbl2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
lbl2.FontUnderline = True
lbl2.ForeColor = &H8000000D

End Sub


Private Sub lbl21_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lbl21.FontUnderline = True
lbl21.ForeColor = &H8000000D
Call changer
End Sub



Private Sub lbl22_Click()

prevbar = barclick
barclick = 7
If prevbar = barclick Then
Timer1.Enabled = False
Else
aw = 1
swf1.Movie = App.Path & "/flash/user-header.swf"
Timer1.Enabled = True
mouses = 1
End If
End Sub

Private Sub lbl22_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
lbl22.FontUnderline = True
lbl22.ForeColor = &H8000000D
End Sub

Private Sub lbl23_Click()
prevbar = barclick
barclick = 6
If prevbar = barclick Then
Timer1.Enabled = False
Else
swf1.Movie = App.Path & "/flash/reports-header.swf"
Timer1.Enabled = True
mouses = 1
End If
End Sub

Private Sub lbl23_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
lbl23.FontUnderline = True
lbl23.ForeColor = &H8000000D
End Sub

Private Sub lbl24_Click()
prevbar2 = barclick2
barclick2 = 4
If prevbar2 = barclick2 Then
Timer4.Enabled = False
Else
Timer4.Enabled = True
mouses = 1
End If
End Sub

Private Sub lbl24_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lbl24.FontUnderline = True
lbl24.ForeColor = &H8000000D
End Sub

Private Sub lbl3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
lbl3.FontUnderline = True
lbl3.ForeColor = &H8000000D
End Sub
Private Sub lbl4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
lbl4.FontUnderline = True
lbl4.ForeColor = &H8000000D
End Sub
Private Sub lbl5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
lbl5.FontUnderline = True
lbl5.ForeColor = &H8000000D
End Sub
Private Sub lbl3_Click()
prevbar = barclick
barclick = 3
If prevbar = barclick Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
mouses = 1
End If
swf1.Movie = App.Path & "/flash/students-header.swf"
End Sub

Private Sub lbl4_Click()
prevbar = barclick
barclick = 4
If prevbar = barclick Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
mouses = 1
End If
swf1.Movie = App.Path & "/flash/sms-header.swf"
End Sub

Private Sub lbl5_Click()
prevbar = barclick
barclick = 5
If prevbar = barclick Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
mouses = 1
End If
swf1.Movie = App.Path & "/flash/sms-header.swf"
End Sub

Private Sub liststud_Click()
Call picrptstudents_Click
End Sub

Private Sub lock_Click()
Call piclock_Click
End Sub

Private Sub lout_Click()
Unload Me
End Sub

Private Sub muser_Click()
Call Picmanageuser_Click
End Sub


Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
End Sub
Private Sub pic2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
End Sub
Private Sub pic21_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
End Sub
Private Sub pic22_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
End Sub
Private Sub pic23_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
End Sub
Private Sub pic24_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
End Sub
Private Sub pic3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
End Sub
Private Sub pic4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
End Sub
Private Sub pic5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call changer
End Sub


Private Sub picabout_Click()
swf1.Movie = App.Path & "/flash/about.swf"
End Sub
Private Sub picabout_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
picabout.Picture = imgbtn_mask.ListImages(11).Picture
End Sub
Private Sub piclock_Click()
If MsgBox("Lock This System?", vbQuestion + vbYesNo) = vbYes Then
adminbool = True
frmadminpass.Show 1
End If
End Sub
Private Sub piclock_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
piclock.Picture = imgbtn_mask.ListImages(14).Picture
End Sub
Private Sub piclogout_Click()
If MsgBox("Are you sure you want to log-out?", vbYesNo + vbQuestion) = vbYes Then
Unload Me
End If
End Sub
Private Sub piclogout_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
piclogout.Picture = imgbtn_mask.ListImages(15).Picture
End Sub
Private Sub Picmanageuser_Click()
If userlevel = "Admin" Then
frmmanageuser.Show 1
Else
MsgBox "You don't have the rights to access this form"
End If
End Sub
Private Sub Picmanageuser_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
picmanageuser.Picture = imgbtn_mask.ListImages(1).Picture
End Sub
Private Sub picregular_Click()
If userlevel = "Admin" Then
frmregular.Show 1
Else
If valfloater = True Then
MsgBox "The system has determined that you don't have an advisory class in the previous school year level." & vbNewLine & "You cannot access this form.", vbDefaultButton1
Else
frmregular.Show 1
End If
End If
End Sub
Private Sub picregular_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
picregular.Picture = imgbtn_mask.ListImages(4).Picture
End Sub
Private Sub picroommap_Click()
If userlevel = "Teacher" Then
frmroomsched.Show 1
Else
MsgBox "You dont have the rights to access this form"
End If
End Sub
Private Sub picroommap_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
picroommap.Picture = imgbtn_mask.ListImages(7).Picture
End Sub
Private Sub picrptstudents_Click()
dreport = True
frmprint.Show 1
End Sub
Private Sub picrptstudents_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
picrptstudents.Picture = imgbtn_mask.ListImages(12).Picture
End Sub

Private Sub picsms_Click()
FormGsm.Show 1
End Sub

Private Sub picsms_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
picsms.Picture = imgbtn_mask.ListImages(10).Picture
End Sub
Private Sub picsubsked_Click()
If userlevel = "Admin" Then
frmsubsched.Show 1
Else
MsgBox "You dont have the rights to access this form"
End If
End Sub
Private Sub picsubsked_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
picsubsked.Picture = imgbtn_mask.ListImages(6).Picture
End Sub

Private Sub picsy_Click()
frmsy.Show 1
End Sub

Private Sub picsy_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
picsy.Picture = imgbtn_mask.ListImages(17).Picture
End Sub

Private Sub pictransferee_Click()
If userlevel = "Admin" Then
frmtransferee.Show 1
Else
MsgBox "You dont have the rights to access this form"
End If
End Sub
Private Sub pictransferee_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
pictransferee.Picture = imgbtn_mask.ListImages(5).Picture
End Sub
Private Sub Picture1_Click()
If userlevel = "Admin" Then
frmfreshmen.Show 1
Else
MsgBox "You dont have the rights to access this form"
End If
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Picture1.Picture = imgbtn_mask.ListImages(3).Picture
End Sub
Private Sub Picture2_Click()

If MsgBox("Activate Student Information Module?", vbYesNo + vbQuestion) = vbYes Then
frmstudentinfo.Show 1
usrlog ("Activated Student Information Module")
End If
End Sub
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Picture2.Picture = imgbtn_mask.ListImages(8).Picture
End Sub
Private Sub Picture3_Click()
dreport = False
frmprint.Show 1
End Sub
Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Picture3.Picture = imgbtn_mask.ListImages(13).Picture
End Sub

Private Sub Picture4_Click()
If userlevel = "Admin" Then
frmenrollform.Show 1
Else
MsgBox "You don't have the rights to access this form"
End If
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Picture4.Picture = imgbtn_mask.ListImages(9).Picture
End Sub

Private Sub Picture6_Click()
Dim rs9 As New ADODB.Recordset
rs9.Open "SELECT * From roommap", Con
Set DataReport8.DataSource = rs9.DataSource
For Each obj In DataReport8.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rs9.DataMember
    End If
Next
DataReport8.Sections("Section2").Controls("lblyrsec").Caption = sy & "-" & sy + 1
DataReport8.Sections("Section1").Controls("Text1").DataField = "roomname"
DataReport8.Sections("Section1").Controls("Text2").DataField = "building"
DataReport8.Sections("Section1").Controls("Text3").DataField = "floor"
DataReport8.Sections("Section1").Controls("Text4").DataField = "morning"
DataReport8.Sections("Section1").Controls("Text5").DataField = "afternoon"
DataReport8.Refresh
DataReport8.Show 1
End Sub

Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Picture6.Picture = imgbtn_mask.ListImages(16).Picture
End Sub

Private Sub picuserlog_Click()
If userlevel = "Admin" Then
frmuserlog.Show 1
Else
MsgBox "You don't have the rights to access this form"
End If
End Sub
Private Sub picuserlog_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
picuserlog.Picture = imgbtn_mask.ListImages(2).Picture
End Sub
Private Sub reg_Click()
Call picregular_Click
End Sub
Private Sub sched_Click()
Call Picture3_Click
End Sub

Private Sub subsched_Click()
Call picsubsked_Click
End Sub
Private Sub sysd_Click()
frmsy.Show 1
End Sub
Private Sub teachsched_Click()
Call picroommap_Click
End Sub
Private Sub Timer1_Timer()
Select Case barclick
Case 1:
If pic2.Top = 4000 Then
mousemovers
Else
    pic2.Top = pic2.Top + 500
    picmovers
End If
Case 2:
    If pic3.Top = 4000 Then
    mousemovers
    Else
    pic3.Top = pic3.Top + 500
    picmovers
    End If
Case 3:
    If pic4.Top = 4000 Then
    mousemovers
    Else
    pic4.Top = pic4.Top + 500
    picmovers
    End If
Case 4:
    If pic5.Top = 4000 Then
    mousemovers
    Else
    pic5.Top = pic5.Top + 500
    picmovers
    End If
Case 5:
    If pic23.Top = 4000 Then
    mousemovers
    Else
    pic23.Top = pic23.Top + 500
    picmovers
    End If
Case 6:
    If pic22.Top = 4000 Then
    mousemovers
    Else
    pic22.Top = pic22.Top + 500
    picmovers
    End If
Case 7:
   If aw = 8 Then
    mousemovers
    Else
    aw = aw + 1
    picmovers
    End If

End Select
End Sub
Sub picmovers()
        If prevbar = 1 Then
        pic2.Top = pic2.Top - 500
        ElseIf prevbar = 2 Then
        pic3.Top = pic3.Top - 500
        ElseIf prevbar = 3 Then
        pic4.Top = pic4.Top - 500
        ElseIf prevbar = 4 Then
        pic5.Top = pic5.Top - 500
        ElseIf prevbar = 5 Then
        pic23.Top = pic23.Top - 500
        ElseIf prevbar = 6 Then
        pic22.Top = pic22.Top - 500
        End If
End Sub
Sub mousemovers()
Timer1.Enabled = False
lbl1.Enabled = True
lbl2.Enabled = True
lbl4.Enabled = True
lbl3.Enabled = True
lbl5.Enabled = True
lbl23.Enabled = True
lbl22.Enabled = True

mouses = 0
aw = 1
End Sub
Sub changer()
If mouses = 0 Then
lbl1.FontUnderline = False
lbl1.ForeColor = &H0&
lbl2.FontUnderline = False
lbl2.ForeColor = &H0&
lbl3.FontUnderline = False
lbl3.ForeColor = &H0&
lbl4.FontUnderline = False
lbl4.ForeColor = &H0&
lbl5.FontUnderline = False
lbl5.ForeColor = &H0&
lbl21.FontUnderline = False
lbl21.ForeColor = &H0&
lbl22.FontUnderline = False
lbl22.ForeColor = &H0&
lbl23.FontUnderline = False
lbl23.ForeColor = &H0&
picabout.Picture = imgbtn_unmask.ListImages(11).Picture
picsms.Picture = imgbtn_unmask.ListImages(10).Picture
Picture2.Picture = imgbtn_unmask.ListImages(8).Picture
picroommap.Picture = imgbtn_unmask.ListImages(7).Picture
picsubsked.Picture = imgbtn_unmask.ListImages(6).Picture
pictransferee.Picture = imgbtn_unmask.ListImages(5).Picture
picregular.Picture = imgbtn_unmask.ListImages(4).Picture
picmanageuser.Picture = imgbtn_unmask.ListImages(1).Picture
picuserlog.Picture = imgbtn_unmask.ListImages(2).Picture
Picture1.Picture = imgbtn_unmask.ListImages(3).Picture
picrptstudents.Picture = imgbtn_unmask.ListImages(12).Picture
Picture4.Picture = imgbtn_unmask.ListImages(9).Picture
Picture3.Picture = imgbtn_unmask.ListImages(13).Picture
piclock.Picture = imgbtn_unmask.ListImages(14).Picture
piclogout.Picture = imgbtn_unmask.ListImages(15).Picture
Picture6.Picture = imgbtn_unmask.ListImages(16).Picture
picsy.Picture = imgbtn_unmask.ListImages(17).Picture
Else
lbl1.Enabled = False
lbl2.Enabled = False
lbl3.Enabled = False
lbl4.Enabled = False
lbl5.Enabled = False
lbl23.Enabled = False
lbl22.Enabled = False
End If
End Sub
Private Sub Timer2_Timer()
Label2.Caption = "Date: " & Date
Label3.Caption = "Time: " & Time
Label4.Caption = "User ID: " & userid
Label13.Caption = " " & username
Label14.Caption = names
Label15.Caption = userlevel
Label10.Caption = number
Label11.Caption = address
Label18.Caption = usrsubject
Label8.Caption = "SY: " & sy & "-" & Int(sy + 1)
End Sub
Private Sub transferee_Click()
Call pictransferee_Click
End Sub
Private Sub ulog_Click()
Call picuserlog_Click
End Sub
