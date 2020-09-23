VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmtransferee 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Enrollment System"
   ClientHeight    =   8130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10725
   LinkTopic       =   "Form4"
   ScaleHeight     =   8130
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Input Grades"
      Height          =   2415
      Left            =   6480
      TabIndex        =   45
      Top             =   4320
      Width           =   3375
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
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   50
         Top             =   480
         Width           =   495
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
         MaxLength       =   3
         TabIndex        =   49
         Top             =   960
         Width           =   495
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
         MaxLength       =   3
         TabIndex        =   48
         Top             =   1440
         Width           =   495
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
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   47
         Top             =   960
         Width           =   495
      End
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
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   46
         Top             =   1440
         Width           =   495
      End
      Begin TNHSES.lvButtons_H lvButtons_H2 
         Height          =   375
         Left            =   2160
         TabIndex        =   51
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
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
         cGradient       =   12648384
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   8454016
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   2430
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   3390
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade"
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
         TabIndex        =   57
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Average Grade"
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
         Index           =   0
         Left            =   480
         TabIndex        =   56
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "English"
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
         TabIndex        =   55
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Math"
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
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Science"
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
         Index           =   2
         Left            =   1800
         TabIndex        =   53
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Filipino"
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
         Index           =   3
         Left            =   1800
         TabIndex        =   52
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Image Image8 
         Height          =   375
         Left            =   -720
         Picture         =   "frmtransferee.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10755
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   4560
      Top             =   480
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5535
      Left            =   160
      TabIndex        =   26
      Top             =   2400
      Width           =   3015
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
         TabIndex        =   31
         Top             =   720
         Width           =   1215
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf1 
         Height          =   2535
         Left            =   0
         TabIndex        =   27
         Top             =   3000
         Width           =   3015
         _cx             =   5318
         _cy             =   4471
         FlashVars       =   ""
         Movie           =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\transferee.swf"
         Src             =   "C:\Documents and Settings\Vargas\My Documents\TNHS Enrollment System\flash\transferee.swf"
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
         Left            =   720
         TabIndex        =   35
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         Caption         =   "Add"
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
      Begin TNHSES.lvButtons_H Command3 
         Height          =   615
         Left            =   720
         TabIndex        =   36
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   32
         Top             =   840
         Width           =   1095
      End
      Begin VB.Image Image7 
         Height          =   495
         Left            =   2400
         Picture         =   "frmtransferee.frx":3C5E
         Top             =   600
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   3500
         Y1              =   2985
         Y2              =   2985
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Menu"
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
         TabIndex        =   29
         Top             =   120
         Width           =   7095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00400000&
         Height          =   2970
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   3015
      End
      Begin VB.Image Image6 
         Height          =   495
         Left            =   0
         Picture         =   "frmtransferee.frx":40BE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   5535
      Left            =   3240
      TabIndex        =   14
      Top             =   2400
      Width           =   7335
      Begin VB.CheckBox Check3 
         BackColor       =   &H00400000&
         Caption         =   "Good Moral Certificate"
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
         Height          =   375
         Left            =   3840
         TabIndex        =   43
         Top             =   4440
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00400000&
         Caption         =   "Previous School Card"
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
         Height          =   375
         Left            =   3840
         TabIndex        =   42
         Top             =   4080
         Width           =   2655
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00400000&
         Caption         =   "Birth Certificate"
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
         Height          =   375
         Left            =   3840
         TabIndex        =   41
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00400000&
         Caption         =   "Form 137"
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
         Height          =   375
         Left            =   3840
         TabIndex        =   40
         Top             =   5160
         Width           =   2535
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
         Left            =   4800
         TabIndex        =   10
         Top             =   3300
         Width           =   2415
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
         ItemData        =   "frmtransferee.frx":45EE
         Left            =   1560
         List            =   "frmtransferee.frx":45F8
         TabIndex        =   5
         Text            =   "Select Gender.."
         Top             =   3000
         Width           =   2175
      End
      Begin VB.ComboBox cmb11 
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
         ItemData        =   "frmtransferee.frx":460A
         Left            =   5520
         List            =   "frmtransferee.frx":4617
         TabIndex        =   11
         Text            =   "Select.."
         Top             =   3720
         Width           =   1695
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
         Left            =   1560
         TabIndex        =   2
         Top             =   1560
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
         Left            =   1560
         TabIndex        =   1
         Top             =   1080
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
         Left            =   1560
         TabIndex        =   0
         Top             =   600
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
         Left            =   1560
         TabIndex        =   8
         Top             =   4440
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
         Left            =   1560
         TabIndex        =   6
         Top             =   3480
         Width           =   2175
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
         Left            =   1560
         TabIndex        =   7
         Top             =   3960
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2520
         Width           =   495
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
         Left            =   1560
         TabIndex        =   3
         Top             =   2040
         Width           =   2175
      End
      Begin MSMask.MaskEdBox txt6 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   2520
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
         Left            =   1920
         TabIndex        =   9
         Top             =   4920
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
      Begin TNHSES.lvButtons_H Command1 
         Height          =   375
         Left            =   5640
         TabIndex        =   37
         Top             =   2920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Browse..."
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
         Height          =   375
         Left            =   3840
         TabIndex        =   39
         Top             =   2920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Take a Picture"
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
      Begin TNHSES.lvButtons_H lvButtons_H4 
         Height          =   375
         Left            =   6480
         TabIndex        =   44
         Top             =   4200
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "Grade"
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
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Previous School"
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
         Height          =   495
         Left            =   3960
         TabIndex        =   38
         Top             =   3300
         Width           =   975
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   3000
         Width           =   1095
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1480
         TabIndex        =   33
         Top             =   4950
         Width           =   1215
      End
      Begin VB.Label Label17 
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
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   5400
         TabIndex        =   30
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblSetupTitle 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Student Information"
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
         TabIndex        =   28
         Top             =   120
         Width           =   7095
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label4 
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
         TabIndex        =   24
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
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
         TabIndex        =   21
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label9 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label10 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2760
         TabIndex        =   18
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label6 
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
         TabIndex        =   16
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Year to Enroll"
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
         Left            =   3960
         TabIndex        =   15
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   2295
         Left            =   4080
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2655
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00400000&
         Height          =   5535
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   7335
      End
      Begin VB.Image Image5 
         Height          =   495
         Left            =   0
         Picture         =   "frmtransferee.frx":462A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7335
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Browse Picture"
      Filter          =   "JPEG|*.jpg;*.jpeg|BMP|*.bmp"
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      Orientation     =   2
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   8120
      Index           =   0
      Left            =   10
      Top             =   15
      Width           =   10710
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "TNHS Enrollment System- Enrollment Form (Transferee)"
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
      TabIndex        =   13
      Top             =   120
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   10200
      Picture         =   "frmtransferee.frx":4B5A
      Top             =   60
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmtransferee.frx":7B07
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10755
   End
   Begin VB.Image Image2 
      Height          =   2040
      Left            =   -5520
      Picture         =   "frmtransferee.frx":B765
      Stretch         =   -1  'True
      Top             =   480
      Width           =   20160
   End
End
Attribute VB_Name = "frmtransferee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub Command1_Click()
On Error GoTo err
cd1.ShowOpen
If cd1.FileName = "" Then
MsgBox "No File Has Been Selected"
Else
Image4.Picture = LoadPicture(cd1.FileName)
End If
Exit Sub
err:
MsgBox "Invalid Filename"
End Sub

Private Sub Command3_Click()
If Command3.Caption = "Cancel" Then
 i_Clear.cLearMe Me
  cd1.FileName = ""
Image4.Picture = frmmain.imagelist1.ListImages(51).Picture
 Frame1.Enabled = False
 txt6.Text = "__/__/____"
txt11.Text = "_________"
 MsgBox "Transaction Cancelled"
 Check2.Value = 0
Check1.Value = 0
Check3.Value = 0
Check4.Value = 0
Command3.Caption = "Close"
Command4.Caption = "Add"
Image7.Enabled = True
Else
Unload Me
End If
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Add" Then
Frame1.Enabled = True
Call USR_AutoNum
txt2.SetFocus
Command4.Caption = "Enroll"
Command3.Caption = "Cancel"
Image7.Enabled = False
ElseIf Command4.Caption = "Edit" Then
Frame1.Enabled = True
txt2.SetFocus
Command4.Caption = "Update"
Command3.Caption = "Cancel"
Image7.Enabled = False
ElseIf Command4.Caption = "Update" Then
If Len(txt1.Text) = 0 Or Len(txt2.Text) = 0 Or Len(Text1.Text) = 0 Or Len(txt3.Text) = 0 Or Len(txt4.Text) = 0 Or Len(txt5.Text) = 0 Or Len(txt6.Text) = 0 Or Len(txt7.Text) = 0 Or Len(txt8.Text) = 0 Or Len(txt9.Text) = 0 Or Len(txt10.Text) = 0 Or Len(txt11.Text) = 0 Or cmb11.Text = "Select.." Or Combo1.Text = "Select Gender.." Or Len(Text4.Text) = 0 Then
MsgBox "All Fields are Required"
Else
Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
    RS.Open "Select * From tblstudents Where studno='" & txt1.Text & "'"
    With RS
    If .EOF = False Then
     .Fields(0) = txt1.Text
    .Fields(1) = txt2.Text
    .Fields(2) = txt3.Text
    .Fields(3) = txt4.Text
    .Fields(4) = txt5.Text
    .Fields(5) = txt6.Text
    .Fields(6) = txt7.Text
    .Fields(7) = txt8.Text
    .Fields(8) = txt9.Text
    .Fields(9) = txt10.Text
        .Fields(10) = cmb11.Text
    .Fields(12) = "+639" & txt11.Text
    .Fields(13) = "Not Enrolled"
    .Fields(14) = "Transferee"
    .Fields("evaluated") = "Yes"
    .Fields(15) = sy
    .Fields(18) = Text4.Text
    .Fields(19) = Text1.Text
    .Fields("gender") = Combo1.Text
    .Fields("goodmoral") = "Yes"
    .Fields("schoolcard") = "Yes"
    If Check2.Value = 0 Then
    .Fields("bc") = "No"
    Else
    .Fields("bc") = "Yes"
    End If
    If Check4.Value = 0 Then
    .Fields("goodmoral") = "No"
    Else
    .Fields("goodmoral") = "Yes"
    End If
    .Fields("english") = Text2.Text
    .Fields("math") = Text3.Text
    .Fields("science") = Text5.Text
    .Fields("filipino") = Text6.Text

     End If
    End With
    RS.Update
    RS.Close
     usrlog ("Edited transferee Student with Student Number:" & txt1.Text)
 SavePicture Image4, App.Path & "/students/" & txt1.Text & ".jpg"
    MsgBox "Record Updated"
 cd1.FileName = ""
Image4.Picture = frmmain.imagelist1.ListImages(51).Picture
  i_Clear.cLearMe Me
 Command3.Caption = "Exit"
 Frame1.Enabled = False
txt6.Text = "__/__/____"
txt11.Text = "_________"
Command4.Caption = "Add"
 Image7.Enabled = True
 Check2.Value = 0
Check1.Value = 0
Check3.Value = 0
Check4.Value = 0
 End If
Else
If Len(txt1.Text) = 0 Or Len(txt2.Text) = 0 Or Len(Text1.Text) = 0 Or Len(txt3.Text) = 0 Or Len(txt4.Text) = 0 Or Len(txt5.Text) = 0 Or Len(txt6.Text) = 0 Or Len(txt7.Text) = 0 Or Len(txt8.Text) = 0 Or Len(txt9.Text) = 0 Or Len(txt10.Text) = 0 Or Len(txt11.Text) = 0 Or cmb11.Text = "Select.." Or Combo1.Text = "Select Gender.." Or Len(Text4.Text) = 0 Then
MsgBox "All Fields are Required"
Else
Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
    RS.Open "Select * From tblstudents"
    RS.AddNew
    With RS
     .Fields(0) = txt1.Text
    .Fields(1) = txt2.Text
    .Fields(2) = txt3.Text
    .Fields(3) = txt4.Text
    .Fields(4) = txt5.Text
    .Fields(5) = txt6.Text
    .Fields(6) = txt7.Text
    .Fields(7) = txt8.Text
    .Fields(8) = txt9.Text
    .Fields(9) = txt10.Text
        .Fields(10) = cmb11.Text
    .Fields(12) = "+639" & txt11.Text
    .Fields(13) = "Not Enrolled"
    .Fields(14) = "Transferee"
    .Fields("evaluated") = "Yes"
    .Fields(15) = sy
    .Fields(18) = Text4.Text
    .Fields(19) = Text1.Text
    .Fields("gender") = Combo1.Text
    .Fields("goodmoral") = "Yes"
    .Fields("schoolcard") = "Yes"
    If Check2.Value = 0 Then
    .Fields("bc") = "No"
    Else
    .Fields("bc") = "Yes"
    End If
    If Check4.Value = 0 Then
    .Fields("goodmoral") = "No"
    Else
    .Fields("goodmoral") = "Yes"
    End If
    .Fields("english") = Text2.Text
    .Fields("math") = Text3.Text
    .Fields("science") = Text5.Text
    .Fields("filipino") = Text6.Text

    End With
    RS.Update
    RS.Close
       usrlog ("Added transferee Student with Student Number:" & txt1.Text)
 SavePicture Image4, App.Path & "/students/" & txt1.Text & ".jpg"
MsgBox "Student Enrolled"
 cd1.FileName = ""
Image4.Picture = frmmain.imagelist1.ListImages(51).Picture
  i_Clear.cLearMe Me
txt6.Text = "__/__/____"
txt11.Text = "_________"
 Command3.Caption = "Exit"
 Frame1.Enabled = False
    Command4.Caption = "Add"
    Image7.Enabled = True
    Check2.Value = 0
Check1.Value = 0
Check3.Value = 0
Check4.Value = 0
End If
End If
End Sub

Private Sub Form_Load()
Label17.Caption = "SY: " & sy & "-" & Int(sy) + 1
Call dbConnection
swf1.Movie = App.Path & "\flash\transferee.swf"
swf1.Play
Image4.Picture = frmmain.imagelist1.ListImages(51).Picture
Me.Width = 225
Frame3.Visible = False
End Sub

Private Sub Image3_Click()
Unload Me
End Sub
Private Function USR_AutoNum()
Dim conwan As New ADODB.Connection
Dim rswan As New ADODB.Recordset
conwan.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\db1.mdb;Persist Security Info=False"
conwan.Open
rswan.Open "select * from tblstudents Where skulyir='" & sy & "' Order by studno DESC", conwan, 3, 2
If rswan.RecordCount = 0 Then
    txt1.Text = sy & "-" & "0000"
Else
    txt1.Text = sy & "-" & Format(Right(rswan!studno, 4) + 1, "0000")
End If
rswan.Close
End Function


Private Sub Image7_Click()
viewstudrec = "transferee"
frmviewstudrec.Show 1
End Sub



Private Sub lvButtons_H1_Click()
viewstudrec = "transferee"
frmwebcam.Show 1
End Sub

Private Sub lvButtons_H2_Click()
If Text1.Text = Clear Or Text2.Text = Clear Or Text3.Text = Clear Or Text5.Text = Clear Or Text6.Text = Clear Then
MsgBox "all input grades is a required field"
Else
If Text1.Text > 101 Or Text1.Text < 60 Or Text2.Text > 101 Or Text2.Text < 60 Or Text3.Text > 101 Or Text3.Text < 60 Or Text5.Text > 101 Or Text5.Text < 60 Or Text6.Text > 101 Or Text6.Text < 60 Then
MsgBox "Please input Valid grade value"
Else
Frame3.Visible = False
Frame2.Enabled = True
Frame1.Enabled = True
End If
End If
End Sub
Private Sub lvButtons_H4_Click()
Frame3.Visible = True
Frame2.Enabled = False
Frame1.Enabled = False
End Sub

Private Sub Timer1_Timer()
If Me.Width >= 10725 Then
Timer1.Enabled = False
Else
Me.Left = Me.Left - 250
Me.Width = Me.Width + 500
End If
End Sub
Private Sub txt6_LostFocus()
On Error GoTo eror
txt7.Text = Year(Now) - Year(txt6.Text)
If txt7.Text >= 70 Or txt7.Text <= 0 Then GoTo eror
Exit Sub
eror:
MsgBox "Invalid Date Value"
txt6.Text = "__/__/____"
txt7.Text = Clear
txt6.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
 If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
 If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
 If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub

