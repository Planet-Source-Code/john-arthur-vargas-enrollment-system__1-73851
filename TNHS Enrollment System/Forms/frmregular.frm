VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmregular 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Enrollment System"
   ClientHeight    =   8145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10725
   LinkTopic       =   "Form4"
   ScaleHeight     =   8145
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   3480
      Top             =   1440
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5655
      Left            =   160
      TabIndex        =   23
      Top             =   2280
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
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf1 
         Height          =   2535
         Left            =   0
         TabIndex        =   24
         Top             =   3120
         Width           =   3015
         _cx             =   5318
         _cy             =   4471
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
      Begin TNHSES.lvButtons_H cmdenroll 
         Height          =   495
         Left            =   960
         TabIndex        =   37
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "Evaluate"
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
      Begin TNHSES.lvButtons_H Command4 
         Height          =   495
         Left            =   960
         TabIndex        =   38
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "Edit"
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
         Height          =   495
         Left            =   960
         TabIndex        =   39
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Image Image7 
         Height          =   495
         Left            =   2480
         Picture         =   "frmregular.frx":0000
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label1 
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
         TabIndex        =   29
         Top             =   840
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   3360
         Y1              =   3105
         Y2              =   3105
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
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   7095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00400000&
         Height          =   5760
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   3015
      End
      Begin VB.Image Image6 
         Height          =   495
         Left            =   0
         Picture         =   "frmregular.frx":0460
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
      Height          =   5655
      Left            =   3240
      TabIndex        =   12
      Top             =   2280
      Width           =   7335
      Begin VB.CheckBox Check4 
         BackColor       =   &H00400000&
         Caption         =   "Form 137"
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
         Height          =   375
         Left            =   5760
         TabIndex        =   51
         Top             =   5160
         Width           =   2535
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00400000&
         Caption         =   "Birth Certificate"
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
         Height          =   375
         Left            =   5760
         TabIndex        =   50
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00400000&
         Caption         =   "Previous School Card"
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
         Height          =   375
         Left            =   3600
         TabIndex        =   49
         Top             =   4800
         Width           =   2655
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00400000&
         Caption         =   "Good Moral Certificate"
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
         Height          =   375
         Left            =   3600
         TabIndex        =   48
         Top             =   5160
         Width           =   2535
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
         ItemData        =   "frmregular.frx":0990
         Left            =   1320
         List            =   "frmregular.frx":099A
         TabIndex        =   5
         Text            =   "Select Gender.."
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1935
         Left            =   4080
         TabIndex        =   30
         Top             =   2880
         Width           =   3495
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
            Left            =   1800
            TabIndex        =   45
            Top             =   1080
            Width           =   1335
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
            Left            =   1200
            TabIndex        =   43
            Top             =   1560
            Width           =   1935
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
            Left            =   840
            TabIndex        =   41
            Top             =   120
            Width           =   2295
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
            Left            =   840
            TabIndex        =   33
            Top             =   600
            Width           =   735
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
            Left            =   2400
            TabIndex        =   32
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Previous SY Grade"
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
            TabIndex        =   46
            Top             =   1080
            Width           =   1935
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
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   120
            TabIndex        =   44
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
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
            TabIndex        =   42
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label13 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   1560
            TabIndex        =   34
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label11 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   0
            TabIndex        =   31
            Top             =   600
            Width           =   2055
         End
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
         Left            =   1320
         TabIndex        =   2
         Top             =   1680
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
         Left            =   1320
         TabIndex        =   1
         Top             =   1200
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
         Left            =   1320
         TabIndex        =   0
         Top             =   720
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
         Left            =   1320
         TabIndex        =   8
         Top             =   4560
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
         Left            =   1320
         TabIndex        =   6
         Top             =   3600
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
         Left            =   1320
         TabIndex        =   7
         Top             =   4080
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2640
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
         Left            =   1320
         TabIndex        =   3
         Top             =   2160
         Width           =   2175
      End
      Begin MSMask.MaskEdBox txt6 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   2640
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
         Left            =   1680
         TabIndex        =   9
         Top             =   5040
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
         Left            =   3720
         TabIndex        =   40
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
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
         Height          =   615
         Left            =   3720
         TabIndex        =   47
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
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
         Left            =   0
         TabIndex        =   36
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label22 
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
         Left            =   1200
         TabIndex        =   35
         Top             =   5040
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
         TabIndex        =   27
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
         TabIndex        =   25
         Top             =   120
         Width           =   3495
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
         Left            =   0
         TabIndex        =   22
         Top             =   1680
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
         Left            =   0
         TabIndex        =   21
         Top             =   2160
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
         Left            =   0
         TabIndex        =   20
         Top             =   1200
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
         Left            =   0
         TabIndex        =   19
         Top             =   720
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
         Left            =   0
         TabIndex        =   18
         Top             =   4560
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
         Left            =   0
         TabIndex        =   17
         Top             =   3600
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
         Left            =   0
         TabIndex        =   16
         Top             =   4080
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
         Left            =   2520
         TabIndex        =   15
         Top             =   2640
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
         Left            =   0
         TabIndex        =   14
         Top             =   2640
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
         Left            =   0
         TabIndex        =   13
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00400000&
         Height          =   5655
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   7335
      End
      Begin VB.Image Image5 
         Height          =   495
         Left            =   0
         Picture         =   "frmregular.frx":09AC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11055
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1680
      Top             =   960
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
      Height          =   8115
      Index           =   0
      Left            =   15
      Top             =   0
      Width           =   10710
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "TNHS Enrollment System- Enrollment Form (Regular)"
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
      TabIndex        =   11
      Top             =   120
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   10200
      Picture         =   "frmregular.frx":0EDC
      Top             =   60
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmregular.frx":3E89
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10755
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   -6000
      Picture         =   "frmregular.frx":7AE7
      Stretch         =   -1  'True
      Top             =   480
      Width           =   21240
   End
End
Attribute VB_Name = "frmregular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdenroll_Click()
If txt1.Text = "" Then
MsgBox "Please Select Student First"
Else
If Text1.Text = "Enrolled" Then
If MsgBox("The student you have selected is already evaluated." & vbNewLine & "Do you want to change the evaluation of the selected student?", vbYesNo + vbCritical) = vbYes Then
frmenroll.Show 1
End If
Else
frmenroll.Show 1
End If
End If
End Sub

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
Image4.Picture = frmmain.ImageList1.ListImages(51).Picture
 Frame1.Enabled = False
 Command3.Caption = "Close"
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True

 MsgBox "Transaction Cancelled"
 Command4.Enabled = False
Else
Unload Me
End If
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Update" Then
If Len(txt1.Text) = 0 Or Len(txt2.Text) = 0 Or Len(txt3.Text) = 0 Or Len(txt4.Text) = 0 Or Len(txt5.Text) = 0 Or Len(txt6.Text) = 0 Or Len(txt7.Text) = 0 Or Len(txt8.Text) = 0 Or Len(txt9.Text) = 0 Or Len(txt10.Text) = 0 Or Len(txt11.Text) = 0 Or Combo1.Text = "Select Gender.." Then
MsgBox "All Fields are Required"
Else
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblstudents Where studno='" & txt1.Text & "'"
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
    If Check2.Value = 0 Then
    .Fields("bc") = "No"
    Else
    .Fields("bc") = "Yes"
    End If
    If Check4.Value = 0 Then
    .Fields("form137") = "No"
    Else
    .Fields("form137") = "Yes"
    End If
.Fields(12) = "+639" & txt11.Text
.Fields("gender") = Combo1.Text
.Update
End If
SavePicture Image4, App.Path & "/students/" & txt1.Text & ".jpg"
MsgBox "Record Updated"
cd1.FileName = ""
Command4.Caption = "Edit Info"
Frame1.Enabled = False

If .Fields(13) = "Enrolled" Then
cmdenroll.Enabled = False
Image4.Picture = frmmain.ImageList1.ListImages(51).Picture
i_Clear.cLearMe Me
txt6.Text = "__/__/____"
txt11.Text = "_________"
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
cmdenroll.Enabled = False
Command4.Enabled = False
Command3.Caption = "Exit"
Else
cmdenroll.Enabled = True
Command4.Enabled = True
Command3.Caption = "Cancel"
End If
.Close
End With
End If
Else
If txt1.Text = "" Then
MsgBox "Please Select Student First"
Else
Frame1.Enabled = True
txt2.SetFocus
cmdenroll.Enabled = False
Command4.Caption = "Update"
If frmregular.Check1.Value = 1 Then frmregular.Check1.Enabled = False
If frmregular.Check2.Value = 1 Then frmregular.Check2.Enabled = False
If frmregular.Check3.Value = 1 Then frmregular.Check3.Enabled = False
If frmregular.Check4.Value = 1 Then frmregular.Check4.Enabled = False
End If
End If
End Sub

Private Sub Commnad4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Form_Load()
Label17.Caption = "SY: " & sy & "-" & Int(sy + 1)
Call dbConnection
swf1.Movie = App.Path & "\flash\regular.swf"
swf1.Play
Image4.Picture = frmmain.ImageList1.ListImages(51).Picture
Me.Width = 225

End Sub

Private Sub Image3_Click()
Unload Me
End Sub
Private Function USR_AutoNum()
Dim conwan As New ADODB.Connection
Dim rswan As New ADODB.Recordset
conwan.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\db1.mdb;Persist Security Info=False"
conwan.Open
rswan.Open "select * from tblstudents Where skulyir='" & Year(Now) & "'", conwan, 3, 2
If rswan.RecordCount = 0 Then
    txt1.Text = Year(Now) & "-" & "0000"
Else
rswan.MoveLast
    txt1.Text = Year(Now) & "-" & Format(Right(rswan!studno, 4) + 1, "0000")
End If
rswan.Close
End Function


Private Sub Image7_Click()
viewstudrec = "regular"
frmviewstudrec.Show 1
End Sub

Private Sub lvButtons_H1_Click()
viewstudrec = "regular"
frmwebcam.Show 1
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
