VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmenrollform 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   13005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6250
      Left            =   3960
      TabIndex        =   48
      Top             =   3120
      Width           =   8895
      Begin MSComctlLib.ListView ListView2 
         Height          =   5655
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Visible         =   0   'False
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   9975
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilRecordIco"
         ForeColor       =   -2147483640
         BackColor       =   11793649
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Rank"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Student Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Gender"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Year"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Enrollment Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Grade"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "English"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Math"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Science"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Filipino"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5655
         Left            =   120
         TabIndex        =   50
         Top             =   480
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   9975
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilRecordIco"
         ForeColor       =   -2147483640
         BackColor       =   11793649
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student Number"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Gender"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Enrollment Type"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Year-Section"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Grade"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   375
         Left            =   120
         TabIndex        =   94
         Top             =   80
         Visible         =   0   'False
         Width           =   3615
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
            ItemData        =   "frmenrollform.frx":0000
            Left            =   600
            List            =   "frmenrollform.frx":000D
            TabIndex        =   95
            Text            =   "Select Year"
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Year:"
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
            Left            =   0
            TabIndex        =   96
            Top             =   50
            Width           =   975
         End
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00400000&
         Height          =   6225
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   8895
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8640
      Top             =   3720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   6735
      Left            =   3960
      TabIndex        =   3
      Top             =   2640
      Width           =   8895
      Begin VB.TextBox Text10 
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
         TabIndex        =   91
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text9 
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
         TabIndex        =   89
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox Text8 
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
         TabIndex        =   87
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Text7 
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
         TabIndex        =   85
         Top             =   2040
         Width           =   615
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Enabled         =   0   'False
         Height          =   3855
         Left            =   2400
         TabIndex        =   21
         Top             =   600
         Width           =   6495
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
            TabIndex        =   31
            Top             =   1440
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
            TabIndex        =   30
            Top             =   1920
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
            Left            =   1320
            TabIndex        =   29
            Top             =   3360
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
            TabIndex        =   28
            Top             =   2880
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
            Left            =   4440
            TabIndex        =   27
            Top             =   2400
            Width           =   1935
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
            TabIndex        =   26
            Top             =   0
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
            TabIndex        =   25
            Top             =   480
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
            Left            =   1320
            TabIndex        =   24
            Top             =   960
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
            Left            =   4560
            TabIndex        =   23
            Top             =   3360
            Width           =   1815
         End
         Begin VB.TextBox txtgender 
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
            TabIndex        =   22
            Top             =   2400
            Width           =   2175
         End
         Begin MSMask.MaskEdBox txt6 
            Height          =   375
            Left            =   1320
            TabIndex        =   32
            Top             =   1920
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
            Left            =   4920
            TabIndex        =   33
            Top             =   2880
            Width           =   1455
            _ExtentX        =   2566
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
         Begin VB.Image Image4 
            BorderStyle     =   1  'Fixed Single
            Height          =   2295
            Left            =   3720
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2655
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
            Left            =   3600
            TabIndex        =   46
            Top             =   2880
            Width           =   855
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
            TabIndex        =   45
            Top             =   1920
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
            TabIndex        =   44
            Top             =   1920
            Width           =   495
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
            TabIndex        =   43
            Top             =   3360
            Width           =   1095
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
            TabIndex        =   42
            Top             =   2880
            Width           =   1815
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
            Left            =   3600
            TabIndex        =   41
            Top             =   2400
            Width           =   1095
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
            TabIndex        =   40
            Top             =   0
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
            TabIndex        =   39
            Top             =   480
            Width           =   1455
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
            TabIndex        =   38
            Top             =   1440
            Width           =   1095
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
            TabIndex        =   37
            Top             =   960
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
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   4485
            TabIndex        =   36
            Top             =   2910
            Width           =   855
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
            TabIndex        =   35
            Top             =   2400
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
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3600
            TabIndex        =   34
            Top             =   3360
            Width           =   975
         End
      End
      Begin VB.ComboBox cmb12 
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
         ItemData        =   "frmenrollform.frx":0020
         Left            =   1080
         List            =   "frmenrollform.frx":0022
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Enabled         =   0   'False
         Height          =   2295
         Left            =   0
         TabIndex        =   14
         Top             =   4440
         Width           =   8895
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
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   83
            Top             =   600
            Width           =   2895
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
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   600
            Width           =   1215
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
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   600
            Width           =   1215
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
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   120
            Width           =   2895
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Values Educ. Teacher:"
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
            Left            =   4560
            TabIndex        =   84
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Social Studies"
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
            Left            =   7080
            TabIndex        =   82
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label lblmorn 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
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
            Index           =   15
            Left            =   8460
            TabIndex        =   81
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "TLE"
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
            Left            =   6240
            TabIndex        =   80
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label lblmorn 
            Alignment       =   2  'Center
            BackColor       =   &H00FF00FF&
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
            Index           =   14
            Left            =   6640
            TabIndex        =   79
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "MAPEH"
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
            Left            =   4800
            TabIndex        =   78
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblmorn 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
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
            Index           =   13
            Left            =   5640
            TabIndex        =   77
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label29 
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
            Left            =   3600
            TabIndex        =   76
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblmorn 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
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
            Index           =   12
            Left            =   4320
            TabIndex        =   75
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label26 
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
            Left            =   2280
            TabIndex        =   74
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblmorn 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
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
            Index           =   11
            Left            =   3120
            TabIndex        =   73
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label25 
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
            Left            =   1320
            TabIndex        =   72
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblmorn 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
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
            Left            =   1800
            TabIndex        =   71
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label24 
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
            TabIndex        =   70
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label shpenglish 
            Alignment       =   2  'Center
            BackColor       =   &H0000C000&
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
            Left            =   840
            TabIndex        =   69
            Top             =   1920
            Width           =   375
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
            Left            =   3855
            TabIndex        =   68
            Top             =   1440
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
            Left            =   4740
            TabIndex        =   67
            Top             =   1440
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
            Left            =   4395
            TabIndex        =   66
            Top             =   1440
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
            Left            =   3510
            TabIndex        =   65
            Top             =   1440
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
            Left            =   6525
            TabIndex        =   64
            Top             =   1440
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
            Left            =   5640
            TabIndex        =   63
            Top             =   1440
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
            Left            =   2610
            TabIndex        =   62
            Top             =   1440
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
            Left            =   1725
            TabIndex        =   61
            Top             =   1440
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
            Left            =   840
            TabIndex        =   60
            Top             =   1440
            Width           =   870
         End
         Begin VB.Label lblclass 
            BackColor       =   &H00FFFFC0&
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
            Height          =   300
            Index           =   7
            Left            =   720
            TabIndex        =   59
            Top             =   1200
            Width           =   7005
         End
         Begin VB.Line Line8 
            BorderWidth     =   2
            X1              =   6525
            X2              =   6525
            Y1              =   1440
            Y2              =   1800
         End
         Begin VB.Line Line7 
            BorderWidth     =   2
            X1              =   5640
            X2              =   5640
            Y1              =   1440
            Y2              =   1800
         End
         Begin VB.Line Line6 
            BorderWidth     =   2
            X1              =   4740
            X2              =   4740
            Y1              =   1440
            Y2              =   1800
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            X1              =   4395
            X2              =   4395
            Y1              =   1440
            Y2              =   1800
         End
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   3510
            X2              =   3510
            Y1              =   1440
            Y2              =   1800
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   1725
            X2              =   1725
            Y1              =   1440
            Y2              =   1800
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            X1              =   3855
            X2              =   3855
            Y1              =   1440
            Y2              =   1800
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule"
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
            TabIndex        =   58
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
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
            Left            =   2040
            TabIndex        =   57
            Top             =   120
            Width           =   855
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
            Left            =   480
            TabIndex        =   56
            Top             =   600
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
            Left            =   2640
            TabIndex        =   55
            Top             =   600
            Width           =   1095
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
            Left            =   4920
            TabIndex        =   54
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Section Information:"
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
            TabIndex        =   15
            Top             =   120
            Width           =   2535
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00400000&
            Height          =   2295
            Index           =   3
            Left            =   0
            Top             =   0
            Width           =   8895
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00400000&
            BorderWidth     =   4
            Height          =   660
            Left            =   600
            Top             =   1200
            Width           =   7080
         End
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
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin TNHSES.lvButtons_H lvButtons_H6 
         Height          =   495
         Left            =   80
         TabIndex        =   16
         Top             =   3600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "Enroll"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8454016
         cGradient       =   8454016
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   12648384
      End
      Begin TNHSES.lvButtons_H lvButtons_H7 
         Height          =   495
         Left            =   1200
         TabIndex        =   17
         Top             =   3600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8454016
         cGradient       =   8454016
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   12648384
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
         TabIndex        =   47
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Limit"
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
         Left            =   960
         TabIndex        =   92
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Girls"
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
         Left            =   480
         TabIndex        =   90
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Boys"
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
         Left            =   480
         TabIndex        =   88
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Students"
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
         TabIndex        =   86
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label20 
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
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label17 
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
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00400000&
         Height          =   4455
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   8895
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
         TabIndex        =   6
         Top             =   720
         Width           =   1095
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
         TabIndex        =   4
         Top             =   120
         Width           =   3615
      End
      Begin VB.Image Image5 
         Height          =   495
         Left            =   0
         Picture         =   "frmenrollform.frx":0024
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   3735
      Begin MSComctlLib.ImageList imagelist1 
         Left            =   240
         Top             =   120
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
               Picture         =   "frmenrollform.frx":0554
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":1EE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":2BC2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":4554
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":5EE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":7878
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":920A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":9EE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":ABBE
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":B49A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":C176
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":CE52
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":D736
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":E412
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":ECEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":F9CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":1135E
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":12CF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":135CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":13EA8
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":14782
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":1505C
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":15936
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":15ED0
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":161EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":16504
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":16DDE
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":176B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":17F92
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":18C6C
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":18F86
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":193D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":1A22A
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":1A67C
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":1AF56
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":1B830
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":21022
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":218FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":21C16
               Key             =   ""
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":228F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":231CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":23AA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":2437E
               Key             =   ""
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":24C58
               Key             =   ""
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":25532
               Key             =   ""
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":25E0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":266E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":34FCC
               Key             =   ""
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":37190
               Key             =   ""
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":41D57
               Key             =   ""
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmenrollform.frx":42546
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin TNHSES.lvButtons_H lvButtons_H1 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Caption         =   "Enroll Freshmen"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8454016
         cGradient       =   8454016
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   12648384
      End
      Begin TNHSES.lvButtons_H lvButtons_H2 
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Caption         =   "Enroll Transferee"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8454016
         cGradient       =   8454016
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   12648384
      End
      Begin TNHSES.lvButtons_H lvButtons_H3 
         Height          =   615
         Left            =   1920
         TabIndex        =   9
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Caption         =   "Enroll Regular"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8454016
         cGradient       =   8454016
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   12648384
      End
      Begin TNHSES.lvButtons_H lvButtons_H4 
         Height          =   615
         Left            =   1920
         TabIndex        =   10
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Caption         =   "Enroll Dropped"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8454016
         cGradient       =   8454016
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   12648384
      End
      Begin TNHSES.lvButtons_H lvButtons_H5 
         Height          =   495
         Left            =   1920
         TabIndex        =   11
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Caption         =   "Close"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8454016
         cGradient       =   8454016
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   12648384
      End
      Begin MSComctlLib.ListView List 
         Height          =   3375
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Year-Section"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Total"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Girls"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Boys"
            Object.Width           =   1058
         EndProperty
      End
      Begin TNHSES.lvButtons_H lvButtons_H8 
         Height          =   495
         Left            =   120
         TabIndex        =   93
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8454016
         cGradient       =   8454016
         Gradient        =   3
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         Enabled         =   0   'False
         cBack           =   12648384
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Students Enrolled per Section"
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
         Top             =   3000
         Width           =   3735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00400000&
         Height          =   2850
         Index           =   2
         Left            =   -120
         Top             =   0
         Width           =   3855
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
         TabIndex        =   2
         Top             =   120
         Width           =   7095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   3720
         Y1              =   2865
         Y2              =   2865
      End
      Begin VB.Image Image6 
         Height          =   495
         Left            =   0
         Picture         =   "frmenrollform.frx":42C28
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7335
      End
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   12480
      Picture         =   "frmenrollform.frx":43158
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "TNHS Enrollment System- Enrollment Form"
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
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   9465
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   12990
   End
   Begin VB.Image Image1 
      Height          =   2280
      Left            =   -6720
      Picture         =   "frmenrollform.frx":46105
      Stretch         =   -1  'True
      Top             =   480
      Width           =   26280
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmenrollform.frx":6915E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13155
   End
End
Attribute VB_Name = "frmenrollform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim studrc As String, studcount As Integer, yrsect As String, comp As Integer

Private Sub cmb12_Click()
Call secinfo

End Sub


Private Sub Combo1_Click()
ListView1.ListItems.Clear
ListView2.ListItems.Clear
If viewstudrec = "Transferee" Then
If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where Enrolltype='Transferee' And evaluated='Yes' And years='" & Combo1.Text & "'", Con, 1, 3
Do Until RS.EOF
With RS
If RS.Fields("Status") = "Not Enrolled" Then
ListView2.ListItems.Add , , bnt
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , RS.Fields(0)
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields(2) & " " & Mid(.Fields(3), 1, 1) & ". " & .Fields(1)
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("gender")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("years")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("Enrolltype")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("grade")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("math")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("english")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("filipino")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("science")
End If
End With
RS.MoveNext
Loop
ElseIf viewstudrec = "Regular" Then
If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where Enrolltype='Regular' And status='Not Enrolled' And evaluated='Yes' And years='" & Combo1.Text & "' Order by yrsec", Con, 1, 3
Do Until RS.EOF
ListView1.ListItems.Add , , RS.Fields(0)
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields(2) & " " & Mid(RS.Fields(3), 1, 1) & ". " & RS.Fields(1)
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("gender")

ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("Enrolltype")
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("yrsec")
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("grade")
RS.MoveNext
Loop
RS.Close

Else

If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where Enrolltype='Regular' And Status='Dropped' And evaluated='Yes' And years='" & Combo1.Text & "' Order by yrsec", Con, 1, 3
Do Until RS.EOF
ListView1.ListItems.Add , , RS.Fields(0)
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields(2) & " " & Mid(RS.Fields(3), 1, 1) & ". " & RS.Fields(1)
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("gender")

ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("Enrolltype")
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("yrsec")
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("grade")
RS.MoveNext
Loop
RS.Close

End If

End Sub

Private Sub Form_Load()
Me.Width = 500
Call loadtotalno
End Sub

Private Sub Image3_Click()
Unload Me
End Sub


Sub loadtotalno()
List.ListItems.Clear
Dim bilang As Integer
Dim boy As Integer
Dim girl As Integer
Call displaysy
Call dbConnection
For a = 1 To firstsec
bilang = 0
boy = 0
girl = 0
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where status='Enrolled'", Con, 1, 3
Do Until RS.EOF
If RS.Fields(17) = "1st-" & a Then
bilang = bilang + 1
If RS.Fields("gender") = "Male" Then
boy = boy + 1
Else
girl = girl + 1
End If
End If
RS.MoveNext
Loop
List.ListItems.Add , , "1st-" & a
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , bilang
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , boy
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , girl
RS.Close
Next

For a = 1 To secondsec
bilang = 0
boy = 0
girl = 0
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where status='Enrolled'", Con, 1, 3
Do Until RS.EOF
If RS.Fields(17) = "2nd-" & a Then
bilang = bilang + 1
If RS.Fields("gender") = "Male" Then
boy = boy + 1
Else
girl = girl + 1
End If
End If
RS.MoveNext
Loop
List.ListItems.Add , , "2nd-" & a
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , bilang
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , boy
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , girl
RS.Close
Next
For a = 1 To thirdsec
bilang = 0
boy = 0
girl = 0
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where status='Enrolled'", Con, 1, 3
Do Until RS.EOF
If RS.Fields(17) = "3rd-" & a Then
bilang = bilang + 1
If RS.Fields("gender") = "Male" Then
boy = boy + 1
Else
girl = girl + 1
End If
End If
RS.MoveNext
Loop
List.ListItems.Add , , "3rd-" & a
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , bilang
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , boy
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , girl
RS.Close
Next
For a = 1 To fourthsec
bilang = 0
boy = 0
girl = 0
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where status='Enrolled'", Con, 1, 3
Do Until RS.EOF
If RS.Fields(17) = "4th-" & a Then
bilang = bilang + 1
If RS.Fields("gender") = "Male" Then
boy = boy + 1
Else
girl = girl + 1
End If
End If
RS.MoveNext
Loop
List.ListItems.Add , , "4th-" & a
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , bilang
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , boy
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , girl
RS.Close
Next
End Sub
Private Function loadstudinfo(studrc As String)
On Error Resume Next
Set RS = New ADODB.Recordset
With RS
If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where studno='" & studrc & "'", Con, 1, 3
If RS.EOF = False Then
txt1.Text = .Fields(0)
txt2.Text = .Fields(1)
txt3.Text = .Fields(2)
txt4.Text = .Fields(3)
txt5.Text = .Fields(4)
txt6.Text = Format(.Fields(5), "MM/DD/YYYY")
txt7.Text = Year(Now) - Year(.Fields(5))
txt8.Text = .Fields(7)
txtgender.Text = .Fields("gender")
txt9.Text = .Fields(8)
txt10.Text = .Fields(9)
Text4.Text = .Fields(18)
txt11.Text = Mid(.Fields(12), 5, 9)
Text1.Text = .Fields("years")
Image4.Picture = LoadPicture(App.Path & "/students/" & .Fields(0) & ".jpg")
Call updatesec
'MsgBox .Fields("grade")
'On Error GoTo wew
If .Fields("Enrolltype") = "Freshmen" Then
'Select Case .Fields("grade")
'Case Is > 84: cmb12.ListIndex = firstsec - firstsec
'Case Is > 82: cmb12.ListIndex = (firstsec + 1) - firstsec
'Case Is > 80: cmb12.ListIndex = (firstsec + 2) - firstsec
'Case Is > 79: cmb12.ListIndex = (firstsec + 3) - firstsec
'Case Is > 78: cmb12.ListIndex = (firstsec + 4) - firstsec
'Case Is > 77: cmb12.ListIndex = (firstsec + 5) - firstsec
'Case Is > 76: cmb12.ListIndex = (firstsec + 6) - firstsec
'Case Is > 75: cmb12.ListIndex = (firstsec + 7) - firstsec
'Case Else
'cmb12.ListIndex = firstsec - 1
'End Select
If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where Enrolltype='Freshmen'", Con, 1, 3
studcount = RS.RecordCount
comp = ((studcount - (studcount Mod firstsec)) / firstsec) + 1
Dim bc As Integer
Dim cc As Integer
bc = 0
cc = 0
For a = 1 To studcount
If a = ListView2.SelectedItem.Text Then
cmb12.ListIndex = cc
End If
bc = bc + 1
'MsgBox cc < studcount Mod firstsec
If cc < studcount Mod firstsec Then

If bc = comp Then
bc = 0
cc = cc + 1
End If
Else
If bc = comp - 1 Then
bc = 0
cc = cc + 1
End If
End If

Next a

ElseIf .Fields("Enrolltype") = "Regular" Then
cmb12.ListIndex = .Fields("section") - 1
End If
End If
End With


Call secinfo

Exit Function
wew:
cmb12.ListIndex = firstsec - 1
Call secinfo
End Function

Sub updatesec()
cmb12.Clear
If Text1.Text = "1st" Then
For a = 1 To firstsec
cmb12.AddItem a
Next a
ElseIf Text1.Text = "2nd" Then
For a = 1 To secondsec
cmb12.AddItem a
Next a
ElseIf Text1.Text = "3rd" Then
For a = 1 To thirdsec
cmb12.AddItem a
Next a
ElseIf Text1.Text = "4th" Then
For a = 1 To fourthsec
cmb12.AddItem a
Next a
End If
End Sub




Private Sub ListView1_DblClick()
Frame5.Visible = False
Frame1.Enabled = True
loadstudinfo (ListView1.SelectedItem.Text)
End Sub

Private Sub ListView2_DblClick()
Frame5.Visible = False
Frame1.Enabled = True
loadstudinfo (ListView2.SelectedItem.SubItems(1))
End Sub

Private Sub lvButtons_H1_Click()
lvButtons_H8.Enabled = True
lvButtons_H1.Enabled = False
lvButtons_H2.Enabled = False
lvButtons_H3.Enabled = False
lvButtons_H4.Enabled = False
Frame5.Enabled = True
MsgBox "Enroll Freshmen"
ListView2.ListItems.Clear
ListView2.Visible = True
ListView1.Visible = False
Dim bnt As Integer
bnt = 1
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where Enrolltype='Freshmen' And evaluated='Yes' Order by studrank DESC", Con, 1, 3
Do Until RS.EOF
With RS
If RS.Fields("Status") = "Not Enrolled" Then
ListView2.ListItems.Add , , bnt
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , RS.Fields(0)
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields(2) & " " & Mid(.Fields(3), 1, 1) & ". " & .Fields(1)
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("gender")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("years")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("Enrolltype")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("grade")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("math")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("english")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("filipino")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("science")
End If
End With
bnt = bnt + 1
RS.MoveNext
Loop
RS.Close
End Sub


Private Sub lvButtons_H2_Click()
Combo1.Text = "Select Year"
Frame6.Visible = True
viewstudrec = "Transferee"
lvButtons_H8.Enabled = True
lvButtons_H1.Enabled = False
lvButtons_H2.Enabled = False
lvButtons_H3.Enabled = False
lvButtons_H4.Enabled = False
Frame5.Enabled = True
MsgBox "Enroll Transferee"
ListView2.ListItems.Clear
ListView2.Visible = True
ListView1.Visible = False
Dim bnt As Integer
bnt = 1
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where Enrolltype='Transferee' And evaluated='Yes' Order by studrank DESC", Con, 1, 3
Do Until RS.EOF
With RS
If RS.Fields("Status") = "Not Enrolled" Then
ListView2.ListItems.Add , , bnt
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , RS.Fields(0)
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields(2) & " " & Mid(.Fields(3), 1, 1) & ". " & .Fields(1)
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("gender")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("years")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("Enrolltype")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("grade")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("math")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("english")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("filipino")
ListView2.ListItems.Item(ListView2.ListItems.Count).ListSubItems.Add , , .Fields("science")
End If
End With
bnt = bnt + 1
RS.MoveNext
Loop
RS.Close
End Sub
Private Sub lvButtons_H3_Click()
viewstudrec = "Regular"
Frame6.Visible = True
Combo1.Text = "Select Year"
lvButtons_H8.Enabled = True
MsgBox "Enroll Regular"
lvButtons_H1.Enabled = False
lvButtons_H2.Enabled = False
lvButtons_H3.Enabled = False
lvButtons_H4.Enabled = False
Frame5.Enabled = True
ListView2.Visible = False
ListView1.Visible = True
If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where Enrolltype='Regular' And status='Not Enrolled' And evaluated='Yes' Order by yrsec", Con, 1, 3
Do Until RS.EOF
ListView1.ListItems.Add , , RS.Fields(0)
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields(2) & " " & Mid(RS.Fields(3), 1, 1) & ". " & RS.Fields(1)
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("gender")

ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("Enrolltype")
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("yrsec")
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("grade")
RS.MoveNext
Loop
RS.Close
End Sub

Private Sub lvButtons_H4_Click()
Combo1.Text = "Select Year"
viewstudrec = "Dropped"
Frame6.Visible = False
If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where Enrolltype='Regular' And Status='Dropped' And evaluated='Yes' Order by yrsec", Con, 1, 3
Do Until RS.EOF
ListView1.ListItems.Add , , RS.Fields(0)
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields(2) & " " & Mid(RS.Fields(3), 1, 1) & ". " & RS.Fields(1)
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("gender")

ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("Enrolltype")
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("yrsec")
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , RS.Fields("grade")
RS.MoveNext
Loop
RS.Close
lvButtons_H8.Enabled = True
MsgBox "Enroll Dropped Students"
lvButtons_H1.Enabled = False
lvButtons_H2.Enabled = False
lvButtons_H3.Enabled = False
lvButtons_H4.Enabled = False
Frame5.Enabled = True
ListView2.Visible = False
ListView1.Visible = True
End Sub



Private Sub lvButtons_H5_Click()
Unload Me
End Sub

Private Sub lvButtons_H6_Click()
Dim str As String
If Int(Text10.Text) < Int(Text7.Text) + 1 Then
str = "The selected section has reached its total student's limit per section. Bypass enrollee?"
Else
str = "Enroll Student?"

End If
If MsgBox(str, vbYesNo + vbQuestion) = vbYes Then
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from tblstudents Where studno='" & txt1.Text & "'", Con, 1, 3
If RS.EOF = False Then
RS.Fields("years") = Text1.Text
RS.Fields("section") = cmb12.Text
RS.Fields("yrsec") = Label21.Caption
RS.Fields("status") = "Enrolled"
RS.Update
End If
RS.Close
MsgBox "Student Enrolled"
usrlog ("Enrolled student with Student No. " & Label21.Caption)
i_Clear.cLearMe Me
Frame6.Visible = False
Frame5.Enabled = False
lvButtons_H8.Enabled = False
lvButtons_H1.Enabled = True
lvButtons_H2.Enabled = True
lvButtons_H3.Enabled = True
lvButtons_H4.Enabled = True
Frame5.Visible = True
ListView1.ListItems.Clear
ListView2.ListItems.Clear
Call loadtotalno
End If
End Sub

Private Sub lvButtons_H7_Click()
Call lvButtons_H8_Click
End Sub

Private Sub lvButtons_H8_Click()
Frame5.Enabled = False
lvButtons_H1.Enabled = True
lvButtons_H2.Enabled = True
lvButtons_H3.Enabled = True
lvButtons_H4.Enabled = True
lvButtons_H8.Enabled = False
Frame5.Visible = True
Frame6.Visible = False
ListView1.ListItems.Clear
ListView2.ListItems.Clear
MsgBox "Transaction Cancelled"
End Sub


Private Sub Timer1_Timer()
If Me.Width = 13000 Then
Timer1.Enabled = False
Else
Me.Left = Me.Left - 250
Me.Width = Me.Width + 500
End If

End Sub
Sub secinfo()
Label21.Caption = Text1.Text & "-" & cmb12.Text
Select Case Text1.Text
Case "1st": Text10.Text = firstsec1
Case "2nd": Text10.Text = secondsec1
Case "3rd": Text10.Text = thirdsec1
Case "4th": Text10.Text = fourthsec1
End Select
Dim x As Integer
x = List.ListItems.Count
Do Until x = 0
If Label21.Caption = List.ListItems.Item(x).Text Then
Text7.Text = List.ListItems.Item(x).SubItems(1)
Text8.Text = List.ListItems.Item(x).SubItems(3)
Text9.Text = List.ListItems.Item(x).SubItems(2)
End If
x = x - 1
Loop

Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from tblsubjsched", Con, 1, 3
RS.Find ("section='" & Text1.Text & "-" & cmb12.Text & "'")
If RS.EOF = False Then
If RS.Fields("schedstatus") = "Not Set" Then
Text2.Text = "Not Set"
Text3.Text = "Not Set"
Text5.Text = "Not Set"
Text6.Text = "Not Set"
lblclass(7).Caption = Clear

For a = 0 To 8
lblmorn(a).BackColor = &HFFFFC0
Next a
Else
Text2.Text = RS.Fields(1)
Text3.Text = RS.Fields(3)
Text5.Text = RS.Fields(4)
If RS.Fields(1) = "Morning" Then
lblclass(7).Caption = "6:00             6:50           7:40             8:30 8:50     9:20  9:40          10:30           11:20          12:10"
Else
lblclass(7).Caption = "12:30          1:20             2:10            3:00  3:20    3:50  4:10           5:00              5:50             6:40"
End If
yrsect = Label21.Caption
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
Case "1st": Call colored(RS1.Fields("subject"), 1)
Case "2nd": Call colored(RS1.Fields("subject"), 2)
Case "3rd": Call colored(RS1.Fields("subject"), 3)
Case "4th":
Call colored(RS1.Fields("subject"), 4)
If Mid(yrsect, 1, 1) = 1 Or Mid(yrsect, 1, 1) = 2 Then
lblmorn(0).BackColor = &HFFFFC0
Call colored(RS1.Fields("subject"), 8)
Else
lblmorn(8).BackColor = &HFFFFC0
Call colored(RS1.Fields("subject"), 0)
End If

Case "5th": Call colored(RS1.Fields("subject"), 5)
Case "6th": Call colored(RS1.Fields("subject"), 6)
Case "7th": Call colored(RS1.Fields("subject"), 7)
End Select

End If
RS1.MoveNext
Loop
RS1.Close
Set RS1 = New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
    RS1.ActiveConnection = Con
    RS1.CursorLocation = 3
    RS1.LockType = 2
RS1.Open "select * from tblvalsched"
Do Until RS1.EOF
If RS1.Fields(0) = yrsect Then
Text6.Text = RS1.Fields("teacher")
End If
RS1.MoveNext
Loop
RS1.Close
End If
End If

RS.Close
End Sub
Private Sub colored(subj As String, pwesto As Integer)
Select Case subj
Case "English": lblmorn(pwesto).BackColor = &HC000&
Case "Math": lblmorn(pwesto).BackColor = &H800000
Case "Science": lblmorn(pwesto).BackColor = &HFFFF&
Case "Filipino": lblmorn(pwesto).BackColor = &H80FF&
Case "Social Studies": lblmorn(pwesto).BackColor = &HC0C000
Case "TLE": lblmorn(pwesto).BackColor = &HFF00FF
Case "MAPEH": lblmorn(pwesto).BackColor = &HFF&

End Select
End Sub

