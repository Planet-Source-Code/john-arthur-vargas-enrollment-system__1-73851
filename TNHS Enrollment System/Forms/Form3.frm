VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmanageuser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9180
   ClientLeft      =   -30
   ClientTop       =   -120
   ClientWidth     =   10410
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9180
   ScaleMode       =   0  'User
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TNHSES.lvButtons_H Command1 
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      cGradient       =   12648447
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454143
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtcnumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      MaxLength       =   9
      TabIndex        =   20
      Top             =   8640
      Width           =   2175
   End
   Begin VB.TextBox txtfname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   7815
      Width           =   2655
   End
   Begin VB.ComboBox cmbsubject 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form3.frx":0000
      Left            =   7560
      List            =   "Form3.frx":0013
      TabIndex        =   3
      Text            =   "Select"
      Top             =   6945
      Width           =   2655
   End
   Begin VB.ComboBox cmbuserlevel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form3.frx":0030
      Left            =   7560
      List            =   "Form3.frx":003A
      TabIndex        =   2
      Text            =   "Select User Level"
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   8400
      Top             =   0
   End
   Begin VB.TextBox txtsname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   7410
      Width           =   2655
   End
   Begin VB.TextBox txtaddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   6
      Top             =   8220
      Width           =   2655
   End
   Begin VB.TextBox txtpass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   7560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   6075
      Width           =   2655
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox txtuser 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7560
      TabIndex        =   0
      Top             =   5680
      Width           =   2655
   End
   Begin TNHSES.lvButtons_H Command4 
      Height          =   375
      Left            =   3240
      TabIndex        =   24
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Block"
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
      cGradient       =   12648447
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454143
   End
   Begin TNHSES.lvButtons_H Command3 
      Height          =   375
      Left            =   4560
      TabIndex        =   22
      Top             =   3240
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
      cGradient       =   12648447
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454143
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   240
      TabIndex        =   18
      Top             =   3720
      Width           =   5535
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Left            =   3480
         TabIndex        =   26
         Top             =   0
         Width           =   1935
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
         ItemData        =   "Form3.frx":004E
         Left            =   1080
         List            =   "Form3.frx":0064
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   0
         Width           =   2175
      End
      Begin MSComctlLib.ListView list 
         Height          =   4575
         Left            =   0
         TabIndex        =   21
         Top             =   480
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   8070
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilRecordIco"
         ColHdrIcons     =   "icoHeader"
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User ID"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Username"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "User Level"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblListInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "0 Records"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   4440
         TabIndex        =   28
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Select By:"
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
         Left            =   0
         TabIndex        =   27
         Top             =   120
         Width           =   1095
      End
   End
   Begin TNHSES.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   9000
      TabIndex        =   29
      Top             =   4200
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
      cBhover         =   12648447
      cGradient       =   12648447
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   8454143
   End
   Begin TNHSES.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   8640
      TabIndex        =   30
      Top             =   4680
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
      cBhover         =   12648384
      LockHover       =   1
      cGradient       =   12648447
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   8454143
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   120
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Browse Picture"
      Filter          =   "JPEG|*.jpg;*.jpeg|BMP|*.bmp"
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      Orientation     =   2
   End
   Begin MSComctlLib.ImageList icoHeader 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":00A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":0641
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":0BDB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   9165
      Left            =   15
      Top             =   15
      Width           =   10435
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      Top             =   9045
      Width           =   10335
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      Top             =   2370
      Width           =   10215
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   6720
      Left            =   5880
      Top             =   2445
      Width           =   75
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00400000&
      Height          =   6480
      Left            =   120
      Top             =   2565
      Width           =   10215
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Information"
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
      Left            =   6000
      TabIndex        =   32
      Top             =   2640
      Width           =   4095
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
      TabIndex        =   31
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label12 
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
      Height          =   375
      Left            =   7560
      TabIndex        =   19
      Top             =   8640
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   9960
      Picture         =   "Form3.frx":1175
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
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
      Left            =   6120
      TabIndex        =   17
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Year to Modify"
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
      Left            =   6120
      TabIndex        =   16
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "User Level"
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
      Height          =   615
      Left            =   6120
      TabIndex        =   15
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TNHS Enrollment System- Manage Users"
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
      TabIndex        =   14
      Top             =   120
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "Form3.frx":4122
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10515
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
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   8760
      Width           =   1695
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   7440
      Width           =   1335
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
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   6120
      TabIndex        =   10
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
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
      Left            =   6120
      TabIndex        =   9
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      Height          =   615
      Left            =   6120
      TabIndex        =   8
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   120
      Picture         =   "Form3.frx":80DA
      Stretch         =   -1  'True
      Top             =   2565
      Width           =   10215
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      Height          =   6720
      Left            =   120
      Top             =   2400
      Width           =   10215
   End
   Begin VB.Image Image2 
      Height          =   2040
      Left            =   -5520
      Picture         =   "Form3.frx":860A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   20520
   End
   Begin VB.Menu teach 
      Caption         =   "teacher"
      Begin VB.Menu blk 
         Caption         =   "Block"
      End
      Begin VB.Menu info 
         Caption         =   "Teacher's Information"
      End
   End
End
Attribute VB_Name = "frmmanageuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filt As Integer, boolclk As Boolean, indchk As String, sqlstate As String

Private Sub blk_Click()
If blk.Caption = "Block User" Then
     If MsgBox("Are you sure you want to Block this User?", vbYesNo + vbQuestion) = vbYes Then
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from users Where ID='" & txtid.Text & "'"
        With RS
        .Fields("status") = "inactive"
        .Update
        .Close
        End With
          MsgBox "Record Blocked"
                usrlog ("Blocked user's record with ID '" & txtid.Text & "'")

Call loadusrlist
End If
Else
     If MsgBox("Are you sure you want to activate this User?", vbYesNo + vbQuestion) = vbYes Then
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from users Where ID='" & txtid.Text & "'"
        With RS
        .Fields("status") = "active"
        .Update
        .Close
        End With
        MsgBox "Record Activated"
        usrlog ("Activated user's record with ID '" & txtid.Text & "'")
        Call loadusrlist
        End If
End If
End Sub

Private Sub cmbuserlevel_Click()
If cmbuserlevel = "Teacher" Then
cmbsubject.Enabled = True
cmbsubject.Text = "Select Year to Modify"
Else
cmbsubject.Enabled = False
cmbsubject.Text = "n/a"
End If
End Sub



Private Sub Combo1_Click()
filt = Combo1.ListIndex
Call Text1_Change
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Add" Then
lvButtons_H1.Enabled = True
lvButtons_H2.Enabled = True
Image4.Picture = frmmain.imagelist1.ListImages(51).Picture
txtid.Enabled = True
txtuser.Enabled = True
txtpass.Enabled = True
cmbuserlevel.Enabled = True
cmbsubject.Enabled = False
txtsname.Enabled = True
txtfname.Enabled = True
txtaddress.Enabled = True
txtcnumber.Enabled = True
txtid.Text = Clear
txtuser.Text = Clear
txtpass.Text = Clear
cmbuserlevel.Text = "Select User Level"
cmbsubject.Text = "Select Subject"
txtsname.Text = Clear
txtfname.Text = Clear
txtaddress.Text = Clear
txtcnumber.Text = Clear
txtuser.SetFocus
Frame1.Enabled = False
Text1.Enabled = False
Call USR_AutoNum
Command1.Caption = "Save"
Command4.Caption = "Cancel"

Frame1.Enabled = False
ElseIf Command1.Caption = "Save" Then
If Len(txtid.Text) = 0 Or Len(txtuser.Text) = 0 Or Len(txtpass.Text) = 0 Or Len(txtsname.Text) = 0 Or Len(txtfname.Text) = 0 Or cmbuserlevel.Text = "Select User Level" Then
    MsgBox "Username, Password, User Level, Surname and Firstname are Required fields"
    Else
If cmbuserlevel.Text = "Teacher" Then
If cmbsubject = "Select Subject" Then
MsgBox "User Level Teacher must select subject to teach"
Else
Call addrecord

End If
Else
Call addrecord
End If
End If
    End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Cancel" Then
Call loadusrlist
Else
If txtid.Text = "" Then
MsgBox "No Record has been selected"
Else
txtid.Enabled = True
txtuser.Enabled = True
txtpass.Enabled = True
cmbuserlevel.Enabled = True
cmbsubject.Enabled = True
txtsname.Enabled = True
txtfname.Enabled = True
txtaddress.Enabled = True
txtcnumber.Enabled = True
Command1.Caption = "Update"
Command2.Caption = "Cancel"
Command4.Enabled = False
Frame1.Enabled = False
End If
End If
End Sub

Private Sub Command3_Click()
Timer2.Enabled = True
End Sub

Private Sub Command4_Click()
If Command4.Caption = "Cancel" Then
Call loadusrlist
ElseIf Command4.Caption = "&Block" Then
If txtid.Text = "" Then
MsgBox "No Record Has Been Selected"
Else
     If MsgBox("Are you sure you want to Block this User?", vbYesNo + vbQuestion) = vbYes Then
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from users Where ID='" & txtid.Text & "'"
        With RS
        .Fields("status") = "inactive"
        .Update
        .Close
        End With
          MsgBox "Record Blocked"
                usrlog ("Blocked user's record with ID '" & txtid.Text & "'")

Call loadusrlist
End If
End If
ElseIf Command4.Caption = "A&ctivate" Then
If txtid.Text = "" Then
MsgBox "No Record Has Been Selected"
    Else
     If MsgBox("Are you sure you want to activate this User?", vbYesNo + vbQuestion) = vbYes Then
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from users Where ID='" & txtid.Text & "'"
        With RS
        .Fields("status") = "active"
        .Update
        .Close
        End With
        MsgBox "Record Activated"
        usrlog ("Activated user's record with ID '" & txtid.Text & "'")
        Call loadusrlist
        End If
        End If
End If
End Sub

Private Sub Form_Load()
sqlstate = ""
Image4.Picture = frmmain.imagelist1.ListImages(51).Picture
frmmanageuser.Width = 435
Call dbConnection
Call loadusrlist
filt = 0
Combo1.ListIndex = 0
Me.Height = 9210
teach.Enabled = False
teach.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer2.Enabled = True
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

Private Sub info_Click()
teachersrec.txtid.Text = List.SelectedItem.Text
teachersrec.Show 1
End Sub

Private Sub list_Click()
On Error Resume Next
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
    RS.Open "Select * From users"
    RS.Find ("ID='" & List.SelectedItem & "'")
    If RS.EOF = False Then
txtid.Text = RS.Fields(0)
txtuser.Text = RS.Fields(1)
txtpass.Text = RS.Fields(2)
cmbuserlevel.Text = RS.Fields(5)
cmbsubject.Text = RS.Fields(6)
txtsname.Text = RS.Fields(3)
txtfname.Text = RS.Fields(4)
txtaddress.Text = RS.Fields(7)
txtcnumber.Text = Mid(RS.Fields(8), 5, 9)
Image4.Picture = LoadPicture(App.Path & "/usrpic/" & txtid.Text & ".jpg")
If RS.Fields("status") = "active" Then
Command4.Caption = "&Block"
Else
Command4.Caption = "A&ctivate"
End If

    End If
End Sub

Private Sub lvButtons_H3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub list_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
List.ColumnHeaders(1).Icon = LoadPicture("")
List.ColumnHeaders(2).Icon = LoadPicture("")
List.ColumnHeaders(3).Icon = LoadPicture("")
List.ColumnHeaders(4).Icon = LoadPicture("")
List.ColumnHeaders(5).Icon = LoadPicture("")
Select Case ColumnHeader
Case "User ID":
If indchk = ColumnHeader & "1" Then
List.ColumnHeaders(1).Icon = icoHeader.ListImages(1).Index
indchk = ColumnHeader & "2"
sqlstate = " Order by ID DESC"
Else
List.ColumnHeaders(1).Icon = icoHeader.ListImages(2).Index
indchk = ColumnHeader & "1"
sqlstate = " Order by ID"
End If
Case "Username":
If indchk = ColumnHeader & "1" Then
List.ColumnHeaders(2).Icon = icoHeader.ListImages(1).Index
indchk = ColumnHeader & "2"
sqlstate = " Order by Username DESC"
Else
List.ColumnHeaders(2).Icon = icoHeader.ListImages(2).Index
indchk = ColumnHeader & "1"
sqlstate = " Order by Username"
End If
Case "Name":
If indchk = ColumnHeader & "1" Then
List.ColumnHeaders(3).Icon = icoHeader.ListImages(1).Index
indchk = ColumnHeader & "2"
sqlstate = " Order by Fname DESC"
Else
List.ColumnHeaders(3).Icon = icoHeader.ListImages(2).Index
indchk = ColumnHeader & "1"
sqlstate = " Order by Fname"
End If
Case "User Level":
If indchk = ColumnHeader & "1" Then
List.ColumnHeaders(4).Icon = icoHeader.ListImages(1).Index
indchk = ColumnHeader & "2"
sqlstate = " Order by Usertype DESC"
Else
List.ColumnHeaders(4).Icon = icoHeader.ListImages(2).Index
indchk = ColumnHeader & "1"
sqlstate = " Order by Usertype"
End If
Case "Status":
If indchk = ColumnHeader & "1" Then
List.ColumnHeaders(5).Icon = icoHeader.ListImages(1).Index
indchk = ColumnHeader & "2"
sqlstate = " Order by status DESC"
Else
List.ColumnHeaders(5).Icon = icoHeader.ListImages(2).Index
indchk = ColumnHeader & "1"
sqlstate = " Order by status"
End If
End Select

Call Text1_Change
End Sub

Private Sub list_DblClick()

    teach.Enabled = True
            If List.SelectedItem.SubItems(4) = "active" Then
   blk.Caption = "Block User"
    Else
   blk.Caption = "Activate User"
    End If
    If List.SelectedItem.SubItems(3) = "Admin" Then
    info.Enabled = False
    Else
    info.Enabled = True
    End If
    PopupMenu teach, , , , blk

    
End Sub

Private Sub lvButtons_H1_Click()
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

Private Sub lvButtons_H2_Click()
viewstudrec = "user"
frmwebcam.Show 1
End Sub

Private Sub Text1_Change()
On Error Resume Next
List.ListItems.Clear
Dim filts As Integer
filts = filt
If filts = 5 Then filts = 9
If filts = 4 Then filts = 5
If filts = 2 Then filts = 4


Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from users" & sqlstate
With RS
Do Until .EOF
If LCase(Mid(.Fields(filts), 1, Len(Text1.Text))) = LCase(Text1.Text) Then
List.ListItems.Add , , .Fields(0), , ilRecordIco.ListImages(1).Index
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(1)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(4) & " " & .Fields(3)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(5)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields("status")
End If
.MoveNext
Loop
.Close
End With
If List.ListItems.Count = 0 Then
lblListInfo.Caption = "No Record"
ElseIf List.ListItems.Count = 1 Then
lblListInfo.Caption = "1 Record"
Else
lblListInfo.Caption = List.ListItems.Count & " Records"
End If
End Sub

Private Sub Timer1_Timer()
If frmmanageuser.Width >= 10435 Then
Timer1.Enabled = False
Else
frmmanageuser.Left = frmmanageuser.Left - 250
frmmanageuser.Width = frmmanageuser.Width + 500
End If
End Sub
Sub loadusrlist()
On Error Resume Next
List.ListItems.Clear
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from users" & sqlstate
With RS
Do Until .EOF
List.ListItems.Add , , .Fields(0), , ilRecordIco.ListImages(1).Index
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(1)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(4) & " " & .Fields(3)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(5)

List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields("status")
.MoveNext
Loop
.Close
End With
If List.ListItems.Count = 0 Then
lblListInfo.Caption = "No Record"
ElseIf List.ListItems.Count = 1 Then
lblListInfo.Caption = "1 Record"
Else
lblListInfo.Caption = List.ListItems.Count & " Records"
End If
Command1.Caption = "Add"
Command4.Caption = "&Block"
txtid.Enabled = False
txtuser.Enabled = False
txtpass.Enabled = False
cmbuserlevel.Enabled = False
cmbsubject.Enabled = False
txtsname.Enabled = False
lvButtons_H1.Enabled = False
lvButtons_H2.Enabled = False
txtfname.Enabled = False
txtaddress.Enabled = False
txtcnumber.Enabled = False
txtid.Text = Clear
txtuser.Text = Clear
txtpass.Text = Clear
cmbuserlevel.Text = "Select User Level"
cmbsubject.Text = "Select Subject"
txtsname.Text = Clear
txtfname.Text = Clear
txtaddress.Text = Clear
Image4.Picture = frmmain.imagelist1.ListImages(51).Picture
txtcnumber.Text = Clear
Frame1.Enabled = True
Text1.Text = Clear
Text1.Enabled = True
List.Enabled = True
End Sub
Private Function USR_AutoNum()
Dim conwan As New ADODB.Connection
Dim rswan As New ADODB.Recordset
conwan.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\db1.mdb;Persist Security Info=False"
conwan.Open
rswan.Open "select * from users Order By ID DESC", conwan, 3, 2
If rswan.RecordCount = 0 Then
    txtid.Text = "USR-0000"
Else
    txtid.Text = "USR-" & Format(Right(rswan!ID, 4) + 1, "0000")
End If
rswan.Close
End Function


Private Sub Timer2_Timer()
If frmmanageuser.Width <= 435 Then
Timer2.Enabled = False
Unload Me
Else
frmmanageuser.Left = frmmanageuser.Left + 250
frmmanageuser.Width = frmmanageuser.Width - 500
End If
End Sub

Private Sub txtcnumber_KeyPress(KeyAscii As Integer)
 If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub
Sub addrecord()
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from users"
        With RS
            .AddNew
            .Fields(0) = txtid.Text
            .Fields(1) = txtuser.Text
            .Fields(2) = txtpass.Text
            .Fields(3) = txtsname.Text
            .Fields(4) = txtfname.Text
            .Fields(5) = cmbuserlevel.Text
            If cmbuserlevel = "Teacher" Then
            .Fields(6) = "Not Set"
            Else
            .Fields(6) = "n/a"
            End If
            .Fields(7) = txtaddress.Text
            .Fields("yrtomodify") = cmbsubject.Text
            .Fields("status") = "active"
            If Len(txtcnumber.Text) = 9 Then
            .Fields(8) = "+639" & txtcnumber.Text
            End If
            .Update
            .Close
        End With
    If cmbuserlevel = "Teacher" Then
    MsgBox "Modify Subject to Teacher"
        teachersrec.txtid.Text = txtid.Text
    teachersrec.txtsname.Text = txtsname.Text
    teachersrec.txtfname.Text = txtfname.Text
    teachersrec.txtid.Text = txtid.Text
    teachersrec.Show 1
    End If
    SavePicture Image4, App.Path & "/usrpic/" & txtid.Text & ".jpg"
    MsgBox "Record Saved"
    usrlog ("Added user's record with ID '" & txtid.Text & "'")
    Call loadusrlist
End Sub
Private Sub txtuser_LostFocus()
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from users Where Username='" & txtuser.Text & "'"
        If RS.EOF = False Then
        MsgBox "Username Not available"
        txtuser.Text = Clear
        txtuser.SetFocus
        End If
End Sub
