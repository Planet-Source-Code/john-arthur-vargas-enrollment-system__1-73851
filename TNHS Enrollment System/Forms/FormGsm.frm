VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormGsm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "GSM Demo - Built with ActiveXperts SMS and MMS Toolkit"
   ClientHeight    =   9315
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Recipients:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   1920
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   8535
      Begin MSComctlLib.ListView list 
         Height          =   5295
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9340
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imagelist1"
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
            Text            =   "ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Position"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Section"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Contact Number"
            Object.Width           =   3528
         EndProperty
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
         ItemData        =   "FormGsm.frx":0000
         Left            =   2880
         List            =   "FormGsm.frx":000D
         TabIndex        =   25
         Text            =   "All"
         Top             =   480
         Width           =   2175
      End
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
         Height          =   405
         Left            =   5880
         TabIndex        =   24
         Top             =   480
         Width           =   2535
      End
      Begin TNHSES.lvButtons_H lvButtons_H1 
         Height          =   615
         Left            =   6720
         TabIndex        =   20
         Top             =   6360
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
         cBhover         =   16777088
         LockHover       =   1
         cGradient       =   16777152
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16776960
      End
      Begin TNHSES.lvButtons_H lvButtons_H2 
         Height          =   615
         Left            =   5040
         TabIndex        =   21
         Top             =   6360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         Caption         =   "Select All"
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
      Begin MSComctlLib.ImageList imagelist1 
         Left            =   0
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FormGsm.frx":002A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FormGsm.frx":19BC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Select By"
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
         Left            =   1920
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter:"
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
         Left            =   5160
         TabIndex        =   26
         Top             =   600
         Width           =   735
      End
      Begin VB.Image Image5 
         Height          =   225
         Left            =   8040
         Picture         =   "FormGsm.frx":2698
         Top             =   60
         Width           =   420
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "TNHS Enrollment System- Recipients"
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
         Left            =   120
         TabIndex        =   22
         Top             =   50
         Width           =   4215
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   7080
         Left            =   10
         Top             =   10
         Width           =   8520
      End
      Begin VB.Image Image4 
         Height          =   615
         Left            =   -1080
         Picture         =   "FormGsm.frx":5645
         Stretch         =   -1  'True
         Top             =   0
         Width           =   16155
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00400000&
      Caption         =   "Messages"
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
      Height          =   4570
      Left            =   5280
      TabIndex        =   6
      Top             =   3840
      Width           =   5895
      Begin VB.TextBox textMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   1440
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   4305
      End
      Begin VB.CommandButton CommandQuery 
         Caption         =   "&Query Delivery Status"
         Height          =   315
         Left            =   720
         TabIndex        =   8
         Top             =   4080
         Width           =   2535
      End
      Begin MSComctlLib.ListView ListViewTx 
         Height          =   1275
         Left            =   1200
         TabIndex        =   7
         Top             =   2640
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2249
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
            Text            =   "Reference"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Recipient"
            Object.Width           =   3316
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   3316
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Time"
            Object.Width           =   3316
         EndProperty
      End
      Begin TNHSES.lvButtons_H buttonSendOptions 
         Height          =   615
         Left            =   3600
         TabIndex        =   32
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         Caption         =   "Options"
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
      Begin TNHSES.lvButtons_H ButtonSend 
         Height          =   615
         Left            =   1800
         TabIndex        =   33
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Caption         =   "Send Message"
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
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2520
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   37
         Text            =   "FormGsm.frx":8972
         Top             =   720
         Width           =   3105
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Sent Items:"
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
         Left            =   120
         TabIndex        =   31
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Message:"
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
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00400000&
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
      Height          =   1875
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   5175
      Begin VB.TextBox textRecipient 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin TNHSES.lvButtons_H Command2 
         Height          =   615
         Left            =   1920
         TabIndex        =   29
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         Caption         =   "Add Recipient"
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
      Begin TNHSES.lvButtons_H Command1 
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Caption         =   "Recipients list..."
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
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "+639"
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
         Left            =   1560
         TabIndex        =   35
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "R&ecipient:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
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
      Height          =   2900
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   5295
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Recipient"
            Object.Width           =   3881
         EndProperty
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "R&ecipient List"
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
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Count: 0"
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
         Left            =   3600
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   5040
      Top             =   4440
   End
   Begin VB.Timer Timer1 
      Left            =   720
      Top             =   6000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "GSM Modem/Phone Connection Properties"
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
      Height          =   1005
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   11055
      Begin VB.ComboBox comboDevice 
         BackColor       =   &H00C0E0FF&
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
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   4200
      End
      Begin VB.ComboBox comboSpeed 
         BackColor       =   &H00C0E0FF&
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
         Left            =   9600
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Device:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Device &Speed:"
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
         Left            =   8160
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00400000&
      Caption         =   "Result"
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
      Height          =   840
      Left            =   120
      TabIndex        =   9
      Top             =   8400
      Width           =   11055
      Begin TNHSES.lvButtons_H lvButtons_H3 
         Height          =   495
         Left            =   9480
         TabIndex        =   36
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Res&ult:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label textResult 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TNHS Enrollment System- SMS Module"
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
      Left            =   60
      TabIndex        =   12
      Top             =   75
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   9300
      Left            =   0
      Top             =   0
      Width           =   11235
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   10680
      Picture         =   "FormGsm.frx":8997
      Top             =   60
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   2520
      Left            =   -6840
      Picture         =   "FormGsm.frx":B944
      Stretch         =   -1  'True
      Top             =   450
      Width           =   24480
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   -1080
      Picture         =   "FormGsm.frx":29303
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14115
   End
End
Attribute VB_Name = "FormGsm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public objConstants As AXmsCtrl.SmsConstants
Public objMessage As AXmsCtrl.SmsMessage
Public objGsm As AXmsCtrl.SmsProtocolGsm
Public objStatus As AXmsCtrl.SmsDeliveryStatus

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 260

Dim ShowReference As Boolean

Private Sub ListView1_DblClick()
If MsgBox("Delete this Recipient?", vbYesNo) = vbYes Then
ListView1.ListItems.Remove ListView1.SelectedItem.Index
End If
End Sub


Private Sub buttonReceive_Click()
    Dim NumMessages As Long
    Dim i As Long
    
    buttonReceive.Enabled = False
    
    Screen.MousePointer = vbHourglass
    
    ListViewRx.ListItems.Clear
    
    objGsm.Device = comboDevice.Text                  ' Set Device
     
    objGsm.LogFile = textLogfile.Text                 ' Set Logfile
   
    If comboSpeed.Text = "Default" Then               ' Set DeviceSpeed
        objGsm.DeviceSpeed = 0
    Else
        objGsm.DeviceSpeed = comboSpeed.Text
    End If
          
    objGsm.MessageStorage = FormGsmRecvConfig.comboStore.ListIndex             ' Set selected message store
    
    NumMessages = objGsm.Receive                      ' Retrieve messages
            
    If GetResult = 0 Then                     ' Success?
            
        For i = 0 To NumMessages - 1
            On Error Resume Next
            Set objMessage = objGsm.GetFirstMessage
            On Error GoTo 0
            
            If GetResult = 0 Then
                Dim lList As ListItem
            
                Set lList = ListViewRx.ListItems.Add(, , objMessage.Time)   ' Add data to list control
            
                lList.SubItems(1) = objMessage.Sender
                lList.SubItems(2) = objMessage.Data
            End If
        Next
    End If
    
    If (FormGsmRecvConfig.checkDelete.Value) Then
        objGsm.DeleteAllMessages
    End If
    
    Screen.MousePointer = vbDefault
    
    buttonReceive.Enabled = True
    
End Sub

Private Sub buttonReceiveOptions_Click()
    FormGsmRecvConfig.Show 1
End Sub

Private Sub buttonSend_Click()
If ListView1.ListItems.Count = 0 Then
MsgBox "No Recipient to send"
textRecipient.Text = Clear
Else
MsgBox Text2.Text & vbNewLine & textMessage.Text
    ButtonSend.Enabled = False
    textResult.Caption = "Sending message, Please wait..."
      Call sendmsg
    End If
End Sub

Private Sub buttonSendOptions_Click()
    FormSendConfig.Show 1
End Sub
Sub sendmsg()
Dim bc As Integer
bc = ListView1.ListItems.Count
Do Until bc = 0
  Dim MessageType As Long
    Dim strReference As String
    textResult.Refresh
    objGsm.Device = comboDevice.Text
    If comboSpeed.ListIndex = 0 Then
        objGsm.DeviceSpeed = 0               ' use default speed
        Else
        objGsm.DeviceSpeed = comboSpeed.List(comboSpeed.ListIndex)
    End If
        
    ' Create Message Object
    Set objMessage = CreateObject("ActiveXperts.SmsMessage")
     
    ' Set Message Format
    objMessage.Format = objConstants.asMESSAGEFORMAT_TEXT
     
    If FormSendConfig.checkMultipart.Value = 1 Then
        objMessage.Format = objConstants.asMESSAGEFORMAT_TEXT_MULTIPART
    End If
            
    If FormSendConfig.checkFlash.Value = 1 Then
        objMessage.Format = objConstants.asMESSAGEFORMAT_TEXT_FLASH
    End If
    
    If FormSendConfig.checkUDH.Value = 1 Then
        objMessage.Format = objConstants.asMESSAGEFORMAT_DATA_UDH
    End If
    
    ' Set Delivery Report
    objMessage.RequestDeliveryStatus = 0
        
    ' Set recipient
    objMessage.Recipient = ListView1.ListItems.Item(bc).SubItems(1)
    
    ' Set Message parameters
    objMessage.Data = Text2.Text & vbNewLine & textMessage.Text
      
    ' Send the message
    strReference = objGsm.Send(objMessage)
        
    ' Display result
    If GetResult() = 0 Then
        Dim lList As ListItem
        
        Set lList = ListViewTx.ListItems.Add(, , strReference)   ' Add data to list control
        
        lList.SubItems(1) = objMessage.Recipient
           lList.SubItems(2) = "Sent"
           lList.SubItems(3) = Time
        lList.Tag = 0
        MsgBox "Message Sent To: " & ListView1.ListItems.Item(bc).SubItems(1)
    Else
        If FormSendConfig.checkUDH.Value = 1 Then
            FormSendConfig.checkUDH.Value = 0
            textMessage.Text = ""
        End If
        bc = 1
    End If
    ListView1.ListItems.Remove bc
bc = bc - 1
Loop
    ButtonSend.Enabled = True
End Sub
Private Sub Combo1_Click()
Text1.Text = Clear
If Combo1.Text = "All" Then
List.ListItems.Clear
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from users Order by Sname"
Do Until RS.EOF
If RS.Fields("Cnumber") <> Empty Then
List.ListItems.Add , , RS.Fields(0), , ImageList1.ListImages(1).Index
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("fname") & " " & RS.Fields("sname")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("Usertype")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "n/a"
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("Cnumber")

End If
RS.MoveNext
Loop
RS.Close

Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblstudents  Status='Enrolled' Order by sname"
Do Until RS.EOF
If RS.Fields("cpno") <> Empty Then
List.ListItems.Add , , RS.Fields(0), , ImageList1.ListImages(2).Index
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("fname") & " " & RS.Fields("sname")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "Students"
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("yrsec")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("cpno")
End If
RS.MoveNext
Loop
RS.Close
ElseIf Combo1.Text = "Students" Then
List.ListItems.Clear
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblstudents  Status='Enrolled' Order by sname"
Do Until RS.EOF
If RS.Fields("cpno") <> Empty Then
List.ListItems.Add , , RS.Fields(0), , ImageList1.ListImages(2).Index
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("fname") & " " & RS.Fields("sname")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "Students"
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("yrsec")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("cpno")
End If
RS.MoveNext
Loop
RS.Close

Else
List.ListItems.Clear
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from users Order by Sname"
Do Until RS.EOF
If RS.Fields("Cnumber") <> Empty Then
List.ListItems.Add , , RS.Fields(0), , ImageList1.ListImages(1).Index
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("fname") & " " & RS.Fields("sname")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("Usertype")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "n/a"
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("Cnumber")
End If
RS.MoveNext
Loop
RS.Close
End If

End Sub

Private Sub Command1_Click()
Combo1.Text = "All"
Frame6.Visible = True
List.ListItems.Clear
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from users Order by Sname"
Do Until RS.EOF
If RS.Fields("Cnumber") <> Empty Then
List.ListItems.Add , , RS.Fields(0), , ImageList1.ListImages(1).Index
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("fname") & " " & RS.Fields("sname")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("Usertype")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "n/a"
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("Cnumber")

End If
RS.MoveNext
Loop
RS.Close

Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblstudents Where Status='Enrolled' Order by sname"
Do Until RS.EOF
If RS.Fields("cpno") <> Empty Then
List.ListItems.Add , , RS.Fields(0), , ImageList1.ListImages(2).Index
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("fname") & " " & RS.Fields("sname")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "Students"
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("yrsec")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("cpno")
End If
RS.MoveNext
Loop
RS.Close
End Sub

Private Sub Command2_Click()
If Len(textRecipient.Text) = 9 Then
ListView1.ListItems.Add , , "n/a"
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , "+639" & textRecipient.Text
textRecipient.Text = Clear
End If
End Sub

Private Sub CommandQuery_Click()
    Dim i As Integer
    
    CommandQuery.Enabled = False
    Screen.MousePointer = vbHourglass
    
    For i = 1 To ListViewTx.ListItems.Count
        Dim lList As ListItem
        
        Set lList = ListViewTx.ListItems(i)
            
        If (lList.Tag = 0) Then
            Set objStatus = objGsm.QueryStatus(lList.Text)
          
            If (GetResult = 0) Then
                lList.SubItems(2) = objStatus.StatusDescription
                lList.SubItems(3) = objStatus.StatusCompletedTime
                
                lList.Tag = objStatus.IsCompleted
            End If
        End If
    Next
    
    Screen.MousePointer = vbDefault
    CommandQuery.Enabled = True
End Sub


Private Sub CommandWap_Click()
    FormWap.Show vbModal
    
    If (FormWap.Result = True) Then
        
        textMessage.Text = FormWap.strMessage
        
        FormSendConfig.checkFlash = 0
        FormSendConfig.checkMultipart = 0
        FormSendConfig.checkReport = 0
        FormSendConfig.checkUDH = 1
        
    End If
End Sub

Private Sub Form_Load()
Me.Width = 235
Call dbConnection
Frame6.Visible = False
    Dim lDeviceCount As Long
    Dim i As Long
            
    ShowReference = False
    
    Set objGsm = CreateObject("ActiveXperts.SmsProtocolGsm")
    Set objConstants = CreateObject("ActiveXperts.SmsConstants")
    
    lDeviceCount = objGsm.GetDeviceCount()  ' Get number of  devices
    
    For i = 0 To lDeviceCount - 1
        comboDevice.AddItem (objGsm.GetDevice(i)) ' Add devices to list box
    Next
             
    FormSendConfig.checkMultipart = 1
    FormSendConfig.checkReport = 1
    
    comboDevice.AddItem ("COM1")        ' Add serial devices
    comboDevice.AddItem ("COM2")
    comboDevice.AddItem ("COM3")
    comboDevice.AddItem ("COM4")
    comboDevice.AddItem ("COM5")
    comboDevice.AddItem ("COM6")
    comboDevice.AddItem ("COM7")
    comboDevice.AddItem ("COM8")
    
    comboDevice.ListIndex = 0
        
    comboSpeed.AddItem ("Default")      ' Setup devicespeed combo
    comboSpeed.AddItem ("1200")
    comboSpeed.AddItem ("2400")
    comboSpeed.AddItem ("9600")
    comboSpeed.AddItem ("19200")
    comboSpeed.AddItem ("38400")
    comboSpeed.AddItem ("57600")
    comboSpeed.AddItem ("115200")
    
    comboSpeed.ListIndex = 0
    
End Sub

Public Function GetResult() As Long
    
    Dim lResult As Long
    
    lResult = objGsm.LastError
    
    textResult.Caption = "ERROR " & lResult & " : " & objGsm.GetErrorDescription(lResult)     ' Set Result
        
    GetResult = lResult
   
End Function
    
Public Function FileExists(sFileName As String) As Boolean
  FileExists = CBool(Len(Dir$(sFileName))) And CBool(Len(sFileName))
End Function
Private Sub Form_Unload(Cancel As Integer)
    Set objGsm = Nothing
    Set objConstants = Nothing
End Sub

Private Sub LinkNumberFormat_Click()
    Shell "explorer " & Chr(34) & "http://www.activexperts.com/support/xmstoolkit?kb=Q4200015" & Chr(34)
End Sub


Private Sub Image3_Click()
Unload Me
End Sub

Private Sub Image5_Click()
Frame6.Visible = False
End Sub

Private Sub list_DblClick()
If MsgBox("Add this Recipient to List?", vbYesNo) = vbYes Then
ListView1.ListItems.Add , , List.SelectedItem.SubItems(1)
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , List.SelectedItem.SubItems(4)
MsgBox "Added to recipients"
End If
End Sub


Private Sub lvButtons_H1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Frame6.Visible = False
End Sub

Private Sub lvButtons_H2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If MsgBox("Select All Recipients?", vbYesNo) = vbYes Then
Dim os As Integer
os = List.ListItems.Count
Do Until os = 0
ListView1.ListItems.Add , , List.ListItems.Item(os).SubItems(1)
ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add , , List.ListItems.Item(os).SubItems(4)
os = os - 1
Loop
MsgBox "Added to recipients"
End If

End Sub

Private Sub lvButtons_H3_Click()
Unload Me
End Sub



Private Sub Text1_Change()
On Error Resume Next
If Combo1.Text = "All" Then
List.ListItems.Clear
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from users Order by Sname"
Do Until RS.EOF
If RS.Fields("Cnumber") <> Empty Then
If LCase(Mid(RS.Fields(3), 1, Len(Text1.Text))) = LCase(Text1.Text) Or LCase(Mid(RS.Fields(4), 1, Len(Text1.Text))) = LCase(Text1.Text) Or LCase(Mid(RS.Fields(0), 1, Len(Text1.Text))) = LCase(Text1.Text) Then
List.ListItems.Add , , RS.Fields(0), , ImageList1.ListImages(1).Index
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("fname") & " " & RS.Fields("sname")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("Usertype")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "n/a"
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("Cnumber")
End If
End If
RS.MoveNext
Loop
RS.Close

Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblstudents Order by sname"
Do Until RS.EOF
If RS.Fields("cpno") <> Empty Then
If LCase(Mid(RS.Fields(1), 1, Len(Text1.Text))) = LCase(Text1.Text) Or LCase(Mid(RS.Fields(2), 1, Len(Text1.Text))) = LCase(Text1.Text) Or LCase(Mid(RS.Fields("yrsec"), 1, Len(Text1.Text))) = LCase(Text1.Text) Or LCase(Mid(RS.Fields(0), 1, Len(Text1.Text))) = LCase(Text1.Text) Then
List.ListItems.Add , , RS.Fields(0), , ImageList1.ListImages(2).Index
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("fname") & " " & RS.Fields("sname")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "Students"
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("yrsec")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("cpno")
End If
End If
RS.MoveNext
Loop
RS.Close
ElseIf Combo1.Text = "Students" Then
List.ListItems.Clear
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblstudents Order by sname"
Do Until RS.EOF
If RS.Fields("cpno") <> Empty Then
If LCase(Mid(RS.Fields(1), 1, Len(Text1.Text))) = LCase(Text1.Text) Or LCase(Mid(RS.Fields(2), 1, Len(Text1.Text))) = LCase(Text1.Text) Or LCase(Mid(RS.Fields("yrsec"), 1, Len(Text1.Text))) = LCase(Text1.Text) Or LCase(Mid(RS.Fields(0), 1, Len(Text1.Text))) = LCase(Text1.Text) Then
List.ListItems.Add , , RS.Fields(0), , ImageList1.ListImages(2).Index
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("fname") & " " & RS.Fields("sname")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "Students"
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("yrsec")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("cpno")
End If
End If
RS.MoveNext
Loop
RS.Close
Else
List.ListItems.Clear
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from users Order by Sname"
Do Until RS.EOF
If RS.Fields("Cnumber") <> Empty Then
If LCase(Mid(RS.Fields(3), 1, Len(Text1.Text))) = LCase(Text1.Text) Or LCase(Mid(RS.Fields(4), 1, Len(Text1.Text))) = LCase(Text1.Text) Or LCase(Mid(RS.Fields(0), 1, Len(Text1.Text))) = LCase(Text1.Text) Then
List.ListItems.Add , , RS.Fields(0), , ImageList1.ListImages(1).Index
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("fname") & " " & RS.Fields("sname")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("Usertype")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , "n/a"
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , RS.Fields("Cnumber")
End If
End If
RS.MoveNext
Loop
RS.Close
End If
End Sub



Private Sub textRecipient_KeyPress(KeyAscii As Integer)
If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack) And KeyAscii <> 13) Then KeyAscii = 0
If KeyAscii = 13 Then
Call Command2_Click
End If
End Sub

Private Sub Timer1_Timer()
Label10.Caption = "Recipient Count: " & ListView1.ListItems.Count
End Sub

Private Sub Timer2_Timer()
If Me.Width >= 11235 Then
Timer2.Enabled = False
Else
Me.Left = Me.Left - 250
Me.Width = Me.Width + 500
End If
End Sub
