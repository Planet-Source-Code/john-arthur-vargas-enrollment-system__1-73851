VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmmappop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "frmmappop"
   ClientHeight    =   10095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TNHSES.lvButtons_H lvButtons_H1 
      Height          =   320
      Left            =   6240
      TabIndex        =   4
      Top             =   2400
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   1
      Left            =   840
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   52
      ImageHeight     =   53
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   42
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":08F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":1185
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":1A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":273B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":2CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":31EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":37F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":3E5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":454E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":4B1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":53C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":5C0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":63B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":709B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":7D91
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":83B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":89E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":8F29
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":94F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":9B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":DC66
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":1190D
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":15806
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":1906B
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":1C2B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":1F197
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":21DB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":23871
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":25402
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":28895
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":2BA9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":2FCCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":33AB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":3757B
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":3C8BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":41F4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":45231
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":484AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":4B199
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":4E255
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7500
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   13229
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777088
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
      TabPicture(0)   =   "frmmappop.frx":51667
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "2nd Floor"
      TabPicture(1)   =   "frmmappop.frx":51683
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   7185
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   9735
         Begin VB.Image Image3 
            Height          =   705
            Index           =   8
            Left            =   4125
            Top             =   270
            Width           =   900
         End
         Begin VB.Image Image3 
            Height          =   705
            Index           =   9
            Left            =   4900
            Picture         =   "frmmappop.frx":5169F
            Top             =   270
            Width           =   960
         End
         Begin VB.Image Image3 
            Height          =   705
            Index           =   5
            Left            =   1730
            Top             =   285
            Width           =   645
         End
         Begin VB.Image Image3 
            Height          =   705
            Index           =   6
            Left            =   2270
            Picture         =   "frmmappop.frx":51CF1
            Top             =   285
            Width           =   735
         End
         Begin VB.Image Image3 
            Height          =   495
            Index           =   11
            Left            =   6005
            Picture         =   "frmmappop.frx":52298
            Top             =   1410
            Width           =   1095
         End
         Begin VB.Image Image3 
            Height          =   735
            Index           =   10
            Left            =   6000
            Picture         =   "frmmappop.frx":52855
            Top             =   285
            Width           =   915
         End
         Begin VB.Image Image3 
            Height          =   690
            Index           =   7
            Left            =   3600
            Picture         =   "frmmappop.frx":52F38
            Top             =   285
            Width           =   540
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H8000000D&
            BorderWidth     =   3
            Height          =   7140
            Left            =   0
            Top             =   0
            Width           =   8985
         End
         Begin VB.Image Image3 
            Height          =   1005
            Index           =   4
            Left            =   6425
            Picture         =   "frmmappop.frx":53421
            Top             =   2310
            Width           =   975
         End
         Begin VB.Image Image3 
            Height          =   1170
            Index           =   3
            Left            =   6590
            Picture         =   "frmmappop.frx":53B8A
            Top             =   3300
            Width           =   1080
         End
         Begin VB.Image Image3 
            Height          =   1260
            Index           =   2
            Left            =   6815
            Picture         =   "frmmappop.frx":543F7
            Top             =   4440
            Width           =   1110
         End
         Begin VB.Image Image3 
            Height          =   1200
            Index           =   1
            Left            =   6995
            Picture         =   "frmmappop.frx":54C75
            Top             =   5670
            Width           =   1155
         End
         Begin VB.Image Image2 
            Height          =   7065
            Left            =   15
            Picture         =   "frmmappop.frx":5555C
            Top             =   30
            Width           =   8925
         End
         Begin VB.Image Image4 
            Height          =   10500
            Left            =   -720
            Picture         =   "frmmappop.frx":5F9B7
            Top             =   0
            Width           =   10500
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   7140
         Left            =   -75000
         TabIndex        =   3
         Top             =   360
         Width           =   9015
         Begin VB.Image Image3 
            Height          =   735
            Index           =   17
            Left            =   1670
            Picture         =   "frmmappop.frx":7FDA5
            Top             =   290
            Width           =   675
         End
         Begin VB.Image Image3 
            Height          =   735
            Index           =   19
            Left            =   3590
            Picture         =   "frmmappop.frx":803B7
            Top             =   270
            Width           =   585
         End
         Begin VB.Image Image3 
            Height          =   720
            Index           =   20
            Left            =   4110
            Picture         =   "frmmappop.frx":808EA
            Top             =   290
            Width           =   930
         End
         Begin VB.Image Image3 
            Height          =   735
            Index           =   18
            Left            =   2210
            Picture         =   "frmmappop.frx":80EA6
            Top             =   290
            Width           =   780
         End
         Begin VB.Image Image3 
            Height          =   1455
            Index           =   15
            Left            =   3050
            Picture         =   "frmmappop.frx":814C9
            Top             =   2640
            Width           =   1380
         End
         Begin VB.Image Image3 
            Height          =   1650
            Index           =   16
            Left            =   3840
            Picture         =   "frmmappop.frx":8219E
            Top             =   3450
            Width           =   1530
         End
         Begin VB.Image Image3 
            Height          =   735
            Index           =   21
            Left            =   6050
            Picture         =   "frmmappop.frx":82E84
            Top             =   270
            Width           =   915
         End
         Begin VB.Image Image3 
            Height          =   1020
            Index           =   14
            Left            =   6430
            Picture         =   "frmmappop.frx":834E1
            Top             =   2310
            Width           =   1110
         End
         Begin VB.Image Image3 
            Height          =   1170
            Index           =   13
            Left            =   6630
            Picture         =   "frmmappop.frx":83C78
            Top             =   3310
            Width           =   1035
         End
         Begin VB.Image Image3 
            Height          =   1245
            Index           =   12
            Left            =   6810
            Picture         =   "frmmappop.frx":844B7
            Top             =   4470
            Width           =   1095
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H8000000D&
            BorderWidth     =   3
            Height          =   7140
            Left            =   0
            Top             =   0
            Width           =   8985
         End
         Begin VB.Image Image7 
            Height          =   10500
            Left            =   -720
            Picture         =   "frmmappop.frx":84D4C
            Top             =   0
            Width           =   10500
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   1080
      Top             =   2880
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
            Picture         =   "frmmappop.frx":93322
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":97475
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":9B07F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":9EDFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":A2615
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":A5690
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":A84BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":AB034
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":ACAE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":AE695
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":B1A6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":B4B96
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":B8C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":BC9D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":C0389
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":C576C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":CAF30
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":CDF5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":D116C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":D3DD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":D6E10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   2
      Left            =   1200
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   43
      ImageHeight     =   47
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":DA192
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":DD3F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":E0570
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":E376A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":E673A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":E9726
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":EC7B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":EFA23
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":F4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":FA2B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":FF6EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":104BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":107E1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":10B1C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmappop.frx":10E358
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image8 
      Height          =   990
      Left            =   3360
      Top             =   5040
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   7455
      Left            =   240
      Top             =   2400
      Width           =   9015
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "TNHS Enrollment System-Room Schedule(Morning Session)"
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
      Width           =   6015
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   9000
      Picture         =   "frmmappop.frx":11146F
      Top             =   60
      Width           =   420
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   10085
      Index           =   0
      Left            =   10
      Top             =   10
      Width           =   9530
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   -600
      Picture         =   "frmmappop.frx":11441C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10155
   End
   Begin VB.Image Image1 
      Height          =   2040
      Left            =   -9600
      Picture         =   "frmmappop.frx":118012
      Stretch         =   -1  'True
      Top             =   480
      Width           =   27480
   End
End
Attribute VB_Name = "frmmappop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim btn As Integer
Dim a As Double
Dim mapcount(2, 21) As Integer

Private Sub Form_Load()
If session = 1 Then
Label20.Caption = "TNHS Enrollment System-Room Schedule(Morning Session)"
Else
Label20.Caption = "TNHS Enrollment System-Room Schedule(Afternoon Session)"
End If
Call dbConnection
Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from roommap"
        With RS
        Do Until .EOF
        If .Fields(2) <> "vacant" Then
        mapcount(1, .Fields(0)) = 1
        Else
        mapcount(1, .Fields(0)) = 0
        End If
        If .Fields(3) <> "vacant" Then
        mapcount(2, .Fields(0)) = 1
        Else
        mapcount(2, .Fields(0)) = 0
        End If
        .MoveNext
        Loop
        End With
        For a = 1 To 21
    If mapcount(session, a) = 1 Then
    Image3(a).Picture = ImageList1(0).ListImages(a).Picture
    Else
    Image3(a).Picture = ImageList1(1).ListImages((a + (ImageList1(1).ListImages.Count / 2))).Picture

    End If
    Next a
If mapcount(session, 5) = 0 And mapcount(session, 6) = 0 Then
ElseIf mapcount(session, 5) = 0 And mapcount(session, 6) = 1 Then
Image3(5).Picture = ImageList1(2).ListImages(1).Picture
ElseIf mapcount(session, 5) = 1 And mapcount(session, 6) = 0 Then
Image3(5).Picture = ImageList1(2).ListImages(3).Picture
ElseIf mapcount(session, 5) = 1 And mapcount(session, 6) = 1 Then
Image3(5).Picture = ImageList1(2).ListImages(2).Picture
End If
If mapcount(session, 8) = 0 And mapcount(session, 9) = 0 Then
Image3(8).Picture = ImageList1(2).ListImages(4).Picture
ElseIf mapcount(session, 8) = 0 And mapcount(session, 9) = 1 Then
Image3(8).Picture = ImageList1(2).ListImages(5).Picture
ElseIf mapcount(session, 8) = 1 And mapcount(session, 9) = 0 Then
Image3(8).Picture = ImageList1(2).ListImages(7).Picture
ElseIf mapcount(session, 8) = 1 And mapcount(session, 9) = 1 Then
Image3(8).Picture = ImageList1(2).ListImages(6).Picture
End If

If mapcount(session, 15) = 0 And mapcount(session, 16) = 0 Then
Image3(15).Picture = ImageList1(2).ListImages(8).Picture
ElseIf mapcount(session, 15) = 0 And mapcount(session, 16) = 1 Then
Image3(15).Picture = ImageList1(2).ListImages(9).Picture
ElseIf mapcount(session, 15) = 1 And mapcount(session, 16) = 0 Then
Image3(15).Picture = ImageList1(2).ListImages(10).Picture
ElseIf mapcount(session, 15) = 1 And mapcount(session, 16) = 1 Then
Image3(15).Picture = ImageList1(2).ListImages(11).Picture
End If
If mapcount(session, 17) = 0 And mapcount(session, 18) = 0 Then
Image3(17).Picture = ImageList1(2).ListImages(12).Picture
ElseIf mapcount(session, 17) = 0 And mapcount(session, 18) = 1 Then
Image3(17).Picture = ImageList1(2).ListImages(13).Picture
ElseIf mapcount(session, 17) = 1 And mapcount(session, 18) = 0 Then
Image3(17).Picture = ImageList1(2).ListImages(14).Picture
ElseIf mapcount(session, 17) = 1 And mapcount(session, 18) = 1 Then
Image3(17).Picture = ImageList1(2).ListImages(15).Picture
End If
        RS.Close
End Sub



Private Sub Image3_DblClick(Index As Integer)
mapid = Index
frmmapinfo.Show 1
End Sub


'Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'a = (ImageList1(1).ListImages.Count / 2)
'Do Until a = 0
'Image3(a).Visible = True
'Image3(a).Picture = ImageList1(1).ListImages(a).Picture
'a = a - 1
'Loop
'End Sub

Private Sub Image5_Click()
Unload Me
End Sub


Private Sub lvButtons_H1_Click()
Unload Me
End Sub

