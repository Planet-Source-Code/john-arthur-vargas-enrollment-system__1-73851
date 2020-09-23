VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmviewstudrec 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   -120
   ClientWidth     =   9090
   LinkTopic       =   "Form4"
   ScaleHeight     =   7170
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      ItemData        =   "frmviewstudrecord.frx":0000
      Left            =   1200
      List            =   "frmviewstudrecord.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   840
      Width           =   1935
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
      Left            =   3360
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin MSComctlLib.ListView list 
      Height          =   5655
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9975
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imagelist1"
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
         Text            =   "Year"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Enrollment Type"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Enroll Status"
         Object.Width           =   2
      EndProperty
   End
   Begin MSComctlLib.ImageList imagelist1 
      Left            =   1320
      Top             =   0
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
            Picture         =   "frmviewstudrecord.frx":003A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":19CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":26A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":403A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":59CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":735E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":8CF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":99CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":A6A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":AF80
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":BC5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":C938
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":D21C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":DEF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":E7D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":F4B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":10E44
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":127D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":130B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":1398E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":14268
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":14B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":1541C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":159B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":15CD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":15FEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":168C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":1719E
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":17A78
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":18752
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":18A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":18EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":19D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":1A162
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":1AA3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":1B316
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":20B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":213E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":216FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":223D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":22CB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":2358A
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":23E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":2473E
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":25018
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":258F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":261CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":34AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":36C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":4183D
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":4202C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList icoHeader 
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":42408
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewstudrecord.frx":429A2
            Key             =   ""
         EndProperty
      EndProperty
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   5535
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   7320
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Dropped"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Not Enrolled"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   5280
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblListInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00D8E9EC&
      BackStyle       =   0  'Transparent
      Caption         =   "No Record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C25418&
      Height          =   195
      Left            =   8040
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   8565
      Picture         =   "frmviewstudrecord.frx":42F3C
      Top             =   60
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   10200
      Picture         =   "frmviewstudrecord.frx":45EE9
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "TNHS Enrollment System(Student's Record)"
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
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   7155
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   9090
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmviewstudrecord.frx":48E96
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9195
   End
End
Attribute VB_Name = "frmviewstudrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim asd As Integer, ct As Integer, sqlstate As String, indchk As String

Private Sub Command3_Click()
Unload Me
End Sub


Private Sub Form_Load()
sqlstate = ""
Combo1.ListIndex = 0
Image4.Picture = ImageList1.ListImages(4).Picture
Image6.Picture = ImageList1.ListImages(13).Picture
Call dbConnection
List.ListItems.Clear
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
If viewstudrec = "regular" Then
If userlevel = "Admin" Then
RS.Open "select * from tblstudents Where Enrolltype='Regular' And evaluated='No' Order by Status"
Else
Label5.Caption = "Previous Advisory Class Section: " & prevyear
RS.Open "select * from tblstudents Where Enrolltype='Regular' And Status='Not Enrolled' And evaluated='No' and yrsec='" & prevyear & "' Order by studno DESC"
Call visibility
End If
Label14.Caption = "TNHS Enrollment System(Student's Record)-Regular Students"
ElseIf viewstudrec = "transferee" Then
RS.Open "select * from tblstudents Where Enrolltype='Transferee' and Status='Not Enrolled' Order by studno DESC"
Label14.Caption = "TNHS Enrollment System(Student's Record)-Transferee Students"
Call visibility
ElseIf viewstudrec = "studinfo" Then
Label14.Caption = "TNHS Enrollment System(Student's Record)"
RS.Open "select * from tblstudents Where Status='Enrolled' Order by studno DESC"
Call visibility
ElseIf viewstudrec = "freshmen" Then
RS.Open "select * from tblstudents Where Enrolltype='Freshmen' and Status='Not Enrolled' Order by studno DESC"
Label14.Caption = "TNHS Enrollment System(Student's Record)-Freshmen Students"
Call visibility
End If
With RS
Do Until .EOF
List.ListItems.Add , , .Fields(0)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(2) & " " & Mid(.Fields(3), 1, 1) & ". " & .Fields(1)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields("gender")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(10)

List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(14)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(13)
If .Fields(13) = "Enrolled" Then
List.ListItems.Item(List.ListItems.Count).SmallIcon = ImageList1.ListImages(3).Index
ElseIf .Fields(13) = "Not Enrolled" Then
List.ListItems.Item(List.ListItems.Count).SmallIcon = ImageList1.ListImages(4).Index
Else
List.ListItems.Item(List.ListItems.Count).SmallIcon = ImageList1.ListImages(13).Index
End If
.MoveNext
Loop
.Close
End With
frmuserlog.Width = 155
If List.ListItems.Count = 0 Then
lblListInfo.Caption = "No Record"
ElseIf List.ListItems.Count = 1 Then
lblListInfo.Caption = "1 Record"
Else
lblListInfo.Caption = List.ListItems.Count & " Records"
End If
End Sub

Private Sub Image2_Click()
Unload Me
End Sub
Sub visibility()
Label3.Visible = False
Label4.Visible = False
Image4.Visible = False
Image6.Visible = False
End Sub


Private Sub Image5_Click()

End Sub



Private Sub list_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
List.ColumnHeaders(1).Icon = LoadPicture("")
List.ColumnHeaders(2).Icon = LoadPicture("")
List.ColumnHeaders(3).Icon = LoadPicture("")
List.ColumnHeaders(4).Icon = LoadPicture("")
List.ColumnHeaders(5).Icon = LoadPicture("")
Select Case ColumnHeader
Case "Student Number":
If indchk = ColumnHeader & "1" Then
List.ColumnHeaders(1).Icon = icoHeader.ListImages(1).Index
indchk = ColumnHeader & "2"
sqlstate = " Order by studno DESC"
Else
List.ColumnHeaders(1).Icon = icoHeader.ListImages(2).Index
indchk = ColumnHeader & "1"
sqlstate = " Order by studno"
End If
Case "Name":
If indchk = ColumnHeader & "1" Then
List.ColumnHeaders(2).Icon = icoHeader.ListImages(1).Index
indchk = ColumnHeader & "2"
sqlstate = " Order by fname DESC"
Else
List.ColumnHeaders(2).Icon = icoHeader.ListImages(2).Index
indchk = ColumnHeader & "1"
sqlstate = " Order by fname"
End If
Case "Gender":
If indchk = ColumnHeader & "1" Then
List.ColumnHeaders(3).Icon = icoHeader.ListImages(1).Index
indchk = ColumnHeader & "2"
sqlstate = " Order by gender DESC"
Else
List.ColumnHeaders(3).Icon = icoHeader.ListImages(2).Index
indchk = ColumnHeader & "1"
sqlstate = " Order by gender"
End If
Case "Year":
If indchk = ColumnHeader & "1" Then
List.ColumnHeaders(4).Icon = icoHeader.ListImages(1).Index
indchk = ColumnHeader & "2"
sqlstate = " Order by years DESC"
Else
List.ColumnHeaders(4).Icon = icoHeader.ListImages(2).Index
indchk = ColumnHeader & "1"
sqlstate = " Order by years"
End If
Case "Enrollment Type"
If indchk = ColumnHeader & "1" Then
List.ColumnHeaders(5).Icon = icoHeader.ListImages(1).Index
indchk = ColumnHeader & "2"
sqlstate = " Order by Enrolltype DESC"
Else
List.ColumnHeaders(5).Icon = icoHeader.ListImages(2).Index
indchk = ColumnHeader & "1"
sqlstate = " Order by Enrolltype"
End If
End Select

Call Text1_Change
End Sub

Private Sub list_DblClick()
On Error GoTo x
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblstudents Where studno='" & List.SelectedItem.Text & "'"
With RS
If .EOF = False Then
If viewstudrec = "freshmen" Then
frmfreshmen.txt1.Text = .Fields(0)
frmfreshmen.txt2.Text = .Fields(1)
frmfreshmen.txt3.Text = .Fields(2)
frmfreshmen.txt4.Text = .Fields(3)
frmfreshmen.txt5.Text = .Fields(4)
frmfreshmen.txt6.Text = Format(.Fields(5), "MM/DD/YYYY")
frmfreshmen.txt7.Text = Year(Now) - Year(.Fields(5))
frmfreshmen.txt8.Text = .Fields(7)
frmfreshmen.Combo1.Text = .Fields("gender")
If .Fields("goodmoral") = "Yes" Then frmfreshmen.Check3.Value = 1
If .Fields("schoolcard") = "Yes" Then frmfreshmen.Check1.Value = 1
If .Fields("bc") = "No" Then frmfreshmen.Check2.Value = 1
If .Fields("form137") = "Yes" Then frmfreshmen.Check4.Value = 1
    frmfreshmen.Text1.Text = .Fields("grade")
    frmfreshmen.Text2.Text = .Fields("english")
     frmfreshmen.Text3.Text = .Fields("math")
     frmfreshmen.Text5.Text = .Fields("science")
     frmfreshmen.Text6.Text = .Fields("filipino")
frmfreshmen.txt9.Text = .Fields(8)
frmfreshmen.txt10.Text = .Fields(9)
frmfreshmen.Text4.Text = .Fields(18)
frmfreshmen.txt11.Text = Mid(.Fields(12), 5, 9)

frmfreshmen.Command3.Caption = "Cancel"
frmfreshmen.Command4.Caption = "Edit"
frmfreshmen.Image4.Picture = LoadPicture(App.Path & "/students/" & .Fields(0) & ".jpg")
asd = 3
ElseIf viewstudrec = "regular" Then
If List.SelectedItem.SubItems(3) = "4th" And List.SelectedItem.SubItems(5) = "Not Enrolled" Then
If MsgBox("The student you have selected is 4th year in previous schoolyear." & vbNewLine & "Add this student to graduate list?" & vbNewLine & "Press Yes to add, no to Schedule his/her new section", vbQuestion + vbYesNo) = vbYes Then
i_Clear.cLearMe frmregular
frmgrad.Show 1
Else
frmregular.Command3.Caption = "Cancel"
frmregular.cmdenroll.Enabled = True
frmregular.Command4.Enabled = True
Call regularload
Unload Me
End If

Else
frmregular.Command3.Caption = "Cancel"
frmregular.cmdenroll.Enabled = True
frmregular.Command4.Enabled = True
Call regularload
Unload Me
End If

ElseIf viewstudrec = "transferee" Then
frmtransferee.txt1.Text = .Fields(0)
frmtransferee.txt2.Text = .Fields(1)
frmtransferee.txt3.Text = .Fields(2)
frmtransferee.txt4.Text = .Fields(3)
frmtransferee.txt5.Text = .Fields(4)
frmtransferee.txt6.Text = Format(.Fields(5), "MM/DD/YYYY")
frmtransferee.txt7.Text = Year(Now) - Year(.Fields(5))
frmtransferee.Combo1.Text = .Fields("gender")
If .Fields("goodmoral") = "Yes" Then frmtransferee.Check3.Value = 1
If .Fields("schoolcard") = "Yes" Then frmtransferee.Check1.Value = 1
If .Fields("bc") = "No" Then frmtransferee.Check2.Value = 1
If .Fields("form137") = "Yes" Then frmtransferee.Check4.Value = 1
    frmtransferee.Text1.Text = .Fields("grade")
        frmtransferee.Text2.Text = .Fields("english")
     frmtransferee.Text3.Text = .Fields("math")
     frmtransferee.Text5.Text = .Fields("science")
     frmtransferee.Text6.Text = .Fields("filipino")
frmtransferee.txt8.Text = .Fields(7)
frmtransferee.txt9.Text = .Fields(8)
frmtransferee.txt10.Text = .Fields(9)
frmtransferee.cmb11.Text = .Fields(10)
frmtransferee.txt11.Text = Mid(.Fields(12), 5, 9)
frmtransferee.Text4.Text = .Fields(18)
frmtransferee.Image4.Picture = LoadPicture(App.Path & "/students/" & .Fields(0) & ".jpg")
frmtransferee.Command3.Caption = "Cancel"
frmtransferee.Command4.Caption = "Edit"
End If
End If
.Close
End With
Unload Me
Exit Sub
x:
frmtransferee.Image4.Picture = frmmain.ImageList1.ListImages(51).Picture
frmregular.Image4.Picture = frmmain.ImageList1.ListImages(51).Picture
frmfreshmen.Image4.Picture = frmmain.ImageList1.ListImages(51).Picture
Unload Me
End Sub


Private Sub Text1_Change()
ct = Combo1.ListIndex
If Combo1.ListIndex = 2 Then ct = 16
If Combo1.ListIndex = 3 Then ct = 10
On Error Resume Next
List.ListItems.Clear
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
If viewstudrec = "regular" Then
If userlevel = "Admin" Then
RS.Open "select * from tblstudents Where Enrolltype='Regular' And evaluated='No'" & sqlstate
Else
Label5.Caption = "Previous Advisory Class Section: " & prevyear
RS.Open "select * from tblstudents Where Enrolltype='Regular' And Status='Not Enrolled' And evaluated='No' and yrsec='" & prevyear & "'" & sqlstate
Call visibility
End If
Label14.Caption = "TNHS Enrollment System(Student's Record)-Regular Students"
ElseIf viewstudrec = "transferee" Then
RS.Open "select * from tblstudents Where Enrolltype='Transferee' and Status='Not Enrolled'" & sqlstate
Label14.Caption = "TNHS Enrollment System(Student's Record)-Transferee Students"
Call visibility
ElseIf viewstudrec = "studinfo" Then
Label14.Caption = "TNHS Enrollment System(Student's Record)"
RS.Open "select * from tblstudents Where Status='Enrolled'" & sqlstate
Call visibility
ElseIf viewstudrec = "freshmen" Then
RS.Open "select * from tblstudents Where Enrolltype='Freshmen' and Status='Not Enrolled'" & sqlstate
Label14.Caption = "TNHS Enrollment System(Student's Record)-Freshmen Students"
Call visibility
End If
With RS
Do Until .EOF
If LCase(Mid(.Fields(ct), 1, Len(Text1.Text))) = LCase(Text1.Text) Then
List.ListItems.Add , , .Fields(0)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(2) & " " & Mid(.Fields(3), 1, 1) & ". " & .Fields(1)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields("gender")
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(10)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(11)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(14)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(13)
If .Fields(13) = "Enrolled" Then
List.ListItems.Item(List.ListItems.Count).SmallIcon = ImageList1.ListImages(3).Index
ElseIf .Fields(13) = "Not Enrolled" Then
List.ListItems.Item(List.ListItems.Count).SmallIcon = ImageList1.ListImages(4).Index
Else
List.ListItems.Item(List.ListItems.Count).SmallIcon = ImageList1.ListImages(13).Index
End If
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
Sub regularload()
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblstudents Where studno='" & List.SelectedItem.Text & "'"
With RS
If .EOF = False Then
frmregular.txt1.Text = .Fields(0)
frmregular.txt2.Text = .Fields(1)
frmregular.txt3.Text = .Fields(2)
frmregular.txt4.Text = .Fields(3)
frmregular.txt5.Text = .Fields(4)
frmregular.txt6.Text = Format(.Fields(5), "MM/DD/YYYY")
frmregular.txt7.Text = Year(Date) - Year(.Fields(5))
frmregular.Combo1.Text = .Fields("gender")
frmregular.txt8.Text = .Fields(7)
frmregular.txt9.Text = .Fields(8)
frmregular.txt10.Text = .Fields(9)
frmregular.Text3.Text = .Fields(10)
frmregular.Text2.Text = .Fields(11)
If .Fields("goodmoral") = "Yes" Then frmregular.Check3.Value = 1
If .Fields("schoolcard") = "Yes" Then frmregular.Check1.Value = 1
If .Fields("bc") = "No" Then frmregular.Check2.Value = 1
If .Fields("form137") = "Yes" Then frmregular.Check4.Value = 1

frmregular.txt11.Text = Mid(.Fields(12), 5, 9)
If .Fields(19) <> "n/a" Then
frmregular.Text5.Text = .Fields(19)
End If
frmregular.Text1.Text = .Fields(13)
frmregular.Text4.Text = .Fields(18)
frmregular.Image4.Picture = LoadPicture(App.Path & "/students/" & .Fields(0) & ".jpg")
End If
.Close
End With
End Sub

