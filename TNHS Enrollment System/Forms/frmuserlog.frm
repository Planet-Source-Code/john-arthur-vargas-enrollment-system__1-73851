VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmuserlog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form8"
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   LinkTopic       =   "Form8"
   ScaleHeight     =   6915
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   360
      Left            =   3240
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
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
      ItemData        =   "frmuserlog.frx":0000
      Left            =   1320
      List            =   "frmuserlog.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   120
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   8880
      Top             =   0
   End
   Begin MSComctlLib.ListView list 
      Height          =   3375
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilRecordIco"
      SmallIcons      =   "ilRecordIco"
      ColHdrIcons     =   "icoHeader"
      ForeColor       =   -2147483640
      BackColor       =   11793649
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   1
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
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Username"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Log"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ImageList icoHeader 
      Left            =   1920
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
            Picture         =   "frmuserlog.frx":003C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmuserlog.frx":05D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   2640
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
            Picture         =   "frmuserlog.frx":0B70
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TNHSES.lvButtons_H lvbuttons_H1 
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Print"
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
      Left            =   7320
      TabIndex        =   8
      Top             =   2760
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
      cGradient       =   12648447
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454143
   End
   Begin TNHSES.lvButtons_H lvButtons_H2 
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Delete"
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
   Begin VB.Label Label2 
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
      TabIndex        =   11
      Top             =   2280
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      Height          =   4650
      Index           =   1
      Left            =   120
      Top             =   2160
      Width           =   8415
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   120
      Picture         =   "frmuserlog.frx":110A
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   8415
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
      Left            =   240
      TabIndex        =   6
      Top             =   6600
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   8160
      Picture         =   "frmuserlog.frx":163A
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   6905
      Index           =   0
      Left            =   10
      Top             =   10
      Width           =   8645
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TNHS Enrollment System- User's Log"
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
      TabIndex        =   3
      Top             =   75
      Width           =   4215
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
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label8 
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
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmuserlog.frx":45E7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8715
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      Height          =   4650
      Index           =   2
      Left            =   120
      Top             =   2160
      Width           =   8415
   End
   Begin VB.Image Image2 
      Height          =   1800
      Left            =   -9360
      Picture         =   "frmuserlog.frx":859F
      Stretch         =   -1  'True
      Top             =   480
      Width           =   25800
   End
End
Attribute VB_Name = "frmuserlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filt As Integer


Private Sub Combo1_Click()
filt = Combo1.ListIndex
Call Text1_Change
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call urload
filt = 0
Combo1.ListIndex = 0
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

Sub urload()
Text1.Text = ""
List.ListItems.Clear
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from userlog order by datetime DESC"
With RS
Do Until .EOF
List.ListItems.Add , , .Fields(0)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(1)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(2)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(3)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(4)
List.ListItems.Item(List.ListItems.Count).SmallIcon = ilRecordIco.ListImages(1).Index
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



Private Sub lvButtons_H1_Click()
If List.ListItems.Count = 0 Then
MsgBox "No Log to print"
Else
If MsgBox("Print this Log?", vbYesNo) = vbYes Then
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from printlog order by dbdate DESC"
With RS
Do Until .EOF
.Delete adAffectCurrent
.MoveNext
Loop
Dim b As Integer
b = List.ListItems.Count
Do Until b = 0
RS.AddNew
.Fields(0) = List.ListItems.Item(b).Text
.Fields(1) = List.ListItems.Item(b).SubItems(1)
.Fields(2) = List.ListItems.Item(b).SubItems(2)
.Fields(3) = List.ListItems.Item(b).SubItems(3)
.Fields(4) = List.ListItems.Item(b).SubItems(4)
RS.Update
b = b - 1
Loop
.Close
End With
Set RS = New ADODB.Recordset
Dim rs7 As New ADODB.Recordset
rs7.Open "SELECT * From printlog Order by dbdate DESC", Con
Set DataReport3.DataSource = rs7.DataSource
For Each obj In DataReport3.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rs7.DataMember
    End If
Next
DataReport3.Sections("Section2").Controls("lblyrsec").Caption = List.ListItems.Count
DataReport3.Sections("Section1").Controls("Text1").DataField = "UserID"
DataReport3.Sections("Section1").Controls("Text2").DataField = "Username"
DataReport3.Sections("Section1").Controls("Text3").DataField = "dbdate"
DataReport3.Sections("Section1").Controls("Text4").DataField = "dbtime"
DataReport3.Sections("Section1").Controls("Text5").DataField = "dblog"
DataReport3.Refresh
DataReport3.Show 1
End If
End If
End Sub

Private Sub lvButtons_H2_Click()
If List.ListItems.Count = 0 Then
MsgBox "Theres no log to delete"
Else
If MsgBox("Are you sure you want to delete logs in the listview?", vbYesNo + vbQuestion) = vbYes Then
Dim b As Integer
b = List.ListItems.Count
Do Until b = 0
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from userlog"
 Do Until RS.EOF
If RS.Fields(0) = List.ListItems.Item(b).Text And RS.Fields(1) = List.ListItems.Item(b).ListSubItems(1) And RS.Fields(2) = List.ListItems.Item(b).ListSubItems(2) And RS.Fields(3) = List.ListItems.Item(b).ListSubItems(3) And RS.Fields(4) = List.ListItems.Item(b).ListSubItems(4) Then RS.Delete adAffectCurrent
RS.MoveNext
Loop
RS.Close
b = b - 1
Loop
MsgBox "Selected logs has been deleted"
Call Text1_Change
End If
End If
End Sub


Private Sub Text1_Change()
On Error Resume Next
List.ListItems.Clear
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from userlog Order by datetime DESC"
With RS
Do Until .EOF
If LCase(Mid(.Fields(filt), 1, Len(Text1.Text))) = LCase(Text1.Text) Then
List.ListItems.Add , , .Fields(0)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(1)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(2)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(3)
List.ListItems.Item(List.ListItems.Count).ListSubItems.Add , , .Fields(4)
List.ListItems.Item(List.ListItems.Count).SmallIcon = ilRecordIco.ListImages(1).Index
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
If frmuserlog.Width >= 8645 Then
Timer1.Enabled = False
Else
frmuserlog.Left = frmuserlog.Left - 250
frmuserlog.Width = frmuserlog.Width + 500
End If
End Sub

