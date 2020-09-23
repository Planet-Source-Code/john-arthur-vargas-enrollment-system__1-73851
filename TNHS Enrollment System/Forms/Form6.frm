VERSION 5.00
Begin VB.Form frmprint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   840
      TabIndex        =   0
      Text            =   "Select"
      Top             =   1080
      Width           =   2775
   End
   Begin TNHSES.lvButtons_H Command4 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "Select"
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
      cGradient       =   8438015
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   33023
   End
   Begin TNHSES.lvButtons_H Command3 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
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
      cBhover         =   8438015
      LockHover       =   1
      cGradient       =   8438015
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   33023
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Print List of Students"
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
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Section"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   3360
      Picture         =   "Form6.frx":0000
      Top             =   60
      Width           =   420
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2700
      Left            =   15
      Top             =   15
      Width           =   3840
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   0
      Picture         =   "Form6.frx":2FAD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3900
   End
End
Attribute VB_Name = "frmprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim printsection As String, printsession As String, printtotalno As String
Dim printroom As String, printadviser As String, male As Integer, female As Integer
Dim males() As String, females() As String, ctr As Integer, ctr1 As Integer, ctr2 As Integer
Dim malestud() As String, femalestud() As String
Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Command4_Click()
On Error Resume Next
If Combo1.Text = "Select" Then
MsgBox "Select section first"
Else
printsection = Combo1.Text
If dreport = True Then

Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from tblsubjsched"
        Do Until RS.EOF
        If RS.Fields(0) = Combo1.Text Then
        printsession = RS.Fields(1) & " "
        printroom = RS.Fields(3) & " "
        printadviser = RS.Fields(4) & " "
        End If
        RS.MoveNext
        Loop
Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from tblstudents Where yrsec='" & Combo1.Text & "' And gender='Male' And status='Enrolled' Order by sname"
ReDim males(RS.RecordCount) As String
ReDim malestud(RS.RecordCount) As String
ctr = 0
Do Until RS.EOF
ctr = ctr + 1
males(ctr) = RS.Fields("sname") & ", " & RS.Fields("fname") & " " & Mid(RS.Fields("mname"), 1, 1) & "."
malestud(ctr) = RS.Fields(0)
RS.MoveNext
Loop
RS.Close
Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from tblstudents Where yrsec='" & Combo1.Text & "' And gender='Female' And status='Enrolled' Order by sname"
ReDim females(RS.RecordCount) As String
ReDim femalestud(RS.RecordCount) As String
ctr1 = 0
Do Until RS.EOF
ctr1 = ctr1 + 1
females(ctr1) = RS.Fields("sname") & ", " & RS.Fields("fname") & " " & Mid(RS.Fields("mname"), 1, 1) & "."
femalestud(ctr1) = RS.Fields(0)
RS.MoveNext
Loop
RS.Close
Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from tblprintstudreport"
If RS.EOF = False Then
Do Until RS.EOF
RS.Delete adAffectCurrent
RS.MoveNext
Loop
End If
If ctr1 < ctr Then
For ctr2 = 1 To ctr
RS.AddNew
RS.Fields(0) = males(ctr2)
RS.Fields(2) = malestud(ctr2)
If ctr2 <= ctr1 Then
RS.Fields(1) = females(ctr2)
RS.Fields(3) = femalestud(ctr2)
Else
RS.Fields(1) = " "
RS.Fields(3) = " "
End If
RS.Update
Next ctr2
ElseIf ctr1 > ctr Then
For ctr2 = 1 To ctr1
RS.AddNew
RS.Fields(1) = females(ctr2)
RS.Fields(3) = femalestud(ctr2)
If ctr2 <= ctr Then
RS.Fields(0) = males(ctr2)
RS.Fields(2) = malestud(ctr2)
Else
RS.Fields(0) = " "
RS.Fields(2) = " "
End If
RS.Update
Next ctr2
ElseIf ctr1 = ctr Then
For ctr2 = 1 To ctr1
RS.AddNew
RS.Fields(1) = females(ctr2)
RS.Fields(0) = males(ctr2)
RS.Fields(2) = malestud(ctr2)
RS.Fields(3) = femalestud(ctr2)
RS.Update
Next ctr2
End If
RS.Close
Dim rs6 As New ADODB.Recordset
rs6.Open "SELECT * From tblprintstudreport", Con
Set DataReport2.DataSource = rs6.DataSource
For Each obj In DataReport2.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rs6.DataMember
    End If
Next
DataReport2.Sections("Section2").Controls("lblyrsec").Caption = printsection
DataReport2.Sections("Section2").Controls("lblsession").Caption = printsession
DataReport2.Sections("Section2").Controls("lblroom").Caption = printroom
DataReport2.Sections("Section2").Controls("lbladviser").Caption = printadviser
DataReport2.Sections("Section2").Controls("lbltotalstud").Caption = ctr + ctr1
DataReport2.Sections("Section2").Controls("lblboy").Caption = "(" & ctr & ")"
DataReport2.Sections("Section2").Controls("lblgirl").Caption = "(" & ctr1 & ")"
DataReport2.Sections("Section1").Controls("Text1").DataField = "malestud"
DataReport2.Sections("Section1").Controls("Text2").DataField = "male"
DataReport2.Sections("Section1").Controls("Text3").DataField = "femalestud"
DataReport2.Sections("Section1").Controls("Text4").DataField = "female"
DataReport2.Refresh
DataReport2.Show 1
Set rs6 = Nothing
Else
Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = 3
    RS.ActiveConnection = Con
    RS.LockType = 2
        RS.Open "select * from tblsubjsched"
        Do Until RS.EOF
        If RS.Fields(0) = Combo1.Text Then
        If RS.Fields(2) = "Finished" Then
        printsession = RS.Fields(1) & " "
        printroom = RS.Fields(3) & " "
        printadviser = RS.Fields(4) & " "
    Else
    MsgBox "Selected section's schedule not set"
    Exit Sub
    End If
    End If
        RS.MoveNext
        Loop
Dim rs7 As New ADODB.Recordset
rs7.Open "SELECT * From tblsched Where room='" & printroom & "'", Con
Set DataReport1.DataSource = rs7.DataSource
For Each obj In DataReport1.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rs7.DataMember
    End If
Next
DataReport1.Sections("Section2").Controls("lblyrsec").Caption = printsection
DataReport1.Sections("Section2").Controls("lblsession").Caption = printsession
DataReport1.Sections("Section2").Controls("lblroom").Caption = printroom
DataReport1.Sections("Section2").Controls("lbladviser").Caption = printadviser
DataReport1.Sections("Section1").Controls("Text1").DataField = "class"
DataReport1.Sections("Section1").Controls("Text2").DataField = "time"
DataReport1.Sections("Section1").Controls("Text3").DataField = "subject"
DataReport1.Sections("Section1").Controls("Text4").DataField = "teacher"
Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from tblvalsched Where room='" & printroom & "'"
        DataReport1.Sections("Section3").Controls("lblday1").Caption = RS.Fields("day")
        DataReport1.Sections("Section3").Controls("lblclass1").Caption = RS.Fields(4)
        DataReport1.Sections("Section3").Controls("lbltime1").Caption = RS.Fields(5)
        DataReport1.Sections("Section3").Controls("lblsubs1").Caption = RS.Fields("subject")
        DataReport1.Sections("Section3").Controls("lbltsub1").Caption = RS.Fields("substeacher")
        DataReport1.Sections("Section3").Controls("lblteacher1").Caption = RS.Fields("teacher")
        RS.MoveNext
        DataReport1.Sections("Section3").Controls("lblday2").Caption = RS.Fields("day")
        DataReport1.Sections("Section3").Controls("lblclass2").Caption = RS.Fields(4)
        DataReport1.Sections("Section3").Controls("lbltime2").Caption = RS.Fields(5)
        DataReport1.Sections("Section3").Controls("lblsubs2").Caption = RS.Fields("subject")
        DataReport1.Sections("Section3").Controls("lbltsub2").Caption = RS.Fields("substeacher")
      RS.Close
DataReport1.Refresh
DataReport1.Show 1
Set rs7 = Nothing
End If
End If
End Sub

Private Sub Form_Load()
  Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from tblsubjsched"
Do Until RS.EOF
Combo1.AddItem RS.Fields(0)
RS.MoveNext
Loop
RS.Close
If dreport = False Then Label9.Caption = "Print Section's Schedule"
End Sub
Private Sub Image3_Click()
Unload Me
End Sub

