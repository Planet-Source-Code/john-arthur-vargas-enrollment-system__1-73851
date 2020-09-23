VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmeditinfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Enrollment System(Manage Users)"
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8520
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
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
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3840
      Width           =   2655
   End
   Begin VB.TextBox Text2 
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox Text1 
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4440
      Width           =   2655
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
      Left            =   1920
      MaxLength       =   9
      TabIndex        =   5
      Top             =   3000
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
      Left            =   1440
      TabIndex        =   3
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   7920
      Top             =   360
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
      Left            =   1440
      TabIndex        =   2
      Top             =   3480
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
      Left            =   1440
      TabIndex        =   4
      Top             =   2520
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
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2040
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1080
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
      Left            =   1440
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
   End
   Begin TNHSES.lvButtons_H Command3 
      Height          =   375
      Left            =   7080
      TabIndex        =   20
      Top             =   5040
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
   Begin TNHSES.lvButtons_H Command1 
      Height          =   375
      Left            =   5760
      TabIndex        =   21
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Save"
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
      Left            =   5400
      TabIndex        =   22
      Top             =   3360
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
      LockHover       =   1
      cGradient       =   12648447
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454143
   End
   Begin TNHSES.lvButtons_H lvButtons_H3 
      Height          =   375
      Left            =   6720
      TabIndex        =   23
      Top             =   3360
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
      cBhover         =   12648447
      LockHover       =   1
      cGradient       =   12648447
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454143
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   0
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
      Height          =   5655
      Left            =   10
      Top             =   15
      Width           =   8505
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   120
      Left            =   0
      Top             =   5520
      Width           =   8535
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   115
      Left            =   0
      Top             =   360
      Width           =   8535
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
      Left            =   240
      TabIndex        =   26
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject to Teach"
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
      Left            =   4320
      TabIndex        =   25
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2415
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
      Height          =   615
      Left            =   240
      TabIndex        =   19
      Top             =   4920
      Width           =   1095
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
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   4560
      Width           =   2055
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
      Left            =   1440
      TabIndex        =   15
      Top             =   3000
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   8040
      Picture         =   "frmeditinfo.frx":0000
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
      Left            =   240
      TabIndex        =   14
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TNHS Enrollment System- Edit Information"
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
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmeditinfo.frx":2FAD
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
      Left            =   240
      TabIndex        =   12
      Top             =   3120
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
      Left            =   240
      TabIndex        =   11
      Top             =   3600
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
      Left            =   240
      TabIndex        =   10
      Top             =   2640
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
      Left            =   240
      TabIndex        =   9
      Top             =   2160
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
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   2055
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
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   5685
      Left            =   4155
      Top             =   240
      Width           =   75
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00400000&
      Height          =   5085
      Left            =   120
      Top             =   480
      Width           =   8295
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   120
      Picture         =   "frmeditinfo.frx":6F65
      Stretch         =   -1  'True
      Top             =   480
      Width           =   8295
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      Height          =   5085
      Left            =   120
      Top             =   480
      Width           =   8295
   End
End
Attribute VB_Name = "frmeditinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()


    If Len(txtid.Text) = 0 Or Len(txtuser.Text) = 0 Or Len(txtpass.Text) = 0 Or Len(txtsname.Text) = 0 Or Len(txtfname.Text) = 0 Then
    MsgBox "Username, Password, User Level, Surname and Firstname are Required fields"
    Else
    
Call updaterecord
    End If
    End Sub



Private Sub Command3_Click()
Timer2.Enabled = True
End Sub


Private Sub Form_Load()
Text1.Text = userlevel
Text2.Text = usersubject
Me.Width = 20
Call dbConnection
Call loadusr
End Sub
Sub loadusr()
On Error Resume Next
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
    RS.Open "Select * From users"
    RS.Find ("ID='" & userid & "'")
    If RS.EOF = False Then
txtid.Text = RS.Fields(0)
txtuser.Text = RS.Fields(1)
txtpass.Text = RS.Fields(2)
txtsname.Text = RS.Fields(3)
txtfname.Text = RS.Fields(4)
txtaddress.Text = RS.Fields(7)
txtcnumber.Text = Mid(RS.Fields(8), 5, 9)
Text1.Text = RS.Fields(5)
Text2.Text = RS.Fields("yrtomodify")
Text3.Text = RS.Fields("Subject")
Image4.Picture = LoadPicture(App.Path & "/usrpic/" & txtid.Text & ".jpg")
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Timer2.Enabled = True
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

Private Sub lvButtons_H2_Click()
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

Private Sub lvButtons_H3_Click()
viewstudrec = "user1"
frmwebcam.Show 1
End Sub
Private Sub Timer1_Timer()
If Me.Width = 8520 Then
Timer1.Enabled = False
Else
Me.Left = Me.Left - 250
Me.Width = Me.Width + 500
End If
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
If Me.Width <= 395 Then
Timer2.Enabled = False
Unload Me
Else
Me.Left = Me.Left + 250
Me.Width = Me.Width - 500
End If
End Sub

Private Sub txtcnumber_KeyPress(KeyAscii As Integer)
 If (Not IsNumeric(Chr$(KeyAscii)) And (Chr$(KeyAscii) <> vbBack)) Then KeyAscii = 0
End Sub
Sub updaterecord()
If txtid.Text = Clear Or txtuser.Text = Clear Or txtpass.Text = Clear Or txtsname.Text = Clear Or txtfname.Text = Clear Then
MsgBox "Complete All fields"
Else

    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from users Where ID='" & txtid.Text & "'"
        With RS
            .Fields(0) = txtid.Text
            .Fields(1) = txtuser.Text
            .Fields(2) = txtpass.Text
            .Fields(3) = txtsname.Text
            .Fields(4) = txtfname.Text
            .Fields(7) = txtaddress.Text
      .Fields("status") = "active"
      If Len(txtcnumber.Text) = 9 Then
            .Fields(8) = "+639" & txtcnumber.Text
            End If
            .Update
            .Close
        End With
    MsgBox "Record Updated"
    SavePicture Image4, App.Path & "/usrpic/" & txtid.Text & ".jpg"
    
    userid = txtid.Text
username = txtuser.Text
password = txtpass.Text
fname = txtfname.Text
sname = txtsname.Text
names = fname & " " & sname
address = txtaddress.Text & " "
number = "+639" & txtcnumber.Text & " "
      usrlog ("Edited User Information")
    End If
End Sub
Private Sub txtuser_LostFocus()
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from users"
        Do Until RS.EOF
        If txtuser.Text <> username Then
        If txtuser.Text = RS.Fields(1) Then
        
        MsgBox "Username Not available"
        txtuser.Text = Clear
        txtuser.SetFocus
        End If
        End If
        RS.MoveNext
        Loop
        RS.Close
End Sub
