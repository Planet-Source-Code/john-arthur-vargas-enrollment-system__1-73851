Attribute VB_Name = "Module1"
Public i_Clear As New Texthandling
Public Field_Check As New Texthandling
Public firstsec As Integer
Public secondsec As Integer
Public thirdsec As Integer
Public fourthsec As Integer
Public firstsec1 As Integer
Public secondsec1 As Integer
Public thirdsec1 As Integer
Public fourthsec1 As Integer
Public viewstudrec As String
Public sy As String
Public username As String
Public password As String
Public userid As String
Public teacherid(7) As String
Public valteacherid(2) As String
Public valsubteacherid(2) As String
Public dreport As Boolean
Public names As String
Public sname As String
Public fname As String
Public userlevel As String
Public usrsubject As String
Public adminbool As Boolean
Public address As String
Public number As String
Public teacherids As String
Public mapid As Integer
Public session As Integer
Public pid As String
Public indsched As Integer
Public days As String
Public classint As Integer
Public classes(2) As String
Public teacher As String
Public valint As Integer
Public boolpop As Integer
Public substeacher(2) As String
Public subssubject(2) As String
Public formboolean As Integer
Public prevyear As String
Public sessions As String
Public valteacher As String
Public valfloater As Boolean
Public viewteacherpopsession As String
Public viewteacherpopclass As String
Public Con As New ADODB.Connection
Public RS As New ADODB.Recordset
Public RS1 As New ADODB.Recordset
Public Function usrlog(logdesc As String)
Call dbConnection
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from userlog"
        With RS
.AddNew
.Fields(0) = userid
.Fields(1) = username
.Fields(2) = Date
.Fields(3) = Time
.Fields(4) = logdesc
.Fields(5) = Now
.Update
End With
End Function
Sub main()
frmsecurity.Show
End Sub

Sub loadyrsec()
frmsubsched.List.ListItems.Clear
Call dbConnection
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
RS.Open "select * from tblsubjsched"
With RS
Do Until .EOF
frmsubsched.List.ListItems.Add , , .Fields(0)
frmsubsched.List.ListItems.Item(frmsubsched.List.ListItems.Count).ListSubItems.Add , , .Fields(2)
If .Fields(2) = "Not Set" Then
frmsubsched.List.ListItems.Item(frmsubsched.List.ListItems.Count).SmallIcon = frmsubsched.imagelist1.ListImages(12).Index
Else
frmsubsched.List.ListItems.Item(frmsubsched.List.ListItems.Count).SmallIcon = frmsubsched.imagelist1.ListImages(11).Index
End If
.MoveNext
Loop
.Close
End With
End Sub
Public Sub dbConnection()
    Set Con = New ADODB.Connection
    Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\db1.mdb;Persist Security Info=False"
    Con.Open
End Sub
Public Sub displaysy()
Call dbConnection
Set RS = New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.ActiveConnection = Con
    RS.CursorLocation = 3
    RS.LockType = 2
        RS.Open "select * from tblyrsection"
        Do Until RS.EOF
        sy = Mid(RS.Fields(0), 1, 4)
        firstsec = RS.Fields(1)
        secondsec = RS.Fields(2)
        thirdsec = RS.Fields(3)
        fourthsec = RS.Fields(4)
            firstsec1 = RS.Fields(5)
        secondsec1 = RS.Fields(6)
        thirdsec1 = RS.Fields(7)
        fourthsec1 = RS.Fields(8)
        RS.MoveNext
        Loop
        RS.Close
End Sub

