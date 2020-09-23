VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Texthandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function cLearMe(ByRef frm As Form) As String
Dim Form_Object As Control
For Each Form_Object In frm
    If TypeOf Form_Object Is TextBox Then
        Form_Object.Text = ""
    End If
Next Form_Object
End Function
