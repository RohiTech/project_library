Attribute VB_Name = "Module1"
Public cn As New ADODB.Connection
Public cadena As String
Dim tipo As String
Public strServer As String
Public strDB As String
Public strUser As String
Public strpass As String
Public contador As Integer
Public banderas As Integer
Public Sub conectando()
On Error GoTo error
 cadena = "Provider=SQLOLEDB.1;" & _
"Password=" & frmconectar.txtcontraseña.Text & ";" & _
"Persist Security Info=true;" & _
"User Id=" & frmconectar.txtusuario.Text & ";" & _
"Initial Catalog=" & frmconectar.txtbasedatos.Text & ";" & _
"Data Source=" & frmconectar.txtservidor.Text

strServer = frmconectar.txtservidor.Text
 strDB = frmconectar.txtbasedatos.Text
strUser = frmconectar.txtusuario.Text
strpass = frmconectar.txtcontraseña.Text

cn.Open cadena
contador = 0
banderas = 1
error:
 If Err.Number <> 0 Then
MsgBox "No se conecto", vbCritical, "Error"
contador = contador + 1
banderas = 0
If contador = 3 Then
End
Else
End If
End If
End Sub



