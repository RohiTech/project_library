Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient

Public Class DAL

    Public Shared Conn As String = ConfigurationManager.ConnectionStrings("NorthwindConnectionString").ConnectionString

    Shared Function Leer(ByVal Comando As String) As DataTable
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(Comando, Conn)
        da.Fill(dt)
        Return dt
    End Function

    Shared Sub Escribir(ByVal Comando As String)
        Dim Connect As New SqlConnection(Conn)
        Dim cmd As New SqlCommand(Comando, Connect)
        Connect.Open()
        cmd.ExecuteNonQuery()
        Connect.Close()
    End Sub

End Class
