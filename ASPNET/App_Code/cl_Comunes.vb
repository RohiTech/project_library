Imports Microsoft.VisualBasic
Imports System.Data.SqlClient


Public MustInherit Class ClComun
    Inherits ClConexion
    Dim vCodError As Integer
    Shared VMensajeError As String
    Dim vUsuario As String

    Public Enum TieneTransaccion
        Si = 1
        No = 0
    End Enum

    Public Property Usuario() As String
        Get
            Return vUsuario
        End Get
        Set(ByVal value As String)
            vUsuario = value
        End Set
    End Property

    Protected Sub InicializarMensajeError()
        vCodError = 0
        VMensajeError = ""
    End Sub

    Public Shared Property MensajeError() As String
        Get
            Return VMensajeError
        End Get
        Set(ByVal value As String)
            VMensajeError = value
        End Set
    End Property

    Public Property CodigoError() As Integer
        Get
            Return vCodError
        End Get
        Set(ByVal value As Integer)
            vCodError = value
        End Set
    End Property

    Public Function ObtenerParametrosWebConfig(ByVal Param As String) As String
        Dim cadena As String
        cadena = DAL.Conn
        Return (cadena)
    End Function

    Protected Sub SendEmail(ByVal MensajeError As String)
        Dim rf As Reflection.MethodInfo
        Dim NombreClase As String
        Dim flag As Boolean = True
        For Each rf In Me.GetType.GetMethods
            NombreClase = rf.DeclaringType.FullName
            If flag Then
                SendEmail(ObtenerParametrosWebConfig("EmailError"), NombreClase, MensajeError)
                flag = False
            End If
        Next
    End Sub

    Private Sub SendEmail(ByVal email As String, ByVal nombreclase As String, ByVal mensajeerror As String)

        Dim mail As Object
        Dim Mensaje As String
        Dim cdoConfig As Object
        Dim ip As String = ObtenerParametrosWebConfig("ServerCorreo")
        Dim emailfrom As String = ObtenerParametrosWebConfig("EmailError")
        Try

            cdoConfig = CreateObject("CDO.Configuration")

            With cdoConfig.Fields
                .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = ip
                .Update()
            End With
            mail = CreateObject("CDO.Message")
            mail.Configuration = cdoConfig
            mail.to = email
            mail.From = emailfrom

            mail.Subject = "Error en Sistema de Orders-IST-UPOLI"
            Mensaje = nombreclase & " <br> " & mensajeerror
            mail.htmlbody = Mensaje

            mail.Fields.Update()
            mail.Send()
            mail = Nothing
        Catch ex As Exception

        End Try
    End Sub

    Public Function ConfigurarComando(ByVal NombreProcedimiento As String, ByVal ListaValoresParametros() As String) As SqlCommand
        Dim cmd As New SqlCommand
        With cmd
            .Connection = Me.ConeccionSql
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = NombreProcedimiento
            Data.SqlClient.SqlCommandBuilder.DeriveParameters(cmd)
            .CommandTimeout = 0
        End With
        AsignarParametros(cmd, ListaValoresParametros)
        Return cmd
    End Function

    Public Function ConfigurarComando(ByVal Trans As TieneTransaccion, ByVal NombreProcedimiento As String, ByVal ListaValoresParametros() As String) As SqlCommand
        Dim cmd As New SqlCommand
        cmd = AsignaComando(NombreProcedimiento)
        If Trans = TieneTransaccion.Si Then
            cmd.Transaction = Me.BeginTranSql
        End If
        '' Recupera los parametros del sp
        Data.SqlClient.SqlCommandBuilder.DeriveParameters(cmd)
        ''asigna parametros
        AsignarParametros(cmd, ListaValoresParametros)
        Return cmd
    End Function
    Public Function ConfigurarComando(ByVal Trans As TieneTransaccion, ByVal NombreProcedimiento As String, ByVal ListaValoresParametros() As String, ByVal varOpcion As String) As SqlCommand
        Dim cmd As New SqlCommand
        cmd = AsignaComando(NombreProcedimiento)
        If Trans = TieneTransaccion.Si Then
            cmd.Transaction = Me.BeginTranSql
        End If

        '' Recupera los parametros del sp
        Data.SqlClient.SqlCommandBuilder.DeriveParameters(cmd)
        ''asigna parametros

        AsignarParametros(cmd, ListaValoresParametros, varOpcion)
        Return cmd
    End Function

    Private Function AsignaComando(ByVal NombreProcedimiento As String) As SqlCommand
        Dim cmd As New SqlCommand
        With cmd
            .Connection = Me.ConeccionSql
            .CommandType = Data.CommandType.StoredProcedure
            .CommandText = NombreProcedimiento
            .CommandTimeout = 0
        End With
        Return cmd
    End Function

    Private Sub AsignarParametros(ByRef cmd As SqlCommand, ByVal ListaValoresParametros() As String, ByVal varOpcion As String)
        Dim i As Integer
        SqlCommandBuilder.DeriveParameters(cmd)

        For i = 1 To ListaValoresParametros.Length - 1

            cmd.Parameters(i).Value = ListaValoresParametros(i - 1)
            If i = 1 Then
                cmd.Parameters(i).Direction = Data.ParameterDirection.InputOutput
            Else
                cmd.Parameters(i).Direction = Data.ParameterDirection.Input
            End If
        Next
    End Sub

    Private Sub AsignarParametros(ByRef cmd As SqlCommand, ByVal ListaValoresParametros() As String)
        Dim i As Integer
        SqlCommandBuilder.DeriveParameters(cmd)

        For i = 1 To ListaValoresParametros.Length - 1
            cmd.Parameters(i).Value = ListaValoresParametros(i - 1)
        Next
    End Sub

    Public Function ConfigurarComando(ByVal NombreProcedimiento As String) As SqlCommand
        Dim cmd As New SqlCommand
        cmd = AsignaComando(NombreProcedimiento)
        Return cmd
    End Function

    Public Function ConfigurarComando(ByVal Trans As TieneTransaccion, ByVal NombreProcedimiento As String) As SqlCommand
        Dim cmd As New SqlCommand
        cmd = AsignaComando(NombreProcedimiento)
        If Trans = TieneTransaccion.Si Then
            cmd.Transaction = Me.BeginTranSql
        End If
        Return cmd
    End Function
End Class


