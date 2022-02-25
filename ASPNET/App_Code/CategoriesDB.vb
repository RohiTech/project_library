Imports System.Data.SqlClient

Namespace NombreSistema.Transacciones
    Public Class ClCategories
        Inherits ClComun
        Dim vCategoriesTabla As New NombreSistema.Estructura.Categories
        '-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- 
        Public Property CategoriesTabla() As NombreSistema.Estructura.Categories
            Get
                Return vCategoriesTabla
            End Get
            Set(ByVal value As NombreSistema.Estructura.Categories)
                vCategoriesTabla = value
            End Set
        End Property

        '-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
        Public Sub ins_Categories(ByVal Categories As NombreSistema.Estructura.Categories)
            Dim ObjParameter(4) As String
            ObjParameter(0) = Categories.CategoryID
            ObjParameter(1) = Categories.CategoryName
            ObjParameter(2) = Categories.Description
            ObjParameter(3) = Categories.Picture

            Try
                Me.InicializarMensajeError()
                Me.OpenSqlBD()
                vCategoriesTabla.CategoryID = Me.ConfigurarComando(TieneTransaccion.Si, "sp_ins_Categories", ObjParameter).ExecuteNonQuery()
                Me.Commit()
            Catch ex As Exception
                Me.CodigoError = -1
                Me.Rollback()
                Me.MensajeError = ex.Message
                Me.SendEmail(ex.Message)
            Finally
                Me.CloseSqlBD()
            End Try
        End Sub


        '-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
        Public Sub upd_Categories(ByVal Categories As NombreSistema.Estructura.Categories)
            Dim ObjParameter(4) As String
            ObjParameter(0) = Categories.CategoryID
            ObjParameter(1) = Categories.CategoryName
            ObjParameter(2) = Categories.Description
            ObjParameter(3) = Categories.Picture

            Try
                Me.InicializarMensajeError()
                Me.OpenSqlBD()
                vCategoriesTabla.CategoryID = Me.ConfigurarComando(TieneTransaccion.Si, "sp_upd_Categories", ObjParameter).ExecuteNonQuery()
                Me.Commit()
            Catch ex As Exception
                Me.CodigoError = -1
                Me.Rollback()
                Me.MensajeError = ex.Message
                Me.SendEmail(ex.Message)
            Finally
                Me.CloseSqlBD()
            End Try
        End Sub


        '-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
        Public Function sel_Categories() As Data.DataTable
            Dim dt As New Data.DataTable
            Try
                Me.InicializarMensajeError()
                Me.OpenSqlBD()
                dt.Load(Me.ConfigurarComando("sp_sel_Categories").ExecuteReader)
            Catch ex As Exception
                Me.CodigoError = -1
                Me.MensajeError = ex.Message
                Me.SendEmail(ex.Message)
            Finally
                Me.CloseSqlBD()
            End Try
            Return dt
        End Function

        '-- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- -- --
        Public Sub del_Categories(ByVal Categories As NombreSistema.Estructura.Categories)
            Dim ObjParameter(1) As String
            ObjParameter(0) = Categories.CategoryID
            Try
                Me.InicializarMensajeError()
                Me.OpenSqlBD()
                Me.ConfigurarComando(TieneTransaccion.Si, "sp_del_Categories", ObjParameter).ExecuteNonQuery()
                Me.Commit()
            Catch ex As Exception
                Me.CodigoError = -1
                Me.Rollback()
                Me.MensajeError = ex.Message
                Me.SendEmail(ex.Message)
            Finally
                Me.CloseSqlBD()
            End Try
        End Sub
    End Class
End Namespace
