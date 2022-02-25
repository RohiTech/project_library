
Partial Class Direccion
    Inherits System.Web.UI.UserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub


    Public Property Dire1() As String
        Get
            Return TextBox2.Text
        End Get
        Set(ByVal value As String)
            TextBox2.Text = value
        End Set
    End Property

    Public Property Dire2() As String
        Get
            Return TextBox1.Text
        End Get
        Set(ByVal value As String)
            TextBox1.Text = value
        End Set
    End Property

End Class
