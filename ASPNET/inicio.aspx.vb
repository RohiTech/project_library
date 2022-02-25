
Partial Class inicio
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnIngresar.Click
        Response.Redirect("orden.aspx?Nombre=" & Encriptar.Encrypt(txtNombre.Text, "Financia"))
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        lblInfo.Text = String.Format("{0}, {1:C2},  {2:N}", _
                                     Today.Date.ToString, 12342.46, 12345678.9)


        'lblNombre.Text = Resources.Resource.Nombre
        'lblTelefono.Text = Resources.Resource.Telefono
        'lblTitulo.Text = Resources.Resource.Titulo
        'btnIngresar.Text = Resources.Resource.btnEnviar

    End Sub
End Class
