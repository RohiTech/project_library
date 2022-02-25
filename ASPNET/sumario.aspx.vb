
Partial Class sumario
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim Nombre As String = Request.QueryString("Nombre")
        Dim Prod As String = Request.QueryString("Prod")

        Label1.Text = _
        String.Format("Gracias, {0} por comprar {1} en nuestro sitio web", Nombre, Prod)

    End Sub
End Class
