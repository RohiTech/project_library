
Partial Class Orden
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim Nombre As String
        Nombre = Encriptar.Decrypt(Request.QueryString("Nombre"), "Financia")
        lblNombre.Text = Nombre

    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Response.Redirect("sumario.aspx?Nombre=" & Request.QueryString("Nombre") & "&Prod=" & TextBox1.Text)
    End Sub
End Class
