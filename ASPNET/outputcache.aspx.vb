
Partial Class outputcache
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        lblTiempo.Text = Date.Now.ToString





    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Response.Redirect("outputcache_Datos.aspx?Cat=" & txtCategoria.Text)
    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Response.Redirect("outpucache_Datos2.aspx?Cat=" & txtCategoria.Text & "&Precio=" & txtPrecio.Text)
    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Response.Redirect("categorias.aspx")
    End Sub
End Class
