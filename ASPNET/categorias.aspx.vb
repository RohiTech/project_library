
Partial Class categorias
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        lblTiempo.Text = DateTime.Now.ToString

        GridView1.DataSource = DAL.Leer("select * from categories")
        GridView1.DataBind()

    End Sub
End Class
