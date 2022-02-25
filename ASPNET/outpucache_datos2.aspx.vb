
Partial Class outpucache_datos2
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Label1.Text = Date.Now.ToString

        Dim cat As String = Request.QueryString("Cat")
        Dim Precio As String = Request.QueryString("Precio")

        GridView1.DataSource = DAL.Leer("select * from products where categoryID='" & cat & "' AND UnitPrice<" & Precio)
        GridView1.DataBind()

    End Sub
End Class
