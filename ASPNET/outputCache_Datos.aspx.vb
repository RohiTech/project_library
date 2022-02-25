
Partial Class outputCache_Datos
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim cat As String = Request.QueryString("Cat")

        Label1.Text = Date.Now.ToString

        GridView1.DataSource = DAL.Leer("select * from products where categoryID=" & cat)
        GridView1.DataBind()



    End Sub


    Shared Function MostrarInfo(ByVal Context As HttpContext) As String
        Return Date.Now.ToString
    End Function

End Class
