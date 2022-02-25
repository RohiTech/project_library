
Partial Class Guardar_Categorias
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim Cat As New NombreSistema.Estructura.Categories
        Dim CatDB As New NombreSistema.Transacciones.ClCategories

        Cat.CategoryName = txtNombre.Text
        Cat.Description = txtDescripcion.Text
        Cat.Picture = 0

        CatDB.ins_Categories(Cat)

        Response.Redirect("categorias.aspx")

    End Sub
End Class
