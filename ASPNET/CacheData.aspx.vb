Imports System.Data

Partial Class CacheData
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim dt As New DataTable

        Dim key As String = txtCat.Text

        If Cache(key) Is Nothing Then

            dt = DAL.Leer("select * from products where categoryid = " & txtCat.Text)
            Cache(key) = dt

        Else

            dt = Cache(key)

        End If


        GridView1.DataSource = dt
        GridView1.DataBind()


    End Sub
End Class
