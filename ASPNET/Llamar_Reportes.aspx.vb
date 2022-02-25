
Partial Class Llamar_Reportes
    Inherits System.Web.UI.Page

    Dim strCadena As String

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        
        '   If DropDownList1.SelectedValue = "rptAgrupacion" Then

        strCadena = ConfigurationManager.AppSettings("Servidor").ToString _
        & ConfigurationManager.AppSettings("Carpeta").ToString _
        & DropDownList1.SelectedValue & "&rs:Format=CSV&p_Pais=Brazil&p_Ciudad=Sao Paulo"

        '  Else




        '   End If

        Response.Redirect(strCadena)


    End Sub
End Class
