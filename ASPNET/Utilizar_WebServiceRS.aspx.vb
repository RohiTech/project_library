Imports RSwebservice

Partial Class Utilizar_WebServiceRS
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rs As New ReportingService2005

        'rs.Credentials = System.Net.CredentialCache.DefaultCredentials
        rs.Credentials = New System.Net.NetworkCredential("Bernardo Robelo", "narco")

        Dim Items As CatalogItem() = _
        rs.ListChildren(ConfigurationManager.AppSettings("Proyecto").ToString, False)

        For Each itm As CatalogItem In Items
            ddlReportes.Items.Add(itm.Name)
        Next

        Dim Formats As Extension() = _
        rs.ListExtensions(ExtensionTypeEnum.Render)

        For Each itm As Extension In Formats
            ddlFormatos.Items.Add(itm.Name)
        Next


    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim strCadena As String

        strCadena = ConfigurationManager.AppSettings("Servidor").ToString _
        & ConfigurationManager.AppSettings("Carpeta").ToString _
        & ddlReportes.SelectedValue & "&rs:Format=" & ddlFormatos.SelectedValue

        Response.Redirect(strCadena)


    End Sub
End Class
