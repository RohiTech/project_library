Imports RSwebservice
Imports Microsoft.Reporting.WebForms

Partial Class ReportViewer
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Page.IsPostBack = False Then

            Dim rs As New ReportingService2005

            'rs.Credentials = System.Net.CredentialCache.DefaultCredentials
            rs.Credentials = New System.Net.NetworkCredential("Bernardo Robelo", "123")

            Dim Items As CatalogItem() = rs.ListChildren("/Reportes", False)

            For Each itm As CatalogItem In Items
                ddlReportes.Items.Add(itm.Name)
            Next
        End If


    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        ReportViewer1.ServerReport.ReportServerUrl = _
            New System.Uri("http://localhost/ReportServer/")

        ReportViewer1.ServerReport.ReportPath = _
        "/Reportes/" & ddlReportes.text

        ReportViewer1.ServerReport.Refresh()

        ReportViewer1.ShowParameterPrompts = False


        'Dim parametros As New List(Of Microsoft.Reporting.WebForms.ReportParameter)

        'parametros.Add _
        '(New Microsoft.Reporting.WebForms.ReportParameter("p_Pais", "Brazil"))

        'parametros.Add _
        '(New Microsoft.Reporting.WebForms.ReportParameter("p_Ciudad", "Sao Paulo"))

        If ddlReportes.text = "rptAgrupacion" Then

            Dim param(1) As  _
            Microsoft.Reporting.WebForms.ReportParameter

            param(0) = _
New Microsoft.Reporting.WebForms.ReportParameter("p_Pais", "Brazil")

            param(1) = _
New Microsoft.Reporting.WebForms.ReportParameter("p_Ciudad", "Sao Paulo")

            'ReportViewer1.ServerReport.SetParameters(parametros)
            ReportViewer1.ServerReport.SetParameters(param)

        End If



    End Sub
End Class
