Imports System.Configuration
Imports System.Web.Configuration

Partial Class encriptarcadena
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Conf As Configuration = WebConfigurationManager.OpenWebConfiguration("~")
        Dim section As ConfigurationSection = Conf.Sections("connectionStrings")
        section.SectionInformation.ProtectSection("DataProtectionConfigurationProvider")
        Conf.Save()
    End Sub
End Class
