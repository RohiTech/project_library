
Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        lblMensaje.Text = _
        String.Format("Bienvenido {0} al Curso ASP.NET", txtNombre.Text)

    End Sub

    Protected Sub dlDepartamentos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dlDepartamentos.SelectedIndexChanged
        lblMensaje.Text = _
        String.Format("Gracias, {0} su departamento es {1}", txtNombre.Text, dlDepartamentos.SelectedItem.Text)
    End Sub


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim b As Double

        b = "asassa"

    End Sub
End Class
