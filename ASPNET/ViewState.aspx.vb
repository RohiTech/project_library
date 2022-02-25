
Partial Class ViewState
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        For Each c As ListItem In ListBox1.Items

            If c.Selected Then

                lblMensaje.Text &= c.Text & "<br/>"

            End If

        Next
    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        ViewState.Add("de", txtDe.Text)
        ViewState("Nombre") = txtNombre.Text
        ViewState("hasta") = txtHasta.Text
    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        txtDe.Text = ViewState("de").ToString
        txtHasta.Text = ViewState("hasta").ToString
        txtNombre.Text = ViewState("Nombre").ToString
    End Sub
End Class
