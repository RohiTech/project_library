
Partial Class portal
    Inherits System.Web.UI.Page

    Protected Sub grid_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grid.SelectedIndexChanged

    End Sub

    Protected Sub SqlDataSource1_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs) Handles SqlDataSource1.Selecting

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Page.IsPostBack = False Then
            Dim rootItem As MenuItem = New MenuItem("Seleccionar Vista")
            For Each Modo As WebPartDisplayMode In WebPartManager1.DisplayModes
                rootItem.ChildItems.Add(New MenuItem(Modo.Name))
            Next
            Menu1.Items.Add(rootItem)
        End If
    End Sub

    Protected Sub Menu1_MenuItemClick(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.MenuEventArgs) Handles Menu1.MenuItemClick
        For Each Modo As WebPartDisplayMode In WebPartManager1.DisplayModes
            If Modo.Name = e.Item.Text Then
                WebPartManager1.DisplayMode = Modo
                Exit For
            End If
        Next
    End Sub




    Protected Sub Calendar1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.Load

        Dim webpart As GenericWebPart = Calendar1.Parent

        webpart.Title = "Calendario"

    End Sub



    Protected Sub AdRotator1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles AdRotator1.Load

        Dim webpart As GenericWebPart = AdRotator1.Parent

        webpart.Title = "Anuncios"
        webpart.AllowClose = False
        webpart.AllowMinimize = False
        webpart.AllowZoneChange = False
        webpart.ChromeType = PartChromeType.BorderOnly


    End Sub
End Class
