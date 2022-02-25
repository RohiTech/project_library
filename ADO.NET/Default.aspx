<%@ Page Language="C#" AutoEventWireup="true"  CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>ADO.NET con C#</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp;
        <asp:DetailsView ID="DetailsView1" runat="server" AllowPaging="True" BackColor="White"
            BorderColor="#DEDFDE" BorderStyle="None" BorderWidth="1px" CellPadding="4" DataKeyNames="ID"
            DataSourceID="ObjectDataSource1" ForeColor="Black" GridLines="Vertical" Height="50px"
            Width="125px">
            <FooterStyle BackColor="#CCCC99" />
            <EditRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
            <RowStyle BackColor="#F7F7DE" />
            <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
            <Fields>
                <asp:CommandField ShowDeleteButton="True" ShowEditButton="True" ShowInsertButton="True" />
            </Fields>
            <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
            <HeaderTemplate>
                Tabla Usuarios:
            </HeaderTemplate>
            <AlternatingRowStyle BackColor="White" />
        </asp:DetailsView>
        <asp:ObjectDataSource ID="ObjectDataSource1" runat="server" DeleteMethod="eliminarDato"
            InsertMethod="insertarDato" SelectMethod="SeleccionarDatos" TypeName="DALUsuario"
            UpdateMethod="actualizarDatos">
            <DeleteParameters>
                <asp:Parameter Name="id" Type="Int32" />
            </DeleteParameters>
            <UpdateParameters>
                <asp:Parameter Name="id" Type="Int32" />
                <asp:Parameter Name="nombre" Type="String" />
                <asp:Parameter Name="apellidos" Type="String" />
                <asp:Parameter Name="email" Type="String" />
            </UpdateParameters>
            <InsertParameters>
                <asp:Parameter Name="id" Type="Int32" />
                <asp:Parameter Name="nombre" Type="String" />
                <asp:Parameter Name="apellidos" Type="String" />
                <asp:Parameter Name="email" Type="String" />
            </InsertParameters>
        </asp:ObjectDataSource>

    </div>
    </form>
</body>
</html>
