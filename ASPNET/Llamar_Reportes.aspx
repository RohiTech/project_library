<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Llamar_Reportes.aspx.vb" Inherits="Llamar_Reportes" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:DropDownList ID="DropDownList1" runat="server">
            <asp:ListItem Value="rptProductos">Reporte de Productos</asp:ListItem>
            <asp:ListItem Value="rptClientes">Reporte de Clientes</asp:ListItem>
            <asp:ListItem Value="rptAgrupacion">Reporte de Agrupacion</asp:ListItem>
        </asp:DropDownList>
    
    </div>
    <asp:Button ID="Button1" runat="server" Text="Ver Reportes" />
    </form>
</body>
</html>
