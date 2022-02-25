<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Orden.aspx.vb" Inherits="Orden" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    <style type="text/css">
        .style1
        {
            width: 67px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <h1>
        Seleccione su Producto
            <asp:Label ID="lblNombre" runat="server" Text="Label"></asp:Label>
        </h1>
    
        <table style="width:100%;">
            <tr>
                <td class="style1">
                    Producto</td>
                <td>
                    <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="style1">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="style1">
                    &nbsp;</td>
                <td>
                    <asp:Button ID="Button1" runat="server" style="height: 26px" Text="Solicitar" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
