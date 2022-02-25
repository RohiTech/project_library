<%@ Page Language="VB" Trace="true" AutoEventWireup="false" CodeFile="ViewState.aspx.vb" Inherits="ViewState" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    <style type="text/css">
        .style1
        {
            width: 94px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:ListBox ID="ListBox1" runat="server" EnableViewState="False" 
            SelectionMode="Multiple">
            <asp:ListItem>Nicaragua</asp:ListItem>
            <asp:ListItem>Honduras</asp:ListItem>
            <asp:ListItem>Costa Rica</asp:ListItem>
            <asp:ListItem>USA</asp:ListItem>
        </asp:ListBox>
    
    </div>
    <asp:Button ID="Button1" runat="server" Text="Button" Width="56px" />
    <br />
    <br />
    <asp:Label ID="lblMensaje" runat="server"></asp:Label>
    <br />
    <br />
    <br />
    <table style="width:100%;">
        <tr>
            <td class="style1">
    <asp:Label ID="Label1" runat="server" Text="Nombre"></asp:Label>
            </td>
            <td>
            <asp:TextBox ID="txtNombre" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td class="style1">
                <asp:Label ID="Label2" runat="server" Text="De"></asp:Label>
            </td>
            <td>
            <asp:TextBox ID="txtDe" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td class="style1">
                <asp:Label ID="lblHasta" runat="server" Text="Hasta"></asp:Label>
            </td>
            <td>
                <asp:TextBox ID="txtHasta" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
        <td class="style1">
        
            &nbsp;</td>
        <td>
        
            <asp:Button ID="Button2" runat="server" Text="Copiar" Width="54px" />
            <asp:Button ID="Button3" runat="server" Height="26px" Text="Pegar" />
        
        </td>
        </tr>
    </table>
    <br />
        </form>
</body>
</html>
