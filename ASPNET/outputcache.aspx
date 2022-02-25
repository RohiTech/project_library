<%@ Page Language="VB" AutoEventWireup="false" CodeFile="outputcache.aspx.vb" Inherits="outputcache" %>
<%@ OutputCache Duration="30" VaryByParam="None" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <p>
        <br />
    </p>
    <form id="form1" runat="server">
    <p>
        <asp:Label ID="lblTiempo" runat="server" Text="Label"></asp:Label>
    </p>
    <p>
        &nbsp;</p>
    <p>
        Categoria<asp:TextBox ID="txtCategoria" runat="server"></asp:TextBox>
        <asp:Button ID="Button1" runat="server" Text="Buscar Por Categoria" />
    </p>
    <p>
        Precio
        <asp:TextBox ID="txtPrecio" runat="server"></asp:TextBox>
        <asp:Button ID="Button2" runat="server" Text="Buscar Por Precio y Categoria" />
    </p>
    <p>
        <asp:Button ID="Button3" runat="server" Text="Cache x Depencia" />
    </p>
    <div>
    
    </div>
    </form>
</body>
</html>
