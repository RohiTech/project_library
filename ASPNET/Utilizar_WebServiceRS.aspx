<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Utilizar_WebServiceRS.aspx.vb" Inherits="Utilizar_WebServiceRS" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:DropDownList ID="ddlReportes" runat="server">
        </asp:DropDownList>
        <br />
        <br />
        <asp:DropDownList ID="ddlFormatos" runat="server">
        </asp:DropDownList>
    
        <br />
        <br />
        <asp:Button ID="Button1" runat="server" Text="Mostrar" />
    
    </div>
    </form>
</body>
</html>
