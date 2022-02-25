<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Guardar_Categorias.aspx.vb" Inherits="Guardar_Categorias" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
    </div>
    Nombre&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:TextBox ID="txtNombre" runat="server"></asp:TextBox>
    <br />
    Descripcion
    <asp:TextBox ID="txtDescripcion" runat="server"></asp:TextBox>
    <br />
    <asp:Button ID="Button1" runat="server" Text="Guardar" />
    </form>
</body>
</html>
