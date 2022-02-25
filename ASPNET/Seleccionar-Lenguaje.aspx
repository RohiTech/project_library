<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Seleccionar-Lenguaje.aspx.vb" Inherits="Seleccionar_Lenguaje" culture="auto" meta:resourcekey="PageResource1" uiculture="auto" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Registro</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:Label ID="lblIdioma" runat="server" Text="Seleccionar Idioma" 
            meta:resourcekey="lblIdiomaResource1"></asp:Label>
           <asp:DropDownList ID="ddlLan" runat="server" AutoPostBack="True" 
            Height="16px" meta:resourcekey="ddlLanResource1">
               <asp:ListItem meta:resourcekey="ListItemResource1">es-NI</asp:ListItem>
               <asp:ListItem meta:resourcekey="ListItemResource2">en-US</asp:ListItem>
        </asp:DropDownList>
    
    <br /><br />
    
        <asp:Label ID="lblNombre" runat="server" Text="Nombre" 
            meta:resourcekey="lblNombreResource1"></asp:Label>
        <asp:TextBox ID="TextBox1" runat="server" Width="118px" 
            meta:resourcekey="TextBox1Resource1"></asp:TextBox>
        <br />
        <asp:Label ID="lblDireccion" runat="server" Text="Direccion" 
            meta:resourcekey="lblDireccionResource1"></asp:Label>
        <asp:TextBox ID="TextBox2" runat="server" Width="118px" 
            meta:resourcekey="TextBox2Resource1"></asp:TextBox>
        <br />
        <asp:Label ID="lblTelefono" runat="server" Text="Telefono" 
            meta:resourcekey="lblTelefonoResource1"></asp:Label>
        <asp:TextBox ID="TextBox3" runat="server" Width="118px" 
            meta:resourcekey="TextBox3Resource1"></asp:TextBox>
        <br />
        <asp:Label ID="lblCedula" runat="server" Text="Cedula" 
            meta:resourcekey="lblCedulaResource1"></asp:Label>
        <asp:TextBox ID="TextBox4" runat="server" Width="118px" 
            meta:resourcekey="TextBox4Resource1"></asp:TextBox>
    
    </div>
    <asp:Button ID="btnGuardar" runat="server" Text="Guardar" 
        meta:resourcekey="btnGuardarResource1" />
    </form>
</body>
</html>
