<%@ Page Language="VB" AutoEventWireup="false" CodeFile="inicio.aspx.vb" Inherits="inicio" culture="es-NI" meta:resourcekey="PageResource1" uiculture="auto" %>
<%@ Register src="Controles/Direccion.ascx" tagname="Direccion" tagprefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    <style type="text/css">
        .style1
        {
            width: 91px;
        }
        .style2
        {
            width: 91px;
            height: 23px;
        }
        .style3
        {
            height: 23px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <h1>
            <asp:Label ID="lblTitulo" runat="server" 
                Text="Bienvenido a Nuestra Tienda Virtual" 
                meta:resourcekey="lblTituloResource1"></asp:Label>
    </h1>
    
    <div>
    

    
        <table style="width:100%;">
            <tr>
                <td class="style1">
                    <asp:Label ID="lblNombre" runat="server" Text="Nombre" 
                        meta:resourcekey="lblNombreResource1"></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtNombre" runat="server" 
                        meta:resourcekey="txtNombreResource1"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="style2">
                    <asp:Label ID="lblTelefono" runat="server" Text="Telefono" 
                        meta:resourcekey="lblTelefonoResource1"></asp:Label>
                    </td>
                <td class="style3">
                    <asp:TextBox ID="txtNombre0" runat="server" 
                        meta:resourcekey="txtNombre0Resource1"></asp:TextBox>
                    </td>
                <td class="style3">
                    </td>
            </tr>
            <tr>
                <td class="style1">
                    &nbsp;</td>
                <td>
                    <asp:Button ID="btnIngresar" runat="server" style="height: 26px" 
                        Text="Ingresar" meta:resourcekey="btnIngresarResource1" />
                    <br />
                    <br />
                    <asp:Label ID="lblInfo" runat="server" Text="Label"></asp:Label>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
        </table>
    
   
    </div>
    <uc1:Direccion ID="Direccion1"  runat="server" />
    <br />
    <br />
    <uc1:Direccion ID="Direccion2" runat="server" />
    <br />
    <uc1:Direccion ID="Direccion3" runat="server" />
    </form>
</body>
</html>
