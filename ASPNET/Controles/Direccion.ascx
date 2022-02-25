<%@ Control Language="VB" AutoEventWireup="false" CodeFile="Direccion.ascx.vb" Inherits="Direccion" %>
<table style="width: 33%;">
    <tr>
        <td style="background-color:Red;color:white">
            Direccion 1</td>
        <td style="background-color:Navy; color:Teal" >
            <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        </td>
    </tr>
    <tr>
        <td style="background-color:Red;color:white">
            Direccion 2</td>
        <td style="background-color:Navy; color:Teal" >
            <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
        </td>
    </tr>
    <tr>
        <td >
            &nbsp;</td>
        <td >
            <asp:Button ID="Button1" runat="server" Text="Button" />
        </td>
    </tr>
</table>
