<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Pagina de Inicio</title>
</head>
<body>
    <form id="form1" runat="server">

    <h1>
    <asp:Label ID="lblTitulo" runat="server" Text="Bienvenidos al Curso ASP.NET !!!"></asp:Label>
    </h1>
    
    <div>
        <asp:SiteMapDataSource ID="SiteMapDataSource1" runat="server" />
        <asp:Menu ID="Menu1" runat="server" BackColor="#FFFBD6" 
            DataSourceID="SiteMapDataSource1" DynamicHorizontalOffset="2" 
            Font-Names="Verdana" Font-Size="0.8em" ForeColor="#990000" 
            Orientation="Horizontal" PathSeparator="-" StaticDisplayLevels="2" 
            StaticSubMenuIndent="10px">
            <StaticSelectedStyle BackColor="#FFCC66" ForeColor="White" />
            <StaticMenuItemStyle HorizontalPadding="5px" VerticalPadding="2px" />
            <DynamicHoverStyle BackColor="#990000" ForeColor="White" />
            <DynamicMenuStyle BackColor="#FFFBD6" />
            <DynamicItemTemplate>
                <%# Eval("Text") %>
            </DynamicItemTemplate>
            <DynamicSelectedStyle BackColor="#FFCC66" ForeColor="White" />
            <DynamicMenuItemStyle HorizontalPadding="5px" VerticalPadding="2px" />
            <StaticHoverStyle BackColor="#990000" ForeColor="White" />
        </asp:Menu>
        <br />
        <asp:TreeView ID="TreeView1" runat="server" DataSourceID="SiteMapDataSource1" 
            ImageSet="Contacts" NodeIndent="10">
            <ParentNodeStyle Font-Bold="True" ForeColor="#5555DD" />
            <HoverNodeStyle Font-Underline="False" />
            <SelectedNodeStyle Font-Underline="True" HorizontalPadding="0px" 
                VerticalPadding="0px" />
            <NodeStyle Font-Names="Verdana" Font-Size="8pt" ForeColor="Black" 
                HorizontalPadding="5px" NodeSpacing="0px" VerticalPadding="0px" />
        </asp:TreeView>
        <br />
        Nombre<asp:TextBox ID="txtNombre" runat="server"></asp:TextBox>
        
        
        <br />
        <br />
        Departmento<asp:DropDownList ID="dlDepartamentos" runat="server" 
            AutoPostBack="True" DataSourceID="SqlDataSource1" 
            DataTextField="RegionDescription" DataValueField="RegionID">
        </asp:DropDownList>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
            ConnectionString="<%$ ConnectionStrings:NorthwindConnectionString %>" 
            SelectCommand="SELECT * FROM [Region]"></asp:SqlDataSource>
        <br />
        <br />
        <br />
        
      
        
        <asp:Label ID="lblMensaje" runat="server"></asp:Label>
        <br />
        <asp:Button ID="Button1" runat="server" Text="Enviar" />
    
    </div>
    </form>
</body>
</html>
