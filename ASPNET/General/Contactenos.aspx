<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Contactenos.aspx.vb" Inherits="General_Contactenos" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:SiteMapDataSource ID="SiteMapDataSource1" runat="server" />
    
    </div>
    <asp:SiteMapPath ID="SiteMapPath1" runat="server" Font-Names="Verdana" 
        Font-Size="0.8em" PathSeparator=" : ">
        <PathSeparatorStyle Font-Bold="True" ForeColor="#990000" />
        <CurrentNodeStyle ForeColor="#333333" />
        <NodeStyle Font-Bold="True" ForeColor="#990000" />
        <RootNodeStyle Font-Bold="True" ForeColor="#FF8000" />
    </asp:SiteMapPath>
    </form>
</body>
</html>
