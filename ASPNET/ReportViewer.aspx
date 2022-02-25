﻿<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ReportViewer.aspx.vb" Inherits="ReportViewer" %>

<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=9.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

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
        <asp:Button ID="Button1" runat="server" Text="Mostrar" />
        <br />
        <br />
    
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Height="600px" 
             Width="100%" ProcessingMode="Remote">
        </rsweb:ReportViewer>
    
    </div>
    </form>
</body>
</html>
