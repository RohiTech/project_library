<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default2.aspx.vb" Inherits="Default2" %>

<%@ Register assembly="DevExpress.Web.ASPxPivotGrid.v8.2, Version=8.2.4.0, Culture=neutral, PublicKeyToken=9b171c9fd64da1d1" namespace="DevExpress.Web.ASPxPivotGrid" tagprefix="dxwpg" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <dxwpg:ASPxPivotGrid ID="ASPxPivotGrid1" runat="server" CssClass="" 
            DataSourceID="SqlDataSource1">
            <Fields>
                <dxwpg:PivotGridField ID="field" Area="RowArea" AreaIndex="0" 
                    FieldName="ProductName" Name="field">
                </dxwpg:PivotGridField>
                <dxwpg:PivotGridField ID="field1" Area="ColumnArea" AreaIndex="0" 
                    FieldName="Mes" Name="field1">
                </dxwpg:PivotGridField>
                <dxwpg:PivotGridField ID="field2" Area="RowArea" AreaIndex="1" 
                    FieldName="ShipCountry" Name="field2">
                </dxwpg:PivotGridField>
                <dxwpg:PivotGridField ID="field3" Area="DataArea" AreaIndex="0" 
                    FieldName="Total" Name="field3">
                </dxwpg:PivotGridField>
            </Fields>
        </dxwpg:ASPxPivotGrid>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
            ConnectionString="<%$ ConnectionStrings:NorthwindConnectionString %>" SelectCommand=" SELECT ShipCountry, ShipCity,
 p.ProductName,
  od.Quantity, od.UnitPrice, od.Quantity*od.UnitPrice AS Total,
 MONTH(o.OrderDate) AS Mes, YEAR(o.OrderDate) AS Año
   FROM Orders o
 INNER JOIN [Order Details] od
 ON o.OrderID = od.OrderID
 INNER JOIN Products p ON od.ProductID = p.ProductID
"></asp:SqlDataSource>
    
    </div>
    </form>
</body>
</html>
