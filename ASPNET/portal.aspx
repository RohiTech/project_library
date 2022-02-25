<%@ Page Language="VB" AutoEventWireup="false" CodeFile="portal.aspx.vb" Inherits="portal" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    <style type="text/css">
        .style1
        {
            width: 219px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:WebPartManager ID="WebPartManager1" runat="server">
        </asp:WebPartManager>
        <table style="width: 100%; height: 4px;">
            <tr>
                <td bgcolor="#CCCCCC" style="color=navy" colspan="2">
                    Bienvenidos a Nuestro Portal Web</td>
                <td>
                    <asp:Menu ID="Menu1" runat="server">
                    </asp:Menu>
                </td>
            </tr>
            <tr>
                <td class="style1" valign="top">
                    <asp:CatalogZone ID="CatalogZone1" runat="server" BackColor="#F7F6F3" 
                        BorderColor="#CCCCCC" BorderWidth="1px" Font-Names="Verdana" Padding="6">
                        <HeaderVerbStyle Font-Bold="False" Font-Size="0.8em" Font-Underline="False" ForeColor="#333333">
                        </HeaderVerbStyle>
                        <PartTitleStyle BackColor="#5D7B9D" Font-Bold="True" Font-Size="0.8em" ForeColor="White">
                        </PartTitleStyle>
                        <FooterStyle HorizontalAlign="Right" BackColor="#E2DED6"></FooterStyle>
                        <PartChromeStyle BorderColor="#E2DED6" BorderWidth="1px" BorderStyle="Solid">
                        </PartChromeStyle>
                        <PartLinkStyle Font-Size="0.8em"></PartLinkStyle>
                        <InstructionTextStyle Font-Size="0.8em" ForeColor="#333333">
                        </InstructionTextStyle>
                        <ZoneTemplate>
                            <asp:PageCatalogPart ID="PageCatalogPart1" runat="server" />
                        </ZoneTemplate>
                        <LabelStyle Font-Size="0.8em" ForeColor="#333333"></LabelStyle>
                        <SelectedPartLinkStyle Font-Size="0.8em"></SelectedPartLinkStyle>
                        <VerbStyle Font-Names="Verdana" Font-Size="0.8em" ForeColor="#333333">
                        </VerbStyle>
                        <HeaderStyle BackColor="#E2DED6" Font-Bold="True" Font-Size="0.8em" ForeColor="#333333">
                        </HeaderStyle>
                        <EditUIStyle Font-Names="Verdana" Font-Size="0.8em" ForeColor="#333333">
                        </EditUIStyle>
                        <PartStyle BorderColor="#F7F6F3" BorderWidth="5px"></PartStyle>
                        <EmptyZoneTextStyle Font-Size="0.8em" ForeColor="#333333"></EmptyZoneTextStyle>
                    </asp:CatalogZone>
                </td>
                <td valign="top">
                    <asp:WebPartZone ID="ZonaContenido" runat="server" BorderColor="#CCCCCC" 
                        Font-Names="Verdana" Padding="6" Height="196px">
                        <EmptyZoneTextStyle Font-Size="0.8em"></EmptyZoneTextStyle>
                        <PartStyle Font-Size="0.8em" ForeColor="#333333"></PartStyle>
                        <TitleBarVerbStyle Font-Size="0.6em" Font-Underline="False" ForeColor="White">
                        </TitleBarVerbStyle>
                        <MenuLabelHoverStyle ForeColor="#E2DED6"></MenuLabelHoverStyle>
                        <MenuPopupStyle BackColor="#5D7B9D" BorderColor="#CCCCCC" BorderWidth="1px" 
                            Font-Names="Verdana" Font-Size="0.6em">
                        </MenuPopupStyle>
                        <MenuVerbStyle BorderColor="#5D7B9D" BorderWidth="1px" BorderStyle="Solid" 
                            ForeColor="White"></MenuVerbStyle>
                        <PartTitleStyle BackColor="#5D7B9D" Font-Bold="True" Font-Size="0.8em" 
                            ForeColor="White"></PartTitleStyle>
                        <ZoneTemplate>
                            <asp:GridView ID="grid" runat="server" AutoGenerateColumns="False" 
                                DataSourceID="SqlDataSource1">
                                <Columns>
                                    <asp:BoundField DataField="productname" HeaderText="productname" 
                                        SortExpression="productname" />
                                    <asp:BoundField DataField="categoryID" HeaderText="categoryID" 
                                        SortExpression="categoryID" />
                                </Columns>
                            </asp:GridView>
                        </ZoneTemplate>
                        <MenuVerbHoverStyle BackColor="#F7F6F3" BorderColor="#CCCCCC" BorderWidth="1px" 
                            BorderStyle="Solid" ForeColor="#333333"></MenuVerbHoverStyle>
                        <PartChromeStyle BackColor="#F7F6F3" BorderColor="#E2DED6" Font-Names="Verdana" 
                            ForeColor="White"></PartChromeStyle>
                        <HeaderStyle HorizontalAlign="Center" Font-Size="0.7em" ForeColor="#CCCCCC">
                        </HeaderStyle>
                        <MenuLabelStyle ForeColor="White"></MenuLabelStyle>
                    </asp:WebPartZone>
                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
                        ConnectionString="<%$ ConnectionStrings:NorthwindConnectionString %>" 
                        SelectCommand="select top 20 productname, categoryID from products">
                    </asp:SqlDataSource>
                </td>
                <td valign="top">
                    <asp:WebPartZone ID="ZonaInformacion" runat="server" BorderColor="#CCCCCC" 
                        Font-Names="Verdana" Padding="6">
                        <EmptyZoneTextStyle Font-Size="0.8em"></EmptyZoneTextStyle>
                        <PartStyle Font-Size="0.8em" ForeColor="#333333"></PartStyle>
                        <TitleBarVerbStyle Font-Size="0.6em" Font-Underline="False" ForeColor="White">
                        </TitleBarVerbStyle>
                        <MenuLabelHoverStyle ForeColor="#E2DED6"></MenuLabelHoverStyle>
                        <MenuPopupStyle BackColor="#5D7B9D" BorderColor="#CCCCCC" BorderWidth="1px" Font-Names="Verdana" Font-Size="0.6em">
                        </MenuPopupStyle>
                        <MenuVerbStyle BorderColor="#5D7B9D" BorderWidth="1px" BorderStyle="Solid" ForeColor="White">
                        </MenuVerbStyle>
                        <PartTitleStyle BackColor="#5D7B9D" Font-Bold="True" Font-Size="0.8em" ForeColor="White">
                        </PartTitleStyle>
                        <ZoneTemplate>
                            <asp:DropDownList ID="DropDownList1" runat="server">
                                <asp:ListItem>Leon</asp:ListItem>
                                <asp:ListItem>Managua</asp:ListItem>
                                <asp:ListItem>Masaya</asp:ListItem>
                                <asp:ListItem>Esteli</asp:ListItem>
                            </asp:DropDownList>
                            <asp:Calendar ID="Calendar1" runat="server"></asp:Calendar>
                            <asp:AdRotator ID="AdRotator1" runat="server" 
                                AdvertisementFile="~/Anuncios.xml" />
                        </ZoneTemplate>
                        <MenuVerbHoverStyle BackColor="#F7F6F3" BorderColor="#CCCCCC" BorderWidth="1px" BorderStyle="Solid" ForeColor="#333333">
                        </MenuVerbHoverStyle>
                        <PartChromeStyle BackColor="#F7F6F3" BorderColor="#E2DED6" Font-Names="Verdana" ForeColor="White">
                        </PartChromeStyle>
                        <HeaderStyle HorizontalAlign="Center" Font-Size="0.7em" ForeColor="#CCCCCC">
                        </HeaderStyle>
                        <MenuLabelStyle ForeColor="White"></MenuLabelStyle>
                    </asp:WebPartZone>
                </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
