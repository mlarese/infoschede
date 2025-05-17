<%@ Control Language="C#" AutoEventWireup="true" CodeFile="Footer.ascx.cs" Inherits="Plugin_Footer" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="cc1" %>
<asp:Panel ID="DivMoreBox" runat="server" CssClass="divmorebox">
      
</asp:Panel>
<asp:Panel ID="DivLeft" runat="server" CssClass="divleft">
    <asp:HyperLink ID="Home" runat="server" CssClass="Home">
    <asp:Image ID="Logo" runat="server"   />
    </asp:HyperLink>
    <p runat="server" id="Indirizzo" class="indirizzo">
    </p>
    <p class="credits">
        <cc1:PoweredBy ID="PoweredBy1" runat="server" />
    </p>
</asp:Panel>
<asp:Panel ID="DivMenu" runat="server" CssClass="divmenu">
    <cc1:NextList ID="ListaMenu" runat="server" OnItemDataBound="ListaMenu_ItemDataBound" CssClass="listamenu">
        <ItemTemplate>
            <div id="Contenitore" runat="server">
                <h1 id="Titolo" runat="server">
                
                </h1>
                <cc1:NextList ID="menu" runat="server" OnItemDataBound="menu_ItemDataBound" CssClass="menu">
                    <ItemTemplate>
                        <a id="Link" runat="server" href="">
                        </a>
                    </ItemTemplate>
                </cc1:NextList>
            </div>
        </ItemTemplate>
    </cc1:NextList>
</asp:Panel>

