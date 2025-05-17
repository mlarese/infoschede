<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ElencoProblemi.ascx.cs" Inherits="Plugin_ElencoProblemi" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="cc1" %>

<h1 id="TitoloSez" class="titolo_sez" runat="server"></h1>

<cc1:NextList ID="Problemi" runat="server" CssClass="problemi" OnItemDataBound="Problemi_ItemDataBound">
    <ItemTemplate>
        <asp:Panel ID="Panel1" runat="server" CssClass="faq_pnl">
            <a href="" id="Link" runat="server"></a>
            <div class="container" id="Container" runat="server">
                <p id="Soluzioni" runat="server" class="soluzioni"></p>
                <p id="Domanda" runat="server" class="domanda"></p>
                <asp:Button OnClick="Ok_Clicked" ID="Ok" runat="server" Text="Sì, risolto" />
                <asp:Button OnClick="Assistenza_Clicked" ID="Assistenza" runat="server" Text="No, richiedi assistenza" />
            </div>
        </asp:Panel>
    </ItemTemplate>
</cc1:NextList>

<div id="IndietroDiv" class="button" runat="server">
    <asp:Button OnClick="Indietro_Click" ID="Indietro" runat="server" CssClass="indietro" Text="Indietro" />
</div>
