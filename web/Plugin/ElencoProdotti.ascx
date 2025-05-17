<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ElencoProdotti.ascx.cs" Inherits="Plugin_ElencoProdotti" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="cc1" %>

<h1 id="TitoloSez" class="titolo_sez" runat="server"></h1>

<cc1:NextList ID="Prodotti" runat="server" CssClass="prodotti" OnItemDataBound="Prodotti_ItemDataBound">
<ItemTemplate>        
    <a href="" id="Link" runat="server"></a>
</ItemTemplate>
    
</cc1:NextList><div id="IndietroDiv" class="button" runat="server">
    <asp:Button OnClick="Indietro_Click" ID="Indietro" runat="server" CssClass="indietro" Text="Indietro" />
</div>
