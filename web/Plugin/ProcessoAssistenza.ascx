<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ProcessoAssistenza.ascx.cs" Inherits="Plugin_ProcessoAssistenza" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="cc1" %>

<cc1:NextList ID="Processo" runat="server" CssClass="processo" OnItemDataBound="Processo_ItemDataBound">
    <ItemTemplate>  
        <div class="passo" id="Passo" runat="server">
            <p class="passo" id="PassoP" runat="server"></p>
        </div>      
        <div class="titolo" id="Titolo" runat="server">
            <h1 class="titolo" id="TitoloH1" runat="server"></h1>
        </div>
        <div class="descrizione" id="Descrizione" runat="server">
            <p class="descrizione" id="DescrizioneP" runat="server"></p>
        </div>
    </ItemTemplate> 
</cc1:NextList>
