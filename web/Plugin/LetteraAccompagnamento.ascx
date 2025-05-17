<%@ Control Language="C#" AutoEventWireup="true" CodeFile="LetteraAccompagnamento.ascx.cs" Inherits="Plugin_LetteraAccompagnamento" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="nextFramework" %>

<div id="TestataDiv" class="testata" runat="server" visible="false">
    <div id="TestataTitoloDiv" class="titolo" runat="server">
        <h1 id="TestataTitolo" runat="server"></h1>
        <div id="TestataLinkDiv" runat="server" class="link">
            <a id="TestataLink" runat="server"></a>
        </div>
    </div>
    <div id="TestataDatiDiv" class="dati" runat="server">
        <div id="IndirizzoDiv" runat="server" class="indirizzo">
            <p id="Indirizzo" runat="server"></p>
            <p id="Citta" runat="server"></p>
            <p id="Fax" runat="server"></p>
            <p><a id="Email" runat="server"></a></p>
            <p id="PartitaIva" runat="server"></p>
        </div>
        <h2 id="Telefono" runat="server"></h2>
    </div>
</div>
<div id="IntestazioneDiv" class="intestazione" runat="server">
    <span id="TitoloLabel" runat="server" class="label"></span>
    <span id="DdtNumeroLabel" runat="server" class="label numero"></span>
    <span id="DdtNumeroValue" runat="server" class="value"></span>
    <span id="DdtDataLabel" runat="server" class="label data"></span>
    <span id="DdtDataValue" runat="server" class="value"></span>
</div>
<div id="DestinatarioDiv" class="destinatario" runat="server">
    <div class="doppio">
        <span id="DestinatarioNomeLabel" runat="server" class="label"></span>
        <span id="DestinatarioNomeValue" runat="server" class="value"></span>
        <span id="NumeroOrdineLabel" runat="server" class="label numero"></span>
    </div>
    <div class="intero">
        <span id="DestinatarioViaLabel" runat="server" class="label"></span>
        <span id="DestinatarioViaValue" runat="server" class="value"></span>
        <span id="DestinatarioCapLabel" runat="server" class="label"></span>
        <span id="DestinatarioCapValue" runat="server" class="value"></span>
    </div>
    <div class="sx">
        <div class="alto">
            <span id="DestinatarioCittaLabel" runat="server" class="label"></span>
            <span id="DestinatarioCittaValue" runat="server" class="value"></span>
            <span id="DestinatarioProvinciaValue" runat="server" class="value corto"></span>
        </div>
        <div class="basso">
            <span id="DestinatarioLocalitaLabel" runat="server" class="label"></span>
            <span id="DestinatarioLocalitaValue" runat="server" class="value"></span>
        </div>
    </div>
    <div class="dx">
        <span id="DdtCausaleLabel" runat="server" class="label"></span>
        <span id="DdtCausaleValue" runat="server" class="value"></span>
    </div>
</div>
<div id="TabellaDiv" class="tabella" runat="server">
    <div id="TabellaTitoloDiv" class="titolo" runat="server" visible="false">
        <span id="CodiceTitolo" runat="server"></span>
        <span id="DescrizioneTitolo" runat="server"></span>
        <span id="QuantitaTitolo" runat="server"></span>
    </div>
    <div id="TabellaDatiDiv" class="dati" runat="server">
        <asp:GridView id="TabellaDati" runat="server" AutoGenerateColumns="true">
        </asp:GridView>
    </div>
</div>
<div id="FondoDiv" class="fondo" runat="server">
    <div id="TrasportatoreDiv" class="trasportatore" runat="server">
        <div>
            <span id="TrasportatoreLabel" runat="server" class="label"></span>
            <span id="TrasportatoreValue" runat="server" class="value"></span>
        </div>
        <span id="DataRitiroLabel" runat="server" class="input"></span>
    </div>
    <div id="FirmeDiv" class="firme" runat="server">
        <div>
            <span id="FirmaConducente" runat="server"></span>
            <input disabled="disabled"/>
        </div>
        <div>
            <span id="FirmaVettore" runat="server"></span>
            <input disabled="disabled"/>
        </div>
        <div>
            <span id="FirmaDestinatario" runat="server"></span>
            <input disabled="disabled"/>
        </div>
    </div>
</div>