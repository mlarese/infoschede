<%@ Control Language="C#" AutoEventWireup="true" CodeFile="Bolla.ascx.cs" Inherits="Plugin_Bolla" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="nextFramework" %>

<span id="DdtTitle" runat="server" class="noborder"></span>
<div class="contenitore">
    <div id="IntestazioneDiv" class="intestazione" runat="server">
        <div>
            <span id="DdtNumeroLabel" runat="server" class="label labelsx"></span>
            <span id="DdtNumeroValue" runat="server" class="value valuesx"></span>
        </div>
        <div>
            <span id="DdtDataLabel" runat="server" class="label labeldx"></span>
            <span id="DdtDataValue" runat="server" class="value valuedx"></span>
        </div>
    </div>
    <div id="ClienteDiv" class="intestazione" runat="server">
        <div>
            <span id="ClienteLabel" runat="server" class="label"></span>
            <span id="ClienteValue" runat="server" class="value"></span>
        </div>
        <div>
            <span id="PartitaIvaLabel" runat="server" class="label"></span>
            <span id="PartitaIvaValue" runat="server" class="value"></span>
        </div>
        <div class="cf">
            <span id="CodiceFiscaleLabel" runat="server" class="label"></span>
            <span id="CodiceFiscaleValue" runat="server" class="value"></span>
        </div>
        <div>
            <span id="AgenteLabel" runat="server" class="label"></span>
            <span id="AgenteValue" runat="server" class="value"></span>
        </div>
        <div>
            <span id="TelefonoLabel" runat="server" class="label"></span>
            <span id="TelefonoValue" runat="server" class="value"></span>
        </div>
    </div>
</div>
<div id="DestinatarioDiv" class="destinatario" runat="server">
    <div>
        <span id="DestinatarioTitle" runat="server" class="title"></span>
        <span id="DestinatarioNomeLabel" runat="server" class="label"></span>
        <span id="DestinatarioNomeValue" runat="server" class="value"></span>
        <span id="DestinatarioViaLabel" runat="server" class="label"></span>
        <span id="DestinatarioViaValue" runat="server" class="value"></span>
        <span id="DestinatarioCittaLabel" runat="server" class="label"></span>
        <span id="DestinatarioCittaValue" runat="server" class="value"></span>
    </div>
</div>
<div id="IndirizzoDiv" class="destinatario" runat="server">
    <div>
        <span id="IndirizzoTitle" runat="server" class="title"></span>
        <span id="IndirizzoNomeLabel" runat="server" class="label"></span>
        <span id="IndirizzoNomeValue" runat="server" class="value"></span>
        <span id="IndirizzoViaLabel" runat="server" class="label"></span>
        <span id="IndirizzoViaValue" runat="server" class="value"></span>
        <span id="IndirizzoCittaLabel" runat="server" class="label"></span>
        <span id="IndirizzoCittaValue" runat="server" class="value"></span>
    </div>
</div>
<div id="TabellaDiv" class="tabella" runat="server">
    <div id="TabellaTitoloDiv" class="titolo" runat="server" visible="false">
        <span id="QuantitaTitolo" runat="server"></span>
        <span id="CodiceTitolo" runat="server"></span>
        <span id="DescrizioneTitolo" runat="server"></span>
        <span id="NumeroSchedaTitolo" runat="server"></span>
        <span id="NumeroVSDDTitolo" runat="server"></span>
        <span id="GaranziaTitolo" runat="server"></span>
    </div>
    <div id="TabellaDatiDiv" class="dati" runat="server">
        <asp:GridView id="TabellaDati" runat="server" AutoGenerateColumns="true">
        </asp:GridView>
    </div>
</div>
<div id="FondoDiv" class="fondo" runat="server">
    <table cellpadding="0" cellspacing="0">
        <tr>
            <td class="annotazioni" colspan="8">
                <span id="AnnotazioniLabel" runat="server" class="label"></span>
                <span id="AnnotazioniValue" runat="server" class="value"></span>
            </td>
        </tr>
        <tr>
            <td colspan="4">
                <span id="DdtCausaleLabel" runat="server" class="label"></span>
                <span id="DdtCausaleValue" runat="server" class="value"></span>
            </td>
            <td colspan="2">
                <span id="DataOraTrasportoLabel" runat="server" class="label"></span>
                <span id="DataOraTrasportoValue" runat="server" class="value"></span>
            </td>
            <td colspan="2">
                <span id="FirmaConducenteLabel" runat="server" class="label"></span>
                <span id="FirmaConducenteValue" runat="server" class="value" visible="false"></span>
            </td>
        </tr>
        <tr>
            <td colspan="2">
                <span id="TrasportoCuraLabel" runat="server" class="label"></span>
                <span id="TrasportoCuraValue" runat="server" class="value"></span>
            </td>
            <td colspan="2">
                <span id="PortoLabel" runat="server" class="label"></span>
                <span id="PortoValue" runat="server" class="value"></span>
            </td>
            <td colspan="1">
                <span id="ColliLabel" runat="server" class="label"></span>
                <span id="ColliValue" runat="server" class="value"></span>
            </td>
            <td colspan="1">
                <span id="PesoLabel" runat="server" class="label"></span>
                <span id="PesoValue" runat="server" class="value"></span>
            </td>
            <td colspan="2">
                <span id="FirmaDestinatarioLabel" runat="server" class="label"></span>
                <span id="FirmaDestinatarioValue" runat="server" class="value" visible="false"></span>
            </td>
        </tr>
        <tr>
            <td colspan="4">
                <span id="VettoreLabel" runat="server" class="label"></span>
                <span id="VettoreValue" runat="server" class="value"></span>
            </td>
            <td colspan="2">
                <span id="DataOraRitiroLabel" runat="server" class="label"></span>
                <span id="DataOraRitiroValue" runat="server" class="value"></span>
            </td>
            <td colspan="2">
                <span id="FirmaVettoreLabel" runat="server" class="label"></span>
                <span id="FirmaVettoreValue" runat="server" class="value"></span>
            </td>
        </tr>
    </table>
    <span id="Note" runat="server" class="noborder note"></span>
</div>
