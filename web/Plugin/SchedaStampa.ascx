<%@ Control Language="C#" AutoEventWireup="true" CodeFile="SchedaStampa.ascx.cs" Inherits="Plugin_SchedaStampa" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="nextFramework" %>

<a class="print_link" href="javascript:window.print()">STAMPA</a>
<h1 id="Titolo" runat="server" visible="false" class="modulostampa"></h1>
<p id="Descrizione" runat="server" visible="false" class="modulostampa"></p>
<table cellpadding="0" cellspacing="0" class="modulostampa">
    <tr>
        <td class="cella">
            <span id="NumeroLabel" runat="server" class="label">NUMERO</span>
            <span id="NumeroValue" runat="server" class="value"></span>
        </td>
        <td class="cella">
            <span id="DataLabel" runat="server" class="label">DATA</span>
            <span id="DataValue" runat="server" class="value"></span>
        </td>
        <td class="cella">
            <span id="StatoLabel" runat="server" class="label">STATO</span>
            <span id="StatoValue" runat="server" class="value"></span>
        </td>
        <td rowspan="5" colspan="3">
            <div class="box">
                <span id="DestinatarioTitle" runat="server" class="label">CLIENTE</span>
                <span id="DestinatarioNomeLabel" runat="server" class="label"></span>
                <span id="DestinatarioNomeValue" runat="server" class="value"></span>
                <span id="DestinatarioViaLabel" runat="server" class="label"></span>
                <span id="DestinatarioViaValue" runat="server" class="value"></span>
                <span id="DestinatarioCittaLabel" runat="server" class="label"></span>
                <span id="DestinatarioCittaValue" runat="server" class="value"></span>
            </div>
        </td>
    </tr>
    <tr>
        <th colspan="3">Ritiro</th>
    </tr>
    <tr>
        <td class="cella">
            <span id="RitiroRifLabel" runat="server" class="label">VS. RIFERIMENTO</span>
            <span id="RitiroRifValue" runat="server" class="value"></span>
        </td>
        <td class="cellacolspan2" colspan="2">
            <span id="RitiroRifDdtLabel" runat="server" class="label">RIF. VS. DDT</span>
            <span id="RitiroRifDdtValue" runat="server" class="value"></span>
        </td>
    </tr>
    <tr>
        <th colspan="3">Riconsegna</th>
    </tr>
    <tr>
        <td class="cella">
            <span id="TrasportatoreLabel" runat="server" class="label">TRASPORTATORE</span>
            <span id="TrasportatoreValue" runat="server" class="value"></span>
        </td>
        <td class="cellacolspan2" colspan="2">
            <span id="ConsegnaRifLabel" runat="server" class="label">DDT</span>
            <span id="ConsegnaRifValue" runat="server" class="value"></span>
        </td>
    </tr>
    <tr><th colspan="6" >MACCHINA</th></tr>
    <tr>
        <td class="cella">
            <span id="CostruttoreLabel" runat="server" class="label">COSTRUTTORE</span>
            <span id="CostruttoreValue" runat="server" class="value"></span>
        </td>
        <td class="cella" colspan="5">
            <span id="ModelloLabel" runat="server" class="label">MODELLO</span>
            <span id="ModelloValue" runat="server" class="value"></span>
        </td>
    </tr>
    <tr>
        <td class="cellacolspan2" colspan="2">
            <span id="MatricolaLabel" runat="server" class="label">MATRICOLA</span>
            <span id="MatricolaValue" runat="server" class="value"></span>
        </td>
        <td class="cella">
            <span id="DataAcquistoLabel" runat="server" class="label">DATA  ACQUISTO</span>
            <span id="DataAcquistoValue" runat="server" class="value"></span>
        </td>
        <td class="cellacolspan2" colspan="2">
            <span id="ScontrinoLabel" runat="server" class="label">NUMERO SCONTRINO</span>
            <span id="ScontrinoValue" runat="server" class="value"></span>
        </td>
        <td class="cella">
            <span id="GaranziaLabel" runat="server" class="label">GARANZIA</span>
            <span id="GaranziaValue" runat="server" class="value"></span>
        </td>
    </tr>
    <tr>
        <td class="cellacolspan6" colspan="6">
            <span id="AccessoriLabel" runat="server" class="label">Accessori presenti</span>
            <span id="AccessoriValue" runat="server" class="value"></span>
        </td>
    </tr>
    <tr><th colspan="6">SEGNALAZIONI CLIENTE</th></tr>
    <tr>
        <td class="cellacolspan6" colspan="6">
            <span id="GuastoSegnalatoLabel" runat="server" class="label">Guasto segnalato</span>
            <span id="GuastoSegnalatoValue" runat="server" class="value"></span>
        </td>
    </tr>
    <tr>
        <td class="cellacolspan6" colspan="6">
            <span id="NoteClienteLabel" runat="server" class="label">Note del cliente</span>
            <span id="NoteClienteValue" runat="server" class="value"></span>
        </td>
    </tr>
    <tr><th colspan="6">RIPARAZIONE</th></tr>
    <tr>
        <td class="cellacolspan5" colspan="4">
            <span id="GuastoRiscontratoLabel" runat="server" class="label">Guasto riscontrato</span>
            <span id="GuastoRiscontratoValue" runat="server" class="value"></span>
        </td>
        <td class="cella">
            <span id="DataFineLavoroLabel" runat="server" class="label">Data fine lavoro</span>
            <span id="DataFineLavoroValue" runat="server" class="value"></span>
        </td>
        <td class="cella">
            <span id="EsitoRiparazioneLabel" runat="server" class="label">Esito riparazione</span>
            <span id="EsitoRiparazioneValue" runat="server" class="value"></span>
        </td>
    </tr>
    <tr>
        <td class="cellacolspan6" colspan="6">
            <span id="NoteRiparazioneLabel" runat="server" class="label">Note di chiusura</span>
            <span id="NoteRiparazioneValue" runat="server" class="value"></span>
        </td>
    </tr>
    <tr><th colspan="6">MANODOPERA E RICAMBI UTILIZZATI</th></tr>
    <tr>
        <td class="cella">
            <span id="ManodoperaOreLabel" runat="server" class="label">Ore manodopera</span>
            <span id="ManodoperaOreValue" runat="server" class="value"></span>
        </td>
        <td class="cella">
            <span id="ManodoperaPrezzoLabel" runat="server" class="label">Prezzo orario</span>
            <span id="ManodoperaPrezzoValue" runat="server" class="value"></span>
        </td>
        <td class="cellacolspan2" colspan="2">
            <span id="ManodoperaTotaleLabel" runat="server" class="label">Totale manodopera</span>
            <span id="ManodoperaTotaleValue" runat="server" class="value"></span>
        </td>
        <td class="cella">
            <span id="CostoPresaLabel" runat="server" class="label">Costo presa</span>
            <span id="CostoPresaValue" runat="server" class="value"></span>
        </td>
        <td class="cella">
            <span id="CostoRiconsegnaLabel" runat="server" class="label">Costo riconsegna</span>
            <span id="CostoRiconsegnaValue" runat="server" class="value"></span>
        </td>
    </tr>
    <tr runat="server" id="RicambiNessunoTr">
        <td colspan="6" class="borded">
            <span class="value">Nessun pezzo sostituito.</span>
        </td>
    </tr>
    <tr runat="server" id="RicambiListaTr">
        <td colspan="6">
            <asp:GridView id="RicambiUtilizzatiLista" runat="server" CssClass="value" AutoGenerateColumns="false"
                          OnRowDataBound="RicambiUtilizzatiLista_RowDataBound" CellPadding="6" CellSpacing="0">
                <Columns>
                    <asp:TemplateField>
                        <HeaderTemplate>
                            <span id="codiceHd" runat="server" class="label"></span>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <span id="codice" runat="server" class="value"></span>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField>
                        <HeaderTemplate>
                            <span id="ricambioHd" runat="server" class="label"></span>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <span id="ricambio" runat="server" class="value"></span>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField>
                        <HeaderTemplate>
                            <span id="prezzoHd" runat="server" class="label"></span>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <span id="prezzo" runat="server" class="value" style="text-align:right;"></span>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField>
                        <HeaderTemplate>
                            <span id="quantitaHd" runat="server" class="label"></span>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <span id="quantita" runat="server" class="value" style="text-align:right;"></span>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField>
                        <HeaderTemplate>
                            <span id="scontoHd" runat="server" class="label"></span>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <span id="sconto" runat="server" class="value" style="text-align:right;"></span>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField>
                        <HeaderTemplate>
                            <span id="totaleHd" runat="server" class="label"></span>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <span id="totale" runat="server" class="value" style="text-align:right;"></span>
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </asp:GridView>
        </td>
    </tr>
    <tr id="TotaleRicambiTr" runat="server">
        <td colspan="4"></td>
        <td class="cellatotali" colspan="2">
            <span id="TotaleRicambiLabel" runat="server" class="label">Totale ricambi</span>
            <span id="TotaleRicambiValue" runat="server" class="value"></span>
        </td>
    </tr>
    <tr id="TotaleIvaTr" runat="server">
        <td colspan="4"></td>
        <td class="cellatotali" colspan="2">
            <span id="TotaleIvaLabel" runat="server" class="label">I.v.a.</span>
            <span id="TotaleIvaValue" runat="server" class="value"></span>
        </td>
    </tr>
    <tr id="TotaleGeneraleTr" runat="server">
        <td colspan="4"></td>
        <td class="cellatotali" colspan="2">
            <span id="TotaleGeneraleLabel" runat="server" class="label">Totale scheda</span>
            <span id="TotaleGeneraleValue" runat="server" class="value"></span>
        </td>
    </tr>
    
    <tr id="DescrittoriTitleTr" runat="server"><th colspan="6">CONTROLLI EFFETTUATI</th></tr>
    
    <tr runat="server" id="DescrittoriListaTr">
        <td colspan="6" class="cellacolspan6">
            <nextFramework:NextList ID="Descrittori" runat="server" OnItemDataBound="Descrittori_ItemDataBound">
                <ItemTemplate>
                    <span id="DescrittoriLabel" runat="server" class="descrittorelabel"></span>
                    <span id="DescrittoriValue" runat="server" class="descrittorevalue"></span>
                </ItemTemplate>
            </nextFramework:NextList>
        </td>
    </tr>
</table>