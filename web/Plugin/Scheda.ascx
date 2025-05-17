<%@ Control Language="C#" AutoEventWireup="true" CodeFile="Scheda.ascx.cs" Inherits="Plugin_Scheda" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="nextFramework" %>

<h1 id="Titolo" class="titolo_sez" runat="server"></h1>
<div id="DescrizioneDiv" class="descrizione" runat="server">
    <p id="Descrizione" runat="server"></p>
</div>
<div id="ErroriListaDiv" class="errori" runat="server" visible="false"></div>

<div id="SchedaDiv" runat="server" class="dati">
    <div id="TitoloPrincipaliDiv" runat="server" class="sezione">
        <h2 id="TitoloPrincipali" runat="server"></h2>
    </div>
    <div id="StatoSchedaDiv" class="stato" runat="server">
        <span id="StatoSchedaLabel" runat="server" class="label"></span>
        <span id="StatoSchedaValue" runat="server" class="value"></span>
    </div>
    <div id="NumeroSchedaDiv" class="numero" runat="server">
        <span id="NumeroSchedaLabel" runat="server" class="label"></span>
        <span id="NumeroSchedaValue" runat="server" class="value"></span>
    </div>
    <div id="DataRicevimentoDiv" class="data" runat="server">
        <span id="DataRicevimentoLabel" runat="server" class="label"></span>
        <span id="DataRicevimentoValue" runat="server" class="value"></span>
    </div>
    <div id="CentroAssistenzaDiv" class="cliente" runat="server">
        <span id="CentroAssistenzaLabel" runat="server" class="label"></span>
        <span id="CentroAssistenzaValue" runat="server" class="value"></span>
    </div>
    <div id="ClienteDiv" class="cliente" runat="server">
        <span id="ClienteLabel" runat="server" class="label"></span>
        <span id="ClienteValue" runat="server" class="value"></span>
    </div>
    <div id="RiferimentoClienteDiv" class="riferimento" runat="server">
        <span id="RiferimentoClienteLabel" runat="server" class="label"></span>
        <span id="RiferimentoClienteValue" runat="server" class="value"></span>
        <input id="RiferimentoClienteInput" runat="server" class="input lungo" maxlength="250" />
    </div>
    
    <div id="TitoloModelloDiv" runat="server" class="sezione">
        <h2 id="TitoloModello" runat="server"></h2>
    </div>
    <div class="modelli">
        <div id="ModelloDiv" class="modello" runat="server">
            <span id="ModelloLabel" runat="server" class="label"></span>
            <span id="ModelloValue" runat="server" class="value"></span>
            <input id="ModelloInput" runat="server" class="input" maxlength="250" />
            <img id="ModelloInfoImg" runat="server" class="info" />
        </div>
        <div id="ModelloVariantiDiv" class="modello" runat="server">
            <span id="ModelloVariantiLabel" runat="server" class="label varianti"></span>
            <nextFramework:NextDropDownList id="ModelloVariantiLista" runat="server" />
        </div>
        <div id="MatricolaDiv" class="matricola" runat="server">
            <span id="MatricolaLabel" runat="server" class="label"></span>
            <span id="MatricolaValue" runat="server" class="value"></span>
            <input id="MatricolaInput" runat="server" class="input" maxlength="250" />
            <img id="MatricolaInfoImg" runat="server" class="info" />
            <img id="MatricolaEsempioImg" runat="server" class="esempio nmatricola" />
        </div>
        <div id="AccessoriListaDiv" class="lista_accessori" runat="server">
            <span id="AccessoriListaLabel" runat="server" class="label"></span>
            <span id="AccessoriListaValue" runat="server" class="value"></span>
            <nextFramework:NextDropDownList id="AccessoriListaDdl" runat="server" />
        </div>
        <div id="AccessoriAltroDiv" class="altri_accessori" runat="server">
            <span id="AccessoriAltroLabel" runat="server" class="label" style="width:93px !important;margin-left:78px;"></span>
            <span id="AccessoriAltroValue" runat="server" class="value"></span>
            <input id="AccessoriAltroInput" runat="server" class="input lungo" maxlength="250" />
        </div>
    </div>
    <div id="LogoMarcaDiv" class="marca" runat="server">
        <img id="LogoMarcaImg" runat="server" />
    </div>
    
    <div id="TitoloAcquistoDiv" runat="server" class="sezione">
        <h2 id="TitoloAcquisto" runat="server"></h2>
    </div>
    <div id="DataAcquistoDiv" class="data dp_container" runat="server">
        <span id="DataAcquistoLabel" runat="server" class="label"></span>
        <span id="DataAcquistoValue" runat="server" class="value"></span>
        <input id="DataAcquistoInput" runat="server" class="input datepicker" type="text" maxlength="10" />
    </div>
    <div id="NegozioAcquistoDiv" class="acquisto" runat="server">
        <span id="NegozioAcquistoLabel" runat="server" class="label"></span>
        <span id="NegozioAcquistoValue" runat="server" class="value"></span>
        <input id="NegozioAcquistoInput" runat="server" class="input" type="text" maxlength="250" />
    </div>
    <div id="NumeroScontrinoDiv" class="nscontrino" runat="server">
        <span id="NumeroScontrinoLabel" runat="server" class="label"></span>
        <span id="NumeroScontrinoValue" runat="server" class="value"></span>
        <input id="NumeroScontrinoInput" runat="server" class="input" maxlength="250" />
        <img id="NumeroScontrinoInfoImg" runat="server" class="info" />
        <img id="NumeroScontrinoEsempioImg" runat="server" class="esempio nscontrino" />
    </div>
    <div id="GaranziaDiv" class="garanzia" runat="server">
        <span id="GaranziaLabel" runat="server" class="label"></span>
        <span id="GaranziaValue" runat="server" class="value"></span>
        <asp:CheckBox id="GaranziaCb" runat="server" Checked="false" />
    </div>
    
    <div id="TitoloRiparazioneDiv" runat="server" class="sezione">
        <h2 id="TitoloRiparazione" runat="server"></h2>
    </div>
    <div id="GuastoSegnalatoDiv" class="guasto" runat="server">
        <span id="GuastoSegnalatoLabel" runat="server" class="label"></span>
        <span id="GuastoSegnalatoValue" runat="server" class="value"></span>
    </div>
    <div id="GuastoSegnalatoAltroDiv" class="guasto" runat="server">
        <span id="GuastoSegnalatoAltroLabel" runat="server" class="label"></span>
        <span id="GuastoSegnalatoAltroValue" runat="server" class="value"></span>
        <input id="GuastoSegnalatoAltroInput" runat="server" class="input lungo" />
    </div>
    <div id="NoteClienteDiv" class="note" runat="server">
        <span id="NoteClienteLabel" runat="server" class="label" style="vertical-align:top;"></span>
        <span id="NoteClienteValue" runat="server" class="value"></span>
        <textarea id="NoteClienteTxtarea" runat="server" class="input" rows="4" cols="110"></textarea>
    </div>
    <div id="GuastoRiscontratoDiv" class="guasto" runat="server">
        <span id="GuastoRiscontratoLabel" runat="server" class="label"></span>
        <span id="GuastoRiscontratoValue" runat="server" class="value"></span>
    </div>
    <div id="GuastoRiscontratoAltroDiv" class="guasto" runat="server" visible="false">
        <span id="GuastoRiscontratoAltroLabel" runat="server" class="label"></span>
        <span id="GuastoRiscontratoAltroValue" runat="server" class="value"></span>
    </div>
    <div id="EsitoInterventoDiv" class="esito" runat="server">
        <span id="EsitoInterventoLabel" runat="server" class="label"></span>
        <span id="EsitoInterventoValue" runat="server" class="value"></span>
    </div>
    <div id="DataFineLavoroDiv" class="data" runat="server">
        <span id="DataFineLavoroLabel" runat="server" class="label"></span>
        <span id="DataFineLavoroValue" runat="server" class="value"></span>
    </div>
    <div id="OreManodoperaDiv" class="ore" runat="server">
        <span id="OreManodoperaLabel" runat="server" class="label"></span>
        <span id="OreManodoperaValue" runat="server" class="value"></span>
    </div>
    <div id="PrezzoManodoperaDiv" class="prezzo" runat="server">
        <span id="PrezzoManodoperaLabel" runat="server" class="label"></span>
        <span id="PrezzoManodoperaValue" runat="server" class="value"></span>
    </div>
    <div id="RicambiUtilizzatiDiv" class="ricambi" runat="server">
        <span id="RicambiUtilizzatiLabel" runat="server" class="label"></span>
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
    </div>
    <div id="NoteChiusuraDiv" class="note" runat="server">
        <span id="NoteChiusuraLabel" runat="server" class="label" style="vertical-align:top;"></span>
        <span id="NoteChiusuraValue" runat="server" class="value"></span>
    </div>
    
    <div id="TitoloControlliDiv" runat="server" class="sezione">
        <h2 id="TitoloControlli" runat="server"></h2>
    </div>
    <div id="DescrittoriDiv" class="descrittori" runat="server">
        <nextFramework:NextList ID="Descrittori" runat="server" OnItemDataBound="Descrittori_ItemDataBound">
            <ItemTemplate>
                <span id="DescrittoriLabel" runat="server" class="label"></span>
                <span id="DescrittoriValue" runat="server" class="value"></span>
            </ItemTemplate>
        </nextFramework:NextList>
    </div>
    
     <div id="TitoloTrasportoDiv" runat="server" class="sezione">
        <h2 id="TitoloTrasporto" runat="server"></h2>
    </div>
    <div id="CostoPresaDiv" class="prezzo" runat="server">
        <span id="CostoPresaLabel" runat="server" class="label"></span>
        <span id="CostoPresaValue" runat="server" class="value"></span>
    </div>
    <div id="NumeroDdtCaricoDiv" class="numero" runat="server">
        <span id="NumeroDdtCaricoLabel" runat="server" class="label"></span>
        <span id="NumeroDdtCaricoValue" runat="server" class="value"></span>
    </div>
    <div id="DataDdtCaricoDiv" class="data" runat="server">
        <span id="DataDdtCaricoLabel" runat="server" class="label"></span>
        <span id="DataDdtCaricoValue" runat="server" class="value"></span>
    </div>
    <div id="CostoRiconsegnaDiv" class="prezzo" runat="server">
        <span id="CostoRiconsegnaLabel" runat="server" class="label"></span>
        <span id="CostoRiconsegnaValue" runat="server" class="value"></span>
    </div>
    <div id="NumeroDdtRiconsegnaDiv" class="numero" runat="server">
        <span id="NumeroDdtRiconsegnaLabel" runat="server" class="label"></span>
        <span id="NumeroDdtRiconsegnaValue" runat="server" class="value"></span>
    </div>
    <div id="DataDdtRiconsegnaDiv" class="data" runat="server">
        <span id="DataDdtRiconsegnaLabel" runat="server" class="label"></span>
        <span id="DataDdtRiconsegnaValue" runat="server" class="value"></span>
    </div>
    <div id="TrasportatoreDiv" class="trasportatore" runat="server">
        <span id="TrasportatoreLabel" runat="server" class="label"></span>
        <span id="TrasportatoreValue" runat="server" class="value"></span>
    </div>
    
    <div id="TitoloRiepilogoDiv" runat="server" class="sezione">
        <h2 id="TitoloRiepilogo" runat="server"></h2>
    </div>
    <div id="CostoPresaRiconsegnaDiv" class="prezzo" runat="server">
        <span id="CostoPresaRiconsegnaLabel" runat="server" class="label"></span>
        <span id="CostoPresaRiconsegnaValue" runat="server" class="value"></span>
    </div>
    <div id="CostoManodoperaDiv" class="prezzo" runat="server">
        <span id="CostoManodoperaLabel" runat="server" class="label"></span>
        <span id="CostoManodoperaValue" runat="server" class="value"></span>
    </div>
    <div id="CostoRicambiDiv" class="prezzo" runat="server">
        <span id="CostoRicambiLabel" runat="server" class="label"></span>
        <span id="CostoRicambiValue" runat="server" class="value"></span>
    </div>
    <div id="CostoTotaleDiv" class="prezzo" runat="server">
        <span id="CostoTotaleLabel" runat="server" class="label"></span>
        <span id="CostoTotaleValue" runat="server" class="value"></span>
    </div>
    
    <div id="InviaDiv" class="button" runat="server">
        <asp:Button OnClick="Invia_Click" ID="Invia" runat="server" CssClass="invia" />
    </div>
    <div id="IndietroDiv" class="button" runat="server">
        <asp:Button OnClick="Indietro_Click" ID="Indietro" runat="server" CssClass="indietro" />
    </div>
</div>
