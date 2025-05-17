<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ElencoRichieste.ascx.cs" Inherits="Plugin_ElencoRichieste" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="nextFramework" %>

<h1 id="Titolo" class="titolo_sez" runat="server"></h1>
<div id="DescrizioneDiv" class="descrizione" runat="server">
    <p id="Descrizione" runat="server"></p>
    <p id="Summary" runat="server"></p>
</div>

<div id="ErroriListaDiv" class="errori" runat="server" visible="false"></div>

<div id="ElencoDiv" runat="server" class="elenco">
    <asp:GridView id="RichiesteElenco" runat="server" CssClass="value" AutoGenerateColumns="false" AllowSorting="true"
                  AllowPaging="true" PageSize="15" CellPadding="6" CellSpacing="0" OnPageIndexChanging="GridView_PageChanging" 
                  OnRowDataBound="RichiesteElenco_RowDataBound" OnSorting="RichiesteElenco_Sorting">
        <Columns>
            <asp:TemplateField SortExpression="stato">
                <HeaderTemplate>
                    <asp:LinkButton id="statoHd" runat="server" CssClass="label" ToolTip="ordina per stato scheda"></asp:LinkButton>
                </HeaderTemplate>
                <ItemTemplate>
                    <span id="stato" runat="server" class="value"></span>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField SortExpression="sc_numero">
                <HeaderTemplate>
                    <asp:LinkButton id="numeroHd" runat="server" CssClass="label" ToolTip="ordina per numero scheda"></asp:LinkButton>
                </HeaderTemplate>
                <ItemTemplate>
                    <span id="numero" runat="server" class="value" style="text-align:right"></span>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField SortExpression="sc_data_ricevimento">
                <HeaderTemplate>
                    <asp:LinkButton id="dataRicevimentoHd" runat="server" CssClass="label" ToolTip="ordina per data ricevimento"></asp:LinkButton>
                </HeaderTemplate>
                <ItemTemplate>
                    <span id="dataRicevimento" runat="server" class="value" style="text-align:right"></span>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField SortExpression="sc_rif_cliente">
                <HeaderTemplate>
                    <asp:LinkButton id="riferimentoClienteHd" runat="server" CssClass="label" ToolTip="ordina per riferimento cliente"></asp:LinkButton>
                </HeaderTemplate>
                <ItemTemplate>
                    <span id="riferimentoCliente" runat="server" class="value"></span>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField SortExpression="nome_rivenditore">
                <HeaderTemplate>
                    <asp:LinkButton id="nomeRivenditoreHd" runat="server" CssClass="label" ToolTip="ordina per nome rivenditore"></asp:LinkButton>
                </HeaderTemplate>
                <ItemTemplate>
                    <span id="nomeRivenditore" runat="server" class="value"></span>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField SortExpression="modello">
                <HeaderTemplate>
                    <asp:LinkButton id="modelloHd" runat="server" CssClass="label" ToolTip="ordina per modello"></asp:LinkButton>
                </HeaderTemplate>
                <ItemTemplate>
                    <span id="modello" runat="server" class="value"></span>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField Visible="false">
                <HeaderTemplate>
                    <span id="guastoHd" runat="server" class="label"></span>
                </HeaderTemplate>
                <ItemTemplate>
                    <span id="guasto" runat="server" class="value"></span>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField Visible="false">
                <HeaderTemplate>
                    <span id="esitoInterventoHd" runat="server" class="label"></span>
                </HeaderTemplate>
                <ItemTemplate>
                    <span id="esitoIntervento" runat="server" class="value"></span>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField SortExpression="sc_numero_DDT_di_carico">
                <HeaderTemplate>
                    <asp:LinkButton id="numeroCaricoHd" runat="server" CssClass="label" ToolTip="ordina per numero DDT di carico"></asp:LinkButton>
                </HeaderTemplate>
                <ItemTemplate>
                    <span id="numeroCarico" runat="server" class="value" style="text-align:right"></span>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField SortExpression="sc_data_DDT_di_carico">
                <HeaderTemplate>
                    <asp:LinkButton id="dataCaricoHd" runat="server" CssClass="label" ToolTip="ordina per data DDT di carico"></asp:LinkButton>
                </HeaderTemplate>
                <ItemTemplate>
                    <span id="dataCarico" runat="server" class="value" style="text-align:right"></span>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderStyle-CssClass="link" ItemStyle-CssClass="link">
                <HeaderTemplate>
                    <span id="linkHd" runat="server" class="label"></span>
                </HeaderTemplate>
                <ItemTemplate>
                    <a id="link" runat="server" class="value"></a>
                </ItemTemplate>
            </asp:TemplateField>
        </Columns>
    </asp:GridView>
</div>
