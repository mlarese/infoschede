<%@ Control Language="C#" AutoEventWireup="true" CodeFile="RicercaProdotti.ascx.cs" Inherits="Plugin_RicercaProdotti" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="nextFramework" %>

<h1 id="Titolo" class="titolo" runat="server"></h1>
<div id="DescrizioneDiv" class="descrizione" runat="server">
    <p id="Descrizione" runat="server"></p>
</div>
<div id="ErroriListaDiv" class="errori" runat="server" visible="false"></div>

<div id="RicercaDiv" runat="server" class="ricerca">
    <div id="ProdottoDiv" class="prodotto" runat="server">
        <span id="ProdottoLabel" runat="server" class="label" visible="false"></span>
        <input id="ProdottoInput" runat="server" class="input lungo" maxlength="250" />
        <asp:RegularExpressionValidator id="ProdottoInputValid" runat="server" EnableClientScript="false"
                                        ErrorMessage="Inserire solo lettere o cifre" ControlToValidate="ProdottoInput"
                                        ValidationExpression="[A-Za-z0-9- ]+" Text="<--">
        </asp:RegularExpressionValidator>
    </div>
</div>
<div id="CercaDiv" class="button" runat="server">
    <asp:Button OnClick="Cerca_Click" ID="Cerca" runat="server" CssClass="cerca" CausesValidation="true" />
</div>
<div id="ViewAllDiv" class="button indietro" runat="server">
    <asp:Button OnClick="Cerca_Click" ID="ViewAll" runat="server" CssClass="cerca" CausesValidation="false" />
</div>
