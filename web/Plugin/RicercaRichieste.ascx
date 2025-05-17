<%@ Control Language="C#" AutoEventWireup="true" CodeFile="RicercaRichieste.ascx.cs" Inherits="Plugin_RicercaRichieste" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="nextFramework" %>

<h1 id="Titolo" class="titolo" runat="server"></h1>
<div id="DescrizioneDiv" class="descrizione" runat="server" visible="false">
    <p id="Descrizione" runat="server"></p>
</div>
<div id="ErroriListaDiv" class="errori" runat="server" visible="false"></div>

<div id="RicercaDiv" runat="server" class="ricerca">
    <div id="NumeroSchedaDiv" class="stato" runat="server">
        <span id="NumeroSchedaLabel" runat="server" class="label"></span>
        <input id="NumeroSchedaInput" runat="server" class="input" maxlength="10" />
        <asp:RegularExpressionValidator id="NumeroSchedaInputValid" runat="server" EnableClientScript="false"
                                        ErrorMessage="Inserire solo cifre" ControlToValidate="NumeroSchedaInput"
                                        ValidationExpression="[0-9-]+" Text="<--">
        </asp:RegularExpressionValidator>
    </div>
    <div id="StatoDiv" class="stato" runat="server">
        <span id="StatoLabel" runat="server" class="label"></span>
        <asp:DropDownList ID="StatoDropDown" cssclass="input lungo" runat="server"></asp:DropDownList>
    </div>
    <div id="MarcaDiv" class="marca" runat="server">
        <span id="MarcaLabel" runat="server" class="label"></span>
        <asp:DropDownList ID="MarcaDropDown" cssclass="input lungo" runat="server"></asp:DropDownList>
    </div>
    <div id="ModelloDiv" class="stato" runat="server">
        <span id="ModelloLabel" runat="server" class="label"></span>
        <input id="ModelloInput" runat="server" class="input lungo" maxlength="250" />
        <asp:RegularExpressionValidator id="ModelloInputValid" runat="server" EnableClientScript="false"
                                        ErrorMessage="Inserire solo lettere o cifre" ControlToValidate="ModelloInput"
                                        ValidationExpression="[A-Za-z0-9- ]+" Text="<--">
        </asp:RegularExpressionValidator>
    </div>
    <div id="RifClienteDiv" class="stato" runat="server">
        <span id="RifClienteLabel" runat="server" class="label"></span>
        <input id="RifClienteInput" runat="server" class="input lungo" maxlength="250" />
        <asp:RegularExpressionValidator id="RifClienteInputValid" runat="server" EnableClientScript="false"
                                        ErrorMessage="Inserire solo lettere o cifre" ControlToValidate="RifClienteInput"
                                        ValidationExpression="[A-Za-z0-9- ]+" Text="<--">
        </asp:RegularExpressionValidator>
    </div>
    <div id="NumeroDdtCaricoDiv" class="stato" runat="server">
        <span id="NumeroDdtCaricoLabel" runat="server" class="label"></span>
        <input id="NumeroDdtCaricoInput" runat="server" class="input" maxlength="10" />
        <asp:RegularExpressionValidator id="NumeroDdtCaricoInputValid" runat="server" EnableClientScript="false"
                                        ErrorMessage="Inserire solo cifre" ControlToValidate="NumeroDdtCaricoInput"
                                        ValidationExpression="[A-Za-z0-9- ]+" Text="<--">
        </asp:RegularExpressionValidator>
    </div>
    <div id="RivenditoreDiv" class="rivenditore" runat="server">
        <span id="RivenditoreLabel" runat="server" class="label"></span>
        <asp:DropDownList ID="RivenditoreDropDown" cssclass="input lungo" runat="server"></asp:DropDownList>
    </div>
    <div id="DataRichiestaDaDiv" class="stato" runat="server">
        <div class="dp_container">
            <span id="DataRichiestaDaLabel" class="label" runat="server"></span>
            <input class="datepicker input" id="DataRichiestaDaInput" runat="server" type="text" />
        </div>
                
    </div>
    <div id="DataRichiestaADiv" class="stato" runat="server">
        <div class="dp_container">
            <span id="DataRichiestaALabel" class="label" runat="server"></span>
            <input class="datepicker input" id="DataRichiestaAInput" runat="server" type="text" />
        </div>
    </div>
    
</div>

<div id="CercaDiv" class="button" runat="server">
    <asp:Button OnClick="Cerca_Click" ID="Cerca" runat="server" CssClass="cerca" CausesValidation="true" />
</div>
<div id="ViewAllDiv" class="button indietro" runat="server">
    <asp:Button OnClick="Cerca_Click" ID="ViewAll" runat="server" CssClass="cerca" CausesValidation="false" />
</div>
