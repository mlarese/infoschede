<%@ Control Language="C#" AutoEventWireup="true" CodeFile="InviaEmail.ascx.cs" Inherits="Plugin_InviaEmail" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="nextFramework" %>

<div id="EmailDiv" runat="server" class="email">
    <h1 id="Titolo" class="titolo" runat="server"></h1>
    <div id="DescrizioneDiv" class="descrizione" runat="server" visible="false">
        <p id="Descrizione" runat="server"></p>
    </div>

    <div id="DatiDiv" class="dati" runat="server">
        <div id="NumeroSchedaDiv" class="numero" runat="server">
            <span id="NumeroSchedaLabel" runat="server" class="label"></span>
            <span id="NumeroSchedaValue" runat="server" class="value"></span>
        </div>
        <div id="DataRicevimentoDiv" class="data" runat="server">
            <span id="DataRicevimentoLabel" runat="server" class="label"></span>
            <span id="DataRicevimentoValue" runat="server" class="value"></span>
        </div>
        <div id="ClienteDiv" class="cliente" runat="server">
            <span id="ClienteLabel" runat="server" class="label"></span>
            <span id="ClienteValue" runat="server" class="value"></span>
        </div>
        <div id="EmailClienteDiv" class="cliente" runat="server">
            <span id="EmailClienteLabel" runat="server" class="label"></span>
            <span id="EmailClienteValue" runat="server" class="value"></span>
        </div>
        <div id="EmailAllegatoDiv" class="allegato" runat="server">
            <span id="EmailAllegatoLabel" runat="server" class="label"></span>
            <a id="EmailAllegato" runat="server" class="link" target="_blank"></a>
        </div>
        <div id="EmailTestoDiv" class="testo" runat="server">
            <span id="EmailTestoLabel" runat="server" class="label"></span>
            <textarea id="EmailTesto" runat="server" class="input" rows="10" cols="73"></textarea>
        </div>
    </div>
    <div class="avviso" id="AvvisoInvioInCopia" runat="server" visible="false">Attenzione! L'e-mail verrà mandata in copia anche al cliente.</div>
    <div id="InviaDiv" class="button" runat="server">
        <asp:Button OnClick="Invia_Click" ID="Invia" runat="server" />
        <asp:Button ID="Chiudi" runat="server" Text="Chiudi" visible="false" OnClientClick="javascript:window.close();" />
    </div>
</div>
<script type="text/javascript" language="JavaScript">
    window.resizeTo(600,500);
</script>
