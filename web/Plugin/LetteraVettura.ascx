<%@ Control Language="C#" AutoEventWireup="true" CodeFile="LetteraVettura.ascx.cs" Inherits="Plugin_LetteraVettura" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="nextFramework" %>

<div id="DivSx" class="sx" runat="server">
    <div id="MittenteDiv" class="mittente" runat="server"></div>
    <div id="DestinatarioDiv" class="destinatario" runat="server"></div>
</div>
<div id="DivDx" class="dx" runat="server">
    <div id="TrasportatoreDiv" class="trasportatore" runat="server"></div>
    <div id="DatiDiv" class="dati" runat="server"></div>
</div>
<div id="DivFirme" class="bottom" runat="server">
    <div id="FirmaMittenteDiv" class="sx firma" runat="server"></div>
    <div id="FirmaRitiroDiv" class="dx firma" runat="server"></div>
</div>
