<%@ Control Language="C#" AutoEventWireup="true" CodeFile="RichiestaRitiro.ascx.cs" Inherits="Plugin_RichiestaRitiro" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="nextFramework" %>

<a class="print_link" href="javascript:window.print()">STAMPA</a>
<div id="DivTestata" class="top" runat="server">
    <div id="TrasportatoreDiv" class="trasportatore" runat="server"></div>
    <div id="DescrizioneDiv" class="avviso" runat="server"></div>
    <div id="MittenteDiv" class="mittente" runat="server"></div>
</div>
<div id="DivSx" class="sx" runat="server">
    <div id="DestinatarioDiv" class="destinatario" runat="server"></div>
    <div id="DataRitiroDiv" class="dataritiro" runat="server"></div>
</div>
<div id="DivDx" class="dx" runat="server">
    <div id="DatiDiv" class="dati" runat="server"></div>
    <div id="ConsegnaDiv" class="consegna" runat="server"></div>
</div>
