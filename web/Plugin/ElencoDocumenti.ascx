<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ElencoDocumenti.ascx.cs" Inherits="Plugin_B2B_ElencoDocumenti" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="cc1" %>

<div class="magiclist">
    <cc1:Magiclist ID="Catalogo" runat="server" OnMagicboxDataBinding="Catalogo_MagicboxDataBinding">
    </cc1:Magiclist>
</div>
