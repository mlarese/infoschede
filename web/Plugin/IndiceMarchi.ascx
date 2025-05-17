<%@ Control Language="C#" AutoEventWireup="true" CodeFile="IndiceMarchi.ascx.cs" Inherits="Plugin_IndiceMarchi" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="cc1" %>

<h1 id="TitoloSez" class="titolo_sez" runat="server"></h1>

<cc1:NextList ID="ListaLettere" runat="server" CssClass="lista_lettere" OnItemDataBound="ListaLettere_ItemDataBound">

    <ItemTemplate>
        <div id="divIniziale" runat="server" class="iniziale">
            <label id="iniziale" runat="server" class="iniziale"></label>
        </div>
        
        <cc1:NextList ID="ListaMarchi" runat="server" CssClass="lista_marchi" ColumnsNumber="4" OnItemDataBound="ListaMarchi_ItemDataBound">

            <ItemTemplate>                        
                
                    <div class="img" id="ImgCont" runat="server">
                        <a id="marchio" runat="server" class="marchio">
                            <img src="" id="logo" runat="server" alt="" />
                            <h1 id="Titolo" runat="server"></h1> 
                        </a>    
                    </div>                     
                
            </ItemTemplate>
        </cc1:NextList>
        
    </ItemTemplate>
</cc1:NextList>