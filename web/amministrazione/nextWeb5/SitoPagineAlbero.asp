<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/ClassIndexAlberi.asp" -->
<%
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - indice delle pagine - albero"
dicitura.puls_new = "NUOVA PAGINA"
dicitura.link_new = "SitoPagineNew.asp?FROM=" & FROM_ALBERO
dicitura.scrivi_con_sottosez()
%>
<div id="content">
    <table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
	    <caption class="border">Indice delle pagine - elenco</caption>
        <tr>
            <td class="content">Visualizza l'indice delle pagine come elenco </td>
            <td class="content_right">
                <a class="button" href="SitoPagine.asp" title="Apre la visualizzazione come elenco.">
				    VISUALIZZA COME ELENCO
                </a>
            </td>
        </tr>
    </table>
    <%
    dim oTree 
    set oTree = new ObjIndexTrees
    set oTree.Index = Index
    CALL oTree.AlberoIndiceByTipoContenuto("tb_pagineSito", _
                                           "SitoPagineNew.asp?FROM="& FROM_ALBERO &"&INDICENEW=", _
                                           "nuova pagina", _
                                           "SitoPagineMod.asp?FROM="& FROM_ALBERO &"&INDICE=")
    
    dim oLegenda
    set oLegenda = new ObjIndexLegend
    set oLegenda.Index = Index
    oLegenda.ContentTabName = "tb_pagineSito"
    CALL oLegenda.AddExtra("#666", "Pagine di contenuto")
    oLegenda.Write()
    
    %>
</div>
</body>
</html>