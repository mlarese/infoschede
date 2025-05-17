<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/ClassIndexAlberi.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_indice_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.sezione = "Indice generale"
dicitura.puls_new = "nuovo:;RAGGRUPPAMENTO;VOCE"
dicitura.link_new = ";IndexRaggruppamentoGestione.asp?FROM="& FROM_ALBERO &";IndexGestione.asp?FROM=" & FROM_ALBERO

if index.ChkPrm(prm_Pubblicazioni_accesso, 0) then
	dicitura.iniz_sottosez(2)
	dicitura.sottosezioni(2) = "PUBBLICAZIONI AUTOMATICHE"
	dicitura.links(2) = "IndexPubblicazioni.asp"
else
	dicitura.iniz_sottosez(1)
end if
dicitura.sottosezioni(1) = "META TAG"
dicitura.links(1) = "IndexMetaTag.asp?FROM=Indice"

dicitura.scrivi_con_sottosez()
%>
<div id="content">
    <table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
        <caption class="border">Indice generale - elenco</caption>
        <tr>
            <td class="content">Visualizza l'indice generale come elenco</td>
            <td class="content_right">
                <a class="button" href="IndexGenerale.asp" title="Apre la visualizzazione come elenco.">
                    VISUALIZZA COME ELENCO
                </a>
            </td>
        </tr>
    </table>
    <%
    dim oTree
    set oTree = new ObjIndexTrees
    set oTree.Index = Index
    CALL oTree.AlberoIndiceCompleto()
    
    
    dim oLegenda
    set oLegenda = new ObjIndexLegend
    set oLegenda.Index = Index
    oLegenda.Write()
    %>
</div>
</body>
</html>
