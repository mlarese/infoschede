<%@ Language=VBScript CODEPAGE=65001%>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1000 %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="SitoAnalisiStat_TOOLS.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/ClassIndexAlberi.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_strumenti_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(3)
dicitura.sottosezioni(1) = "STATISTICHE PAGINE"
dicitura.links(1) = "SitoAnalisiStatPagine.asp"
dicitura.sottosezioni(2) = "STORICO PAGINE"
dicitura.links(2) = "SitoAnalisiStoricoPagine.asp"
dicitura.sottosezioni(3) = "STORICO INDICE"
dicitura.links(3) = "SitoAnalisiStoricoIndice.asp"

dicitura.sezione = "Analisi statistica accessi all'indice"

dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoAnalisi.asp"
dicitura.scrivi_con_sottosez()

dim conn, sql, rs, rsp, lingua, i

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet") %>
<div id="content">
	<% CALL WRITE_StatisticheGenerali(conn, rs, Session("AZ_ID"))
	
	dim oTree
    set oTree = new ObjIndexTrees
	oTree.Tree.TableCaption = "Statistiche di accesso ai nodi dell'indice"
    set oTree.Index = Index
	oTree.GestioneEsterna = true
    
	CALL GetStatTable()
	
    CALL oTree.AlberoIndiceByTipoContenuto("", "", "", "")
	
    dim oLegenda
    set oLegenda = new ObjIndexLegend
    set oLegenda.Index = Index
    oLegenda.Write() %>

</div>
</body>
</html>

<%	
'gestisce il nome del nodo
Sub GestioneNodo(conn, nodo, ByRef nome, ByRef link)
	dim rs, table
	'se sono il sito visualizzo le relative statistiche
	if instr(1, nodo("tab_name"), "tb_webs", vbTextCompare)>0 then
		set rs = conn.Execute("SELECT * FROM tb_webs WHERE id_webs = "& nodo("co_F_key_id"))
		nome = GetStatNome(nome, nodo("tab_name"), nodo("idx_id"), rs("contUtenti"), rs("contCrawler"), rs("contAltro"), rs("contatore"))
	else
		nome = GetStatNome(nome, nodo("tab_name"), nodo("idx_id"), nodo("idx_contUtenti"), nodo("idx_contCrawler"), nodo("idx_contAltro"), nodo("idx_contatore"))
	end if
	set rs = nothing
End Sub

%>