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
dicitura.sottosezioni(2) = "STATISTICHE INDICE"
dicitura.links(2) = "SitoAnalisiStatIndice.asp"
dicitura.sottosezioni(3) = "STORICO PAGINE"
dicitura.links(3) = "SitoAnalisiStoricoPagine.asp"

dicitura.sezione = "Analisi storica accessi all'indice"

dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoAnalisi.asp"
dicitura.scrivi_con_sottosez()

dim conn, sql, rs, rss, rsp, lingua, i, ids
dim totUtenti, totCrawler, totAltro, totCont


set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rss = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")

'gestione dello scroll precedente/successivo tra i report
if request("GOTO") = "PREVIOUS" then
	if not isDate(session("st_data_from")) then
		session("ERRORE") = "Nessun report precedente trovato!"
	else
		session("st_data_to") = session("st_data_from")
		session("st_data_from") = getValueList(conn, rs, "SELECT MAX(sw_data) FROM tb_storico_webs WHERE sw_data < "& SQL_Date(conn, session("st_data_from")))
	end if
elseif request("GOTO") = "NEXT" then
	if cDate(session("st_data_to")) = cDate(getValueList(conn, rs, "SELECT MAX(sw_data) FROM tb_storico_webs")) then
		session("ERRORE") = "Nessun report successivo trovato!"
	else
		session("st_data_from") = session("st_data_to")
		session("st_data_to") = getValueList(conn, rs, "SELECT MIN(sw_data) FROM tb_storico_webs WHERE sw_data > "& SQL_Date(conn, session("st_data_to")))
	end if
end if

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	CALL SearchSession_Reset("st_")
	if request("cerca")<>"" then
		CALL SearchSession_Set("st_")
	end if
end if

'filtra per data
if isDate(Session("st_data_from")) then
	sql = sql & " AND sw_data > "& SQL_DateTime(conn, Session("st_data_from"))
end if
if isDate(Session("st_data_to")) then
	sql = sql & " AND sw_data <= "& SQL_DateTime(conn, Session("st_data_to"))
end if

if sql = "" then		'se non ho ricerca
	'prendo tutto lo storico
	Session("st_data_to") = getValueList(conn, rs, "SELECT MAX(sw_data) FROM tb_storico_webs")
	Session("st_data_from") = ""
end if

rs.open " SELECT * FROM tb_storico_webs WHERE sw_webs_id="& Session("AZ_ID") &" "& sql &" ORDER BY sw_data", conn, adOpenStatic, adLockOptimistic, adAsyncFetch
if not rs.eof then
	'calcolo la stringa degli id dello storico per il filtro delle pagine: IN (ids)
	while not rs.eof
		ids = ids & rs("sw_id") &","
		rs.moveNext
	wend
	ids = left(ids, len(ids)-1)
	rs.close
	
	sql = " SELECT SUM(sw_contatore) AS contatore, SUM(sw_contUtenti) AS contUtenti, SUM(sw_contCrawler) AS contCrawler, SUM(sw_contAltro) AS contAltro "& _
		  " FROM tb_storico_webs WHERE sw_webs_id="& Session("AZ_ID") &" "& sql
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adAsyncFetch
end if
%>

<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="27%" valign="top">
<!-- BLOCCO DI RICERCA -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
				<form action="" method="post" id="ricerca" name="ricerca">
						<tr>
							<td>
								<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
									<caption>Opzioni di ricerca</caption>
									<tr>
										<td class="footer" colspan="2">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 99%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("st_data_from;st_data_to") %>>DATA ARCHIVIAZIONE</td></tr>
									<tr><td class="label" colspan="2">a partire dal:</td></tr>
									<tr>
										<td class="content" colspan="2">
											<% sql = "SELECT sw_data FROM tb_storico_webs WHERE sw_data < (SELECT MAX(sw_data) FROM tb_storico_webs) ORDER BY sw_data"
											rss.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText%>
											<select name="search_data_from">
												<option value="">prima registrazione</option>
												<% while not rss.eof %>
													<option value="<%= rss("sw_data") %>" <%= IIF(cString(rss("sw_data"))=cString(session("st_data_from")), "selected", "") %>><%= rss("sw_data") %></option>
													<% rss.moveNext
												wend %>
											</select>
											<% rss.close %>
										</td>
									</tr>
									<tr><td class="label" colspan="2">fino al:</td></tr>
										<td class="content" colspan="2">
											<% sql = "SELECT sw_data FROM tb_storico_webs ORDER BY sw_data"
											rss.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText%>
											<select name="search_data_to">
												<%while not rss.eof %>
													<option value="<%= rss("sw_data") %>" <%= IIF(cString(rss("sw_data"))=cString(session("st_data_to")), "selected", "") %>><%= rss("sw_data") %></option>
													<% rss.moveNext
												wend %>
											</select>
											<% rss.close %>
										</td>
									</tr>
									<tr>
										<td class="footer" colspan="2">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 99%;">
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</form>
				</table>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
				<% if rs.eof then %>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
						<caption class="border">Nessuna statistica trovata per il periodo scelto.</caption>
					</table>
				<% else %>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
						<caption class="border">Statistiche generali del sito:</caption>
						<tr>
							<td class="label" style="width:30%;" rowspan="4">N&ordm; visitatori dal</td>
							<td class="label" style="width:20%;">utenti:</td>
							<td class="content"><%= cIntero(rs("contUtenti")) %></td>
						</tr>
						<tr>
							<td class="label">motori di ricerca:</td>
							<td class="content"><%= cIntero(rs("contCrawler")) %></td>
						</tr>
							<td class="label">altri visitatori:</td>
							<td class="content"><%= cIntero(rs("contAltro")) %></td>
						</tr>
						<tr>
							<td class="label">totale:</td>
							<td class="content"><%= cIntero(rs("contatore")) %></td>
						</tr>
					</table>
					<% dim tree
					set tree = new ObjJsTree
					tree.Name = "IndexAlbero_"
			        tree.TableCaption = "Statistiche generali del sito - albero"
					
					CALL  GetStatTable()
					
					CALL Tree_Explore(0)
					
					'disegna albero
			        CALL tree.Write()
					
					
					dim oLegenda
				    set oLegenda = new ObjIndexLegend
				    set oLegenda.Index = Index
				    oLegenda.Write() 

			end if			'fine se not rs.eof %>
		</td> 
	</tr>
	<tr><td>&nbsp;</td></tr>
</table>
</div>
</html>
<%
rs.close
conn.close
set rs = nothing
set rss = nothing
set rsp = nothing
set conn = nothing

'esplora l'albero dello storico
Sub Tree_Explore(padre_id)
	dim rst, sql, nome
    
	sql = " SELECT si_titolo_it, tab_name, tab_titolo, tab_colore,"& _
		  " si_idx_id, si_idx_livello, si_idx_foglia, si_co_F_key_id,"& _
		  " SUM(si_contUtenti) AS contUtenti, SUM(si_contCrawler) AS contCrawler, SUM(si_contAltro) AS contAltro, SUM(si_contatore) AS contatore"& _
		  " FROM tb_storico_index i" + _
		  " LEFT JOIN tb_siti_tabelle t ON i.si_tab_id = t.tab_id"& _
          " WHERE "& SQL_IfIsNull(conn, "si_idx_padre_id", "0") & "=" & cIntero(padre_id) & _
		  " AND si_sw_id IN ("& ids &")"& _
		  " GROUP BY si_titolo_it, tab_name, tab_titolo, tab_colore,"& _
		  " si_idx_id, si_idx_padre_id, si_idx_livello, si_idx_foglia, si_co_F_key_id, si_idx_ordine_assoluto"& _
          " ORDER BY si_idx_ordine_assoluto, si_titolo_it"
    set rst = conn.execute(Sql, ,adCmdtext)
    
    while not rst.Eof
    	nome = JSEncode(rst("si_titolo_it"), """")
        
		if CString(rst("tab_titolo")) <> "" then
			nome = nome & " ("& rst("tab_titolo") &")"
		end if
        'evidenzio il colore del tipo
		if CString(rst("tab_colore")) <> "" then
			nome = "<span style='color:"& rst("tab_colore") &"'>"& nome &"</span>"
		end if
        
		if LCase(rst("tab_name")) = "tb_webs" then
			dim rsAux
			set rsAux = conn.Execute("SELECT * FROM tb_storico_webs WHERE sw_webs_id = "& rst("si_co_F_key_id"))
			nome = GetStatNome(nome, rst("tab_name"), rst("si_idx_id"), rsAux("sw_contUtenti"), rsAux("sw_contCrawler"), rsAux("sw_contAltro"), rsAux("sw_contatore"))
			set rsAux = nothing
		else
			nome = GetStatNome(nome, rst("tab_name"), rst("si_idx_id"), rst("contUtenti"), rst("contCrawler"), rst("contAltro"), rst("contatore"))
		end if
		
        if rst("si_idx_foglia") then 
        	CALL tree.AddLeaf(rst("si_idx_livello"), nome, "", "")
        else
        	CALL tree.AddNode(rst("si_idx_livello"), nome, "", "", rst("si_idx_id"))
        end if
        
        if tree.IsNodeExpanded(rst("si_idx_id")) then
        	CALL Tree_Explore(rst("si_idx_id"))
        end if
		
        'visualizzo un nodo figlio qualsiasi altrimenti non appare il +
        if not rst("si_idx_foglia") AND NOT tree.IsNodeExpanded(rst("si_idx_id")) then
			CALL tree.AddNodeNew(rst("si_idx_livello") + 1, rst("si_idx_id"), "NODO DA ESPANDERE", "NODO DA ESPANDERE")
        end if
		
        rst.MoveNext
	wend
	
    rst.close
    set rst = nothing
end sub %>