<%@ Language=VBScript CODEPAGE=65001%>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1000 %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="SitoAnalisiStat_TOOLS.asp" -->
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
dicitura.sottosezioni(3) = "STORICO INDICE"
dicitura.links(3) = "SitoAnalisiStoricoIndice.asp"

dicitura.sezione = "Storico statistiche accesso pagine"

dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoAnalisi.asp"
dicitura.scrivi_con_sottosez()

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

dim conn, sql, rs, rss, rsp, lingua, i, ids
dim totUtenti, totCrawler, totAltro, totCont

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rss = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")


'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	if request("tutti")<>"" then
		CALL SearchSession_Reset("st_")
	elseif request("cerca")<>"" then
		CALL SearchSession_Reset("st_")
		CALL SearchSession_Set("st_")
	elseif request("elimina")<>"" then
		'abilitato solo per next-aim: cancella registrazioni dello storico
		if instr(1, Session("WEB_ADMIN"), "NEXT", vbTextCompare)>0 AND _
		   instr(1, Session("WEB_ADMIN"), "AIM", vbTextCompare)>0 AND _
		   cInteger(request("elimina_registrazione"))>0 then
		   	sql = "DELETE FROM tb_storico_webs WHERE sw_id=" & cInteger(request("elimina_registrazione"))
			CALL conn.execute(sql, , adExecuteNoRecords)
			response.redirect GetPageName()
		end if
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

rs.open " SELECT * FROM tb_storico_webs WHERE sw_webs_id="& Session("AZ_ID") &" "& sql &" ORDER BY sw_data ", conn, adOpenStatic, adLockOptimistic, adAsyncFetch
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
								<table cellspacing="1" cellpadding="0" class="tabella_madre">
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
								<% if instr(1, Session("WEB_ADMIN"), "NEXT", vbTextCompare)>0 AND _
									  instr(1, Session("WEB_ADMIN"), "AIM", vbTextCompare)>0 then %>
									<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-top:30px;">
										<caption>Cancellazione storico</caption>
										<tr><th colspan="2">REGISTRAZIONE DA RIMUOVERE</td></tr>
										<tr>
											<td class="note">Funzione abilitata solo per l'utente NEXTAIM</td>
										</tr>
										<tr>
											<td class="content">
												<% sql = "SELECT * FROM tb_storico_webs ORDER BY sw_data"
												CALL dropDown(conn, sql, "sw_id", "sw_data", "elimina_registrazione", "", false, " style=""width:100%;""", LINGUA_ITALIANO)  %>
											</td>
										</tr>
										<tr>
											<td class="footer" colspan="2">
												<input type="submit" name="elimina" value="CANCELLA REGISTRAZIONE" class="button" style="width: 99%;">
											</td>
										</tr>
									</table>
								<% end if %>
							</td>
						</tr>
					</form>
				</table>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
<!-- BLOCCO RISULTATI -->
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
				
					<table cellspacing="1" cellpadding="0" class="tabella_madre">
						<caption>Statistiche di visualizzazione delle pagine:</caption>
						<tr>
						<th class="center" rowspan="2" style="width:2%;">ID</th>
						<th rowspan="2" colspan="2">TITOLO</th>
						<th class="center" colspan="4" style="border-bottom:0px;">NUMERO VISITE DELLA PAGINA</th>
					</tr>
					<tr>
						<th class="right" style="width:8%;">UTENTI</th>
						<th class="right" style="width:22%;">MOTORI DI RICERCA</th>
						<th class="right" style="width:8%;">ALTRI</th>
						<th class="right" style="width:8%;">TOTALE</th>
					</tr>
					<% rs.close

					totUtenti = 0
					totCrawler = 0
					totAltro = 0

					'selezione le pagineSito per la visualizzazione raggruppata
					sql = " SELECT sp_pagineSito_id FROM tb_storico_pages WHERE sp_sw_id IN ("& ids &") GROUP BY sp_pagineSito_id"
					rs.open sql, conn, adOpenStatic, adLockOptimistic, adAsyncFetch
					'preparo la query per la scelta della pagina data la paginaSito e la lingua
					sql = " SELECT MIN(sp_nomepage) AS nomepage, SUM(sp_contatore) AS contatore, SUM(sp_contUtenti) AS contUtenti, "& _
						  " SUM(sp_contCrawler) AS contCrawler, SUM(sp_contAltro) AS contAltro "& _
						  " FROM tb_storico_pages WHERE sp_sw_id IN ("& ids &") "
					while not rs.eof %>
						<% for each lingua in application("LINGUE")
							if Session("LINGUA_" & lingua) then
								rsp.open sql &" AND sp_lingua='"& lingua &"' AND sp_pagineSito_id="& rs("sp_pagineSito_id") & " GROUP BY sp_page_id", _
										 conn, adOpenStatic, adLockReadOnly, adCmdText
								if uCase(lingua) <> "IT" then %>
									<tr>
								<% else %>
									<td class="content_center" rowspan="<%= Session("LINGUE_ATTIVE") %>"><%= rs("sp_pagineSito_id") %></td>
								<% end if %>
								<td class="content_center" style="width:5%;">
									<img src="../grafica/flag_mini_<%= lingua %>.jpg" alt="" border="0">
								</td>
								<%if rsp.eof then 		'non c'e la lingua nella registrazione %>
									<td class="content alert" colspan="5">
										pagina non presente nell'archivio per la lingua corrente
									</td>
								<% else %>
									<td class="content"><%= rsp("nomepage") %></td>
									<td class="content_right<%= IIF(cIntero(rsp("contUtenti")) = 0, " notes", "") %>"><%= cIntero(rsp("contUtenti")) %></td>
									<td class="content_right<%= IIF(cIntero(rsp("contCrawler")) = 0, " notes", "") %>"><%= cIntero(rsp("contCrawler")) %></td>
									<td class="content_right<%= IIF(cIntero(rsp("contAltro")) = 0, " notes", "") %>"><%= cIntero(rsp("contAltro")) %></td>
									<td class="content_right<%= IIF(cIntero(rsp("contatore")) = 0, " notes", "") %>"><%= cIntero(rsp("contatore")) %></td>
								</tr>
									<% totUtenti = totUtenti + cIntero(rsp("contUtenti"))
									totCrawler = totCrawler + cIntero(rsp("contCrawler"))
									totAltro = totAltro + cIntero(rsp("contAltro"))
									totCont = totCont + cIntero(rsp("contatore"))
								end if
								rsp.close
							end if
						next
						rs.moveNext
					wend %>
					<tr>
						<td class="footer" colspan="3">
							totali pagine viste:
						</td>
						<td class="footer"><%= totUtenti %></td>
						<td class="footer"><%= totCrawler %></td>
						<td class="footer"><%= totAltro %></td>
						<td class="footer"><%= totCont %></td>
					</tr>
				</table>
			<% end if			'fine se not rs.eof %>
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
%>