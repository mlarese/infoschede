<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione marchi / produttori - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVO MARCHIO"
dicitura.link_new = "Tabelle.asp;MarchiNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql, pager, i, sql_filtri

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("marchi_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("marchi_")
	end if
end if

sql = ""
'filtra per codice
if Session("marchi_codice")<>"" then
    sql = sql & " AND (" & SQL_FullTextSearch(Session("marchi_codice"), "mar_codice") & ") "
end if

'filtra per nome
if Session("marchi_nome")<>"" then
	sql = sql & " AND (" & SQL_FullTextSearch(Session("marchi_nome"), FieldLanguageList("mar_nome_")) & ") "
end if

'filtra per nome cliente
if Session("marchi_denominazione")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("marchi_denominazione"))
end if

'filtra per descrizione
if Session("marchi_descrizione")<>"" then
	sql = sql & " AND (" & SQL_FullTextSearch(Session("marchi_descrizione"), FieldLanguageList("mar_descr_")) & ") "
end if

sql_filtri = sql

sql = "SELECT *, (SELECT COUNT(*) FROM gtb_articoli WHERE art_marca_id=mar_id) AS N_ART FROM gtb_marche " 
if Session("marchi_denominazione") <> "" then
	sql = sql + " LEFT JOIN gv_rivenditori ON gtb_marche.mar_anagrafica_id = gv_rivenditori.riv_id "
end if
sql = sql + "  WHERE (1=1) " + sql_filtri + " ORDER BY mar_nome_it"



session("B2B_MARCHI_SQL") = sql
CALL Pager.OpenSmartRecordset(conn, rs, sql, 20)
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
										<td class="footer">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTI" style="width: 49%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("marchi_codice") %>>CODICE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_codice" value="<%= TextEncode(session("marchi_codice")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("marchi_nome") %>>NOME</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_nome" value="<%= TextEncode(session("marchi_nome")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("marchi_denominazione") %>>COSTRUTTORE</th></tr>
									<tr>
										<td class="content" colspan="2">
										<input type="text" name="search_denominazione" value="<%= TextEncode(session("marchi_denominazione")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th <%= Search_Bg("marchi_descrizione") %>>DESCRIZONE</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_descrizione" value="<%= TextEncode(session("marchi_descrizione")) %>" style="width:100%;">
										</td>
									</tr>
									<tr>
										<td class="footer">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_bottom" value="VEDI TUTTI" style="width: 49%;">
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
				<!-- BLOCCO RISULTATI -->
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>Elenco marchi - Trovati n&ordm; <%= Pager.recordcount %> marchi in n&ordm; <%= Pager.PageCount %> pagine</caption>
					<% if not rs.eof then %>
						<tr>
							<th style="width: 10%;">CODICE</th>
							<th>NOME</th>
							<th class="center" style="width: 16%;">LOGO</th>
							<th class="center" colspan="2" style="width: 25%;">OPERAZIONI</th>
						</tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="content"><%= rs("mar_codice") %></td>
								<td class="content"><%= rs("mar_nome_it") %></td>
								<td class="Content_center">
									<% 	if CString(rs("mar_logo")) <> "" then %>
										<img src="http://<%= Application("IMAGE_SERVER") &"/"& Application("AZ_ID") &"/images/"& rs("mar_logo") %>" alt="logo" border="0">
									<% 	else %>
										&nbsp;
									<% 	end if %>
								</td>
								<td class="content_center">
									<table>
										<tr>
											<td style="vertical-align:middle;" class="content_center" colspan="2">
												<% CALL index.WriteButton("gtb_marche", rs("mar_id"), POS_ELENCO) %>
											</td>
										</tr>
										<tr>
											<td style="vertical-align:middle;" class="content_center">
												<a class="button" href="MarchiMod.asp?ID=<%= rs("mar_id") %>">
													MODIFICA
												</a>
											</td>
											<td style="vertical-align:middle;" class="content_center">
												<% if rs("N_ART") > 0 then %>
													<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare il marchio: sono presenti articoli associati">
														CANCELLA
													</a>
												<% else %>
													<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('MARCHE','<%= rs("mar_id") %>');" >
														CANCELLA
													</a>
												<% end if %>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<%rs.movenext
						wend%>
						<tr>
							<td class="footer" style="text-align:left;" colspan="6">
								<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
							</td>
						</tr>
					<%else%>
						<tr><td class="noRecords">Nessun record trovato</th></tr>
					<% end if %>
				</table>
			</td> 
		</tr>
		<tr><td>&nbsp;</td></tr>
	</table>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set conn = nothing%>