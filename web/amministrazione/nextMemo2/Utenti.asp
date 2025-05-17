<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata  
dicitura.iniz_sottosez(1)
dicitura.sottosezioni(1) = "LOG ACCESSI"
dicitura.links(1) = "UtentiAccessi.asp"
dicitura.sezione = "Gestione utenti area risevata NextMemo 2.0 - elenco"
dicitura.puls_new = "NUOVO UTENTE"
dicitura.link_new = "UtentiNew.asp"
dicitura.scrivi_con_sottosez() 


dim conn, rs, sql, Pager, i, rsa, email, headerCss
set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("u_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("u_")
	end if
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")


if Session("ID_AREA_RISERVATA") <> "" then
	sql = " AND ut_id IN (SELECT rel_ut_id FROM rel_utenti_sito WHERE rel_sito_id IN (" & Session("ID_AREA_RISERVATA") & "))"
end if

if Session("PERMESSO_AREA_RISERVATA") <> "" then
	dim sql_aux
	sql_aux = " SELECT " & _
			  "  (CASE WHEN sito_p1 like '"&Session("PERMESSO_AREA_RISERVATA")&"' THEN 1 " & _
			  "  WHEN sito_p2 like '"&Session("PERMESSO_AREA_RISERVATA")&"' THEN 2 " & _
			  "  WHEN sito_p3 like '"&Session("PERMESSO_AREA_RISERVATA")&"' THEN 3 " & _
			  "  WHEN sito_p4 like '"&Session("PERMESSO_AREA_RISERVATA")&"' THEN 4 " & _
			  "  WHEN sito_p5 like '"&Session("PERMESSO_AREA_RISERVATA")&"' THEN 5 " & _
			  "  WHEN sito_p6 like '"&Session("PERMESSO_AREA_RISERVATA")&"' THEN 6 " & _
			  "  WHEN sito_p7 like '"&Session("PERMESSO_AREA_RISERVATA")&"' THEN 7 " & _
			  "  WHEN sito_p8 like '"&Session("PERMESSO_AREA_RISERVATA")&"' THEN 8 " & _
			  "  WHEN sito_p9 like '"&Session("PERMESSO_AREA_RISERVATA")&"' THEN 9 " & _
			  "  END) FROM tb_siti "
	if Session("ID_AREA_RISERVATA") <> "" then
		sql_aux = sql_aux & " WHERE id_sito = " & Session("ID_AREA_RISERVATA")
	else
		sql_aux = sql_aux & " WHERE sito_amministrazione = 0 "
	end if
	sql = sql & " AND ut_id IN (SELECT rel_ut_id FROM rel_utenti_sito WHERE rel_permesso =" & cIntero(GetValueList(conn, NULL, sql_aux)) & " ) "
end if

'filtra per login
if Session("u_login")<>"" then
    sql = sql & " AND "& SQL_FullTextSearch(Session("u_login"), "ut_login")
end if

'filtra per abilitazione dell'account utente
if Session("u_abilitato")<>"" then
	if not (instr(1, Session("u_abilitato"), "A", vbTextCompare)>0 AND instr(1, Session("u_abilitato"), "N", vbTextCompare)>0 ) then
		if instr(1, Session("u_abilitato"), "A", vbTextCompare)>0 then
			sql = sql & " AND " & SQL_IsTrue(conn, "UT_Abilitato")
		elseif instr(1, Session("u_abilitato"), "N", vbTextCompare)>0 then
			sql = sql & " AND " & " NOT (" & SQL_IsTrue(conn, "UT_Abilitato") & ") "
		end if
	end if
end if

'filtra per nome / cognome
if Session("u_nome")<>"" then
   sql = sql & " AND " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("u_nome"))
end if

'filtra per email del contatto
if Session("u_email")<>"" then
	sql = sql & " AND  IDElencoIndirizzi IN (SELECT id_Indirizzario FROM tb_valoriNumeri WHERE " & SQL_FullTextSearch(Session("u_email"), "ValoreNumero") & " AND id_TipoNumero = " & VAL_EMAIL & ")"
end if

'filtra per stato abilitazioni
if Session("u_abilitato")<> "" OR Session("u_accessoscaduto")<> "" then
	if not ( Session("u_abilitato")<> "" AND Session("u_accessoscaduto")<> "") then
		sql = sql + " AND ( "
		if Session("u_abilitato")<> "" then
			sql = sql + " ( ut_id IN (SELECT rel_ut_id FROM rel_utenti_sito) AND ( " & SQL_IfIsNull(conn, "ut_scadenzaAccesso", Sql_Now(conn) & " + 1") & " >= " & SQL_date(conn, Date()) & ") AND " & Sql_IsTrue(conn, "ut_abilitato") & ") OR "
		end if
		if Session("u_accessoscaduto")<>"" then
			sql = sql + " ( (" &  SQL_IfIsNull(conn, "ut_scadenzaAccesso", Sql_Now(conn) & " + 1") & " < " & SQL_Now(conn) & ") OR ( NOT" & Sql_IsTrue(conn, "ut_abilitato") & ") ) OR "
		end if
		sql = left(sql, (len(sql)-3)) + " ) "
	end if
end if


'filtra per profilo
if Session("u_profilo")<>"" then
	sql = sql & " AND ut_id IN (SELECT rpu_utenti_id FROM mrel_profili_utenti WHERE rpu_profilo_id = " & Session("u_profilo") & ")"
end if


if sql<>"" then sql = " WHERE 1=1 " & sql
sql = " SELECT (SELECT COUNT(*) FROM rel_utenti_sito WHERE rel_ut_id = tb_utenti.ut_id) AS N_Abilitazioni, " + _
	  		 " isSocieta, CognomeElencoIndirizzi, NomeElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi, ModoRegistra, " & _
			 " IDElencoIndirizzi, tb_utenti.*" & _
	  " FROM tb_indirizzario INNER JOIN tb_utenti " &_
	  " ON tb_indirizzario.IDElencoIndirizzi=tb_utenti.ut_NextCom_ID " &_
	  sql & " ORDER BY ModoRegistra"

Session("SQL_UTENTI") = sql

CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)
%>
<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="27%" valign="top">
<!-- BLOCCO DI RICERCA -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
					<form action="Utenti.asp" method="post" id="ricerca" name="ricerca">
					<tr>
						<td>
							<table cellspacing="1" cellpadding="0" class="tabella_madre">
								<caption>Opzioni di ricerca</caption>
								<tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("u_nome") %>>Cognome e/o nome</th></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_nome" value="<%= session("u_nome")%>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("u_login") %>>Login</td></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_login" value="<%= session("u_login")%>" style="width:100%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("u_abilitato;u_accessoscaduto") %>>STATO ABILITAZIONE</td></tr>
								<tr>
									<td class="content_b">
										<input type="checkbox" class="checkbox" name="search_abilitato" value="1" <%= chk(Session("u_abilitato")<>"") %>>
										abilitato
									</td>
								</tr>							
								<tr>
									<td class="content warning">
										<input type="checkbox" class="checkbox" name="search_accessoscaduto" value="1" <%= chk(Session("u_accessoscaduto")<>"") %>>
										accesso scaduto / non abilitato
									</td>
								</tr>
								<% if (cBoolean(Session("CONDIVISIONE_INTERNA"), false) OR cBoolean(Session("CONDIVISIONE_PUBBLICA"), false)) then %>
									<% 
									sql = "SELECT * FROM mtb_profili ORDER BY pro_nome_it"
									if GetValueList(conn, NULL, sql) <>"" then %>
										<tr><th colspan="2" <%= Search_Bg("u_profilo") %>>PROFILO</th></tr>
										<tr>
											<td class="content" colspan="2">
												<% CALL dropDown(conn, sql, "pro_id", "pro_nome_it", "search_profilo", session("u_profilo"), false, "style=""width: 100%;""", Session("LINGUA")) %>
											</td>
										</tr>
									<% end if %>
								<% end if %>
								<tr><th <%= Search_Bg("u_email") %>>Email</td></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_email" value="<%= session("u_email")%>" style="width:100%;">
									</td>
								</tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><td style="font-size:4px;">&nbsp;</td></tr>
						<tr>
							<td>
								<table cellspacing="1" cellpadding="0" class="tabella_madre">
									<caption class="border">Strumenti</caption>
									<tr>
										<td class="content_right">
											<a style="width:100%; text-align:center; line-height:12px;" class="button"
												title="Apre la palette di export dei dati" 
												onclick="OpenAutoPositionedScrollWindow('UtentiExport.asp', 'export', 240, 142, true);" href="javascript:void(0);">
												EXPORT DATI
											</a>
										</td>
									</tr>
									<tr>
										<td class="content_center">
											<% 
											sql = session("SQL_UTENTI")
											sql = "SELECT * " & right(sql, len(sql) + 1 - instr(1, sql, "FROM tb_indirizzario INNER JOIN tb_utenti", vbTextCompare))
											%>
											<% CALL ExportContattiInRubrica(sql, "IDElencoIndirizzi", "", "") %>
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
					<caption>Elenco utenti area riservata</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> utenti  in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo
							headerCss = ""
							if not IsNull(rs("ut_ScadenzaAccesso")) then
								if rs("ut_ScadenzaAccesso") < Date() OR not rs("ut_abilitato") then
									headerCss = " warning"
								end if
							elseif not rs("ut_abilitato") then
								headerCss = " warning"
							end if
							if headerCss = "" and cIntero(rs("N_Abilitazioni"))=0 then
								headerCss = " alert"
							end if%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header <%= headerCss %>" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<a class="button" href="UtentiMod.asp?ID=<%= rs("IDElencoIndirizzi") %>">
																MODIFICA
															</a>
															&nbsp;
															<% dim altre_aree_riservate
															   altre_aree_riservate = 0
															   if Session("ID_AREA_RISERVATA") <> "" then
																	sql = " SELECT COUNT(*) FROM rel_utenti_sito WHERE rel_ut_id = " & rs("ut_id") & _
																		  " AND rel_sito_id NOT IN (" & Session("ID_AREA_RISERVATA") & ")"
																	altre_aree_riservate = cIntero(GetValueList(conn, NULL, sql))		 
															   end if

															   if altre_aree_riservate = 0 then %>
																<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('UTENTI','<%= rs("ut_id") %>');" >
																	CANCELLA
																</a>
															<% else %>
																<a class="button_disabled" href="javascript:void(0);" title="Utente non cancellabile perchè collegato ad altre aree riservate.">
																	CANCELLA
																</a>
															<% end if %>
														</td>
													</tr>
												</table>
												<%= ContactName(rs) %>
											</td>
										</tr>
										<% if rs("isSocieta") then 
											if rs("CognomeElencoIndirizzi") & rs("NomeElencoIndirizzi")<>"" then  %>
												<tr>
													<td class="label">contatto:</td>
													<td class="content" colspan="3"><%= rs("CognomeElencoIndirizzi") %>&nbsp;<%= rs("NomeElencoIndirizzi") %></td>
												</tr>
											<% end if
										else 
											if rs("NomeOrganizzazioneElencoIndirizzi")<>"" then  %>
												<tr>
													<td class="label">ente:</td>
													<td class="content" colspan="3"><%= rs("NomeOrganizzazioneElencoIndirizzi") %></td>
												</tr>
											<% end if
										end if %>
										<tr>
											<td class="label">stato accesso:</td>
											<% if rs("ut_abilitato") then %>
												<td class="content_b">abilitato</td>
											<% else %>
												<td class="content">non abilitato</td>
											<% end if %>
											<td class="label_no_width" style="width:20%;">scadenza accesso:</td>
											<td class="content" style="width:14%;"><%= DateIta(rs("ut_scadenzaAccesso")) %></td>
										</tr>
										<tr>
											<td class="label" style="width:22%;">login:</td>
											<td class="content"><%= rs("ut_login") %></td>
											<td class="label_no_width" style="text-align:right;">
												<%
												sql = "SELECT COUNT(*) FROM log_utenti WHERE log_ut_id = " & rs("ut_id") 
												if cIntero(GetValueList(conn, NULL, sql)) > 0 then %>
													<a class="button_L2" href="UtentiAccessi.asp?ID=<%= rs("ut_id") %>" title="Accessi dell'utente">
														ACCESSI
													</a>
												<% else %>
													<a class="button_disabled button_L2" style="padding-bottom:0px;" href="javascript:void(0);" title="Nessun accesso effettuato dall'utente.">
														ACCESSI
													</a>
												<% end if %>
											</td>
											<td class="label_no_width">
												<% CALL WriteCampoCerca("Documenti.asp", "utente_id", rs("ut_id"), "DOCUMENTI", "button_L2") %>
											</td>
										</tr>
										<% sql = "SELECT TOP 1 ValoreNumero FROM tb_ValoriNumeri " &_
												  " WHERE id_TipoNumero=6 AND "& SQL_IsTrue(conn, "email_default") &" AND id_Indirizzario=" & rs("IDElencoIndirizzi") 
										email = GetValueList(conn, rsa, sql)%>
										<tr>
											<td class="label" style="width:22%;">email:</td>
											<td class="content" colspan="3"><a href="mailto:<%= email %>"><%= email %></a></td>
										</tr>
										<tr>
											<td class="label">ultimo accesso:</td>
											<td class="content" colspan="3">
												<% sql = "SELECT TOP 1 log_data FROM log_utenti WHERE log_ut_id=" & rs("ut_id") & " ORDER BY log_data DESC " 
												rsa.open sql, conn, adOpenStatic, adLockOptimistic, adcmdText
												if rsa.eof then%>
													accesso non ancora effettuato
												<% else %>
													<%= DateTimeIta(rsa("log_data")) %>
												<% end if
												rsa.close %>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<% rs.moveNext
						wend%>
						<tr>
							<td class="footer" style="border-top:0px; text-align:left;">
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
set rsa = nothing
set conn = nothing
%>