<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="../../nextPassport/ToolsApplicazioni.asp" -->
<%
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(3)
dicitura.sottosezioni(1) = "APPLICAZIONI"
dicitura.links(1) = "Applicazioni.asp"
dicitura.sottosezioni(2) = "PARAMETRI"
dicitura.links(2) = "ApplicazioniParams.asp"
dicitura.sottosezioni(3) = "GRUPPI DI PARAMETRI"
dicitura.links(3) = "ApplicazioniParamsGruppi.asp"
dicitura.puls_new = "NUOVO PARAMETRO"
dicitura.link_new = "ApplicazioniParamsNew.asp"
dicitura.sezione = "Gestione parametri - elenco"
dicitura.scrivi_con_sottosez()

dim conn, rs, rsa, sql, Pager, applicativi
set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("sid_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("sid_")
	end if
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open GetConfigurationConnectionstring()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")

'aggiunge filtri di ricerca
sql = ""

'filtra per nome
if Session("sid_nome")<>"" then
	sql = sql & IIF(sql = "", " WHERE ", " AND ") & SQL_FullTextSearch(Session("sid_nome"), FieldLanguageList("sid_nome_"))
end if

'filtra per tipo
if CIntero(Session("sid_tipo")) > 0 then
	sql = sql & IIF(sql = "", " WHERE ", " AND ") & " sid_tipo = "& Session("sid_tipo")
end if

'filtra per raggruppamento
if CIntero(Session("sid_raggruppamento")) > 0 then
	sql = sql & IIF(sql = "", " WHERE ", " AND ") & " sid_raggruppamento_id = "& Session("sid_raggruppamento")
end if

'filtra per applicativo
if CIntero(Session("sid_applicativo")) > 0 then
	sql = sql & IIF(sql = "", " WHERE ", " AND ") &" EXISTS (SELECT 1 FROM rel_siti_descrittori WHERE rsd_descrittore_id = sid_id AND rsd_sito_id = "& Session("sid_applicativo") &")"
end if

sql = " SELECT * FROM tb_siti_descrittori d" & _
	  " LEFT JOIN tb_siti_descrittori_raggruppamenti r ON d.sid_raggruppamento_id = r.sdr_id" & _
      sql & _
	  " ORDER BY sid_nome_it"
Session("SQL_DESCRITTORI_APPLICATIVI") = sql
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
										<td class="footer" colspan="2">
											<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" id="tutti_top" value="VEDI TUTTI" style="width: 49%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("sid_nome") %>>TITOLO</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_nome" value="<%= TextEncode(session("sid_nome")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("sid_tipo") %>>TIPO</th></tr>
									<tr>
										<td class="content" colspan="2">
                                            <% CALL DesAdvancedDropTipi("search_tipo", "width:100%;", cIntero(session("sid_tipo")), false) %>
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("sid_raggruppamento") %>>RAGGRUPPAMENTO</th></tr>
									<tr>
										<td class="content" colspan="2">
                                            <%	sql = "SELECT * FROM tb_siti_descrittori_raggruppamenti ORDER BY sdr_titolo_it"
                                            CALL DropDown(conn, sql, "sdr_id", "sdr_titolo_it", "search_raggruppamento", session("sid_raggruppamento"), false, "style=""width:100%;""", LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg("sid_applicativo") %>>APPLICATIVO</th></tr>
									<tr>
										<td class="content" colspan="2">
                                            <%	sql = "SELECT * FROM tb_siti ORDER BY sito_nome"
                                            CALL DropDown(conn, sql, "id_sito", "sito_nome", "search_applicativo", session("sid_applicativo"), false, "style=""width:100%;""", LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr>
										<td class="footer" colspan="2">
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
					<caption>
						Elenco parametri degli applicativi
					</caption>
                    <% if not rs.eof then %>
						<tr><th>Trovate n&ordm; <%= Pager.recordcount %> parametri in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo %>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size: 1px;">
															<a class="button" href="ApplicazioniParamsMod.asp?ID=<%= rs("sid_id") %>">
                                       							MODIFICA
                                    						</a>&nbsp;
                                                            <a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('APPLICAZIONI_PARAMS','<%= rs("sid_id") %>');" >
                                                                CANCELLA
                                                            </a>
														</td>
													</tr>
												</table>
												<%= rs("sid_nome_IT") %>
											</td>
										</tr>
                                        <tr>
											<td class="label">codice:</td>
					                        <td class="content"><%= rs("sid_codice") %></td>
                                            <td class="label">raggruppamento:</td>
					                        <td class="content"><%= rs("sdr_titolo_it") %></td>
                                        </tr>
                                        <tr>
                                            <td class="label">tipo:</td>
                                            <td class="content" style="width: 40%;"><%= DesVisTipo(rs("sid_tipo")) %></td>
                                            <td class="label" style="white-space: nowrap;">visibile agli utenti:</td>
                                            <td class="content"><input type="checkbox" disabled class="Checkbox" <%= chk(NOT rs("sid_admin")) %>></td>
                                        </tr>
										<% 
										sql = " SELECT * FROM tb_siti INNER JOIN rel_siti_descrittori ON tb_siti.id_sito = rel_siti_descrittori.rsd_sito_id " + _
											  	 " WHERE rsd_descrittore_id = " & rs("sid_id")
										rsa.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText 
										if rsa.eof then %>
											<tr>
												<td class="label">applicazioni:</td>
												<td class="content_disabled" colspan="3">
													parametro non utilizzato
												</td>
											</tr>
										<% else 
											while not rsa.eof%>
												<tr>
													<% if rsa.absoluteposition=1 then %>
														<td class="label" rowspan="<%= rsa.recordcount %>">applicazioni:</td>
													<% end if %>
													<td class="content" colspan="3">
														<%= rsa("sito_nome") %>
													</td>
												</tr>
												<% rsa.movenext
											wend
										end if 
										rsa.close%>
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
set conn = nothing%>

	
