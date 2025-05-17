<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../ClassPageNavigator.asp" -->
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
dicitura.puls_new = "NUOVA APPLICAZIONE"
dicitura.link_new = "ApplicazioniNew.asp"
dicitura.sezione = "Gestione applicazioni - elenco"
dicitura.scrivi_con_sottosez()

dim conn, rs, rsp, sql, sqlAdmin, sqlUtenti, prm, Pager
set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open GetConfigurationConnectionstring()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("app_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("app_")
	end if
end if

sql = ""
'filtra per nome applicazione
if Session("app_nome")<>"" then
	sql = sql & IIF(sql <> "", " AND ", " WHERE ")
	sql = sql & SQL_FullTextSearch(Session("app_nome"), "sito_nome;sito_dir")
end if

sql = "SELECT * FROM tb_siti " + sql + _
	  "ORDER BY sito_nome"
Session("SQL_APPLICAZIONI") = sql
CALL Pager.OpenSmartRecordset(conn, rs, sql, 20)

sql = "SELECT COUNT(*) FROM tb_siti"
prm = CInt(GetValueList(conn, NULL, sql))
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
										<input type="submit" class="button" name="cerca" value="CERCA" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
								<tr><th <%= Search_Bg("app_nome") %>>NOME</td></tr>
								<tr>
									<td class="content">
										<input type="text" name="search_nome" value="<%= session("app_nome")%>" style="width:100%;">
									</td>
								</tr>
								<tr>
									<td class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
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
					<caption>Elenco applicazioni installate</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> applicazioni in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
										<tr>
											<td class="header<%= IIF(rs("sito_amministrazione"), "", " warning") %>" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
													<tr>
														<td style="font-size:1px;">
															<a class="button" href="ApplicazioniParamsModifica.asp?ID=<%= rs("id_sito") %>">
																PARAMETRI
															</a>
															&nbsp;
															<a class="button" href="ApplicazioniMod.asp?ID=<%= rs("id_sito") %>">
																MODIFICA
															</a>
															&nbsp;
															<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('APPLICAZIONI','<%= rs("id_sito") %>');" >
																CANCELLA
															</a>
														</td>
													</tr>
												</table>
												<%= rs("sito_nome") %>
											</td>
										</tr>
										<tr>
											<td class="label">dati dell'applicazione:</td>
											<td class="label_right" style="width:78%;" colspan="3">
												<a class="button_L2" href="ApplicazioniTabelle.asp?ID=<%= rs("id_sito") %>">
													TABELLE DATI
												</a>
											</td>
										</tr>
										<tr>
											<td class="label" style="width:22%;">tipo</td>
											<td class="content" colspan="3">
												<% if rs("sito_amministrazione") then %>
													applicazione area amministrativa
												<% else %>
													applicazione su area riservata pubblica
												<% end if %>
											</td>
										</tr>
										<tr>
											<td class="label">percorso:</td>
											<td class="content"><%= rs("sito_dir") %></td>
											<td class="label">id:</td>
											<td class="content_right" style="width:5%;"><%= rs("id_sito") %></td>
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
	&nbsp;
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set rsp = nothing
set conn = nothing
%>