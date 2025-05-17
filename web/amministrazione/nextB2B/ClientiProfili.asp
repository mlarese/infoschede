<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(2)
dicitura.sottosezioni(1) = "CLIENTI"
dicitura.links(1) = "Clienti.asp"
dicitura.sezione = "Gestione profili clienti - elenco"
dicitura.puls_new = "NUOVO PROFILO"
dicitura.link_new = "ClientiProfiliNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, rsv, sql, pager, i

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")

set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("pro_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("pro_")
	end if
end if

'filtra per nome
if Session("pro_nome")<>"" then
	sql = sql & " AND (" & SQL_FullTextSearch(Session("pro_nome"), FieldLanguageList("pro_nome_")) & ") "
end if

sql = " SELECT *, (SELECT COUNT(*) FROM gv_rivenditori WHERE riv_profilo_id= gtb_profili.pro_id ) AS N_ANAGRAFICHE " + _
      " FROM gtb_profili WHERE (1=1) " + sql + " ORDER BY pro_nome_it"
session("CLI_PROFILI_SQL") = sql
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
									<tr><th <%= Search_Bg("pro_nome") %>>NOME</th></tr>
									<tr>
										<td class="content">
											<input type="text" name="search_nome" value="<%= TextEncode(session("pro_nome")) %>" style="width:100%;">
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
					<caption>
						Elenco profili
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovati n&ordm; <%= Pager.recordcount %> profili in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
                            <tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0">
										<tr>
											<td class="header" colspan="4">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
                                                    <tr>
                                                        <td style="font-size: 1px;">
                                                            <a class="button" href="ClientiProfiliMod.asp?ID=<%= rs("pro_id") %>">MODIFICA</a>
                                                            &nbsp;
                                                            <% if rs("N_ANAGRAFICHE") > 0 then %>
                        										<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare il marchio: sono presenti anagrafiche che utilizzano questo profilo">
                        											CANCELLA
                        										</a>
                        									<% else %>
                        										<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('CLIENTI_PROFILI','<%= rs("pro_id") %>');" >
                        											CANCELLA
                        										</a>
                        									<% end if %>
														</td>
													</tr>
												</table>
												<%= rs("pro_nome_it") %>
											</td>
										</tr>
                                        <tr>
                                            <td class="label_no_width" style="width:18%;">n&ordm; anagrafiche</td>
                                            <td class="content"><%= cIntero(rs("N_ANAGRAFICHE")) %></td>
										</tr>
									</table>
								</td>
							</tr>
							<% rs.moveNext
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
set rsv = nothing
set conn = nothing%>