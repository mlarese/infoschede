<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione profili - elenco"
dicitura.puls_new = "NUOVO PROFILO"
dicitura.link_new = "ProfiliNew.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, Pager

set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT * FROM mtb_profili " + _
	  " WHERE (1=1) " + sql + _
	  " ORDER BY pro_nome_it"
Session("SQL_PROFILI") = sql

CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)
%>
<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
			<td valign="top">
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						Elenco documenti - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine
					</caption>
					<% if not rs.eof then %>
						<tr>
							<th>NOME PROFILO</th>
							<th style="width:30%;">N&deg; DOCUMENTI ASSOCATI</th>
							<th style="width:21%; text-align:center;">OPERAZIONI</th>
						</tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo
							dim doc_associati
							doc_associati = CIntero(GetValueList(conn, NULL, "SELECT COUNT(rdp_doc_id) FROM mrel_doc_profili WHERE rdp_profilo_id = " & rs("pro_id")))
							%>
							<tr>
								<td class="content">
									<%= rs("pro_nome_it") %>
								</td>
								<td class="content"><%= IIF(doc_associati > 0, doc_associati, " - ")%>
								</td>
								<td class="content">
									<table border="0" cellspacing="0" cellpadding="0" align="right">
										<tr>
											<td style="font-size: 1px;">
												<a class="button" href="ProfiliMod.asp?ID=<%= rs("pro_id") %>">
													MODIFICA
												</a>
												&nbsp;
												<% if doc_associati > 0 then %>
													<a class="button_disabled" href="javascript:void(0);" title="Profilo non cancellabile. <%=doc_associati%> documenti/circolari associati.">
														CANCELLA
													</a>
												<% else %>
													<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('PROFILI','<%= rs("pro_id") %>');" >
														CANCELLA
													</a>
												<% end if %>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<% rs.moveNext
						wend %>
						<tr>
							<td class="footer" colspan="3" style="border-top:0px; text-align:left;">
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
set conn = nothing
%>