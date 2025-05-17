<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(2)
dicitura.sottosezioni(1) = "IMPEGNI"
dicitura.links(1) = "Impegni.asp"
dicitura.sottosezioni(2) = "CONFIGURAZIONE"
dicitura.links(2) = "AgendaConfigura.asp"

dicitura.sezione = "Gestione tipologie impegni/appuntamenti - elenco"
dicitura.puls_new = "NUOVA TIPOLOGIA"
dicitura.link_new = "ImpegniTipologieNew.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, Pager

set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT * FROM mtb_tipi_impegni " + _
	  " WHERE (1=1) " + sql + _
	  " ORDER BY tim_nome_it"
Session("SQL_TIPI_IMPEGNI") = sql

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
							<th style="width:5%;">COLORE</th>
							<th>NOME TIPOLOGIA</th>
							<th style="width:25%;">N&deg; DOCUMENTI ASSOCATI</th>
							<th style="width:21%; text-align:center;">OPERAZIONI</th>
						</tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo
							dim impegni_associati
							impegni_associati = CIntero(GetValueList(conn, NULL, "SELECT COUNT(imp_id) FROM mtb_impegni WHERE imp_tipo_id = " & rs("tim_id")))
							%>
							<tr>
								<td class="content_center"><% WriteColor(rs("tim_colore"))%></td>
								<td class="content">
									<%= rs("tim_nome_it") %>
								</td>
								<td class="content"><%= IIF(impegni_associati > 0, impegni_associati, " - ")%>
								</td>
								<td class="content">
									<table border="0" cellspacing="0" cellpadding="0" align="right">
										<tr>
											<td style="font-size: 1px;">
												<a class="button" href="ImpegniTipologieMod.asp?ID=<%= rs("tim_id") %>">
													MODIFICA
												</a>
												&nbsp;
												<% if impegni_associati > 0 then %>
													<a class="button_disabled" href="javascript:void(0);" 
															title="Tipologia non cancellabile. <%=impegni_associati%> documenti/circolari associati.">
														CANCELLA
													</a>
												<% else %>
													<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('IMPEGNI_TIPOLOGIA','<%= rs("tim_id") %>');" >
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
							<td class="footer" colspan="4" style="border-top:0px; text-align:left;">
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