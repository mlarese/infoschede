<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
dim Pager
set Pager = new PageNavigator

'--------------------------------------------------------
sezione_testata = "Selezione del cliente" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
dim conn, sql, rs, rsr
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("cli_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("cli_")
	end if
end if

'filtra per nome
if Session("cli_nome")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("cli_nome"))
end if

'filtra per login
if Session("cli_login")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("cli_login"), "ut_login")
end if

sql = "SELECT * FROM gv_rivenditori "& _
	  "WHERE (1=1) "& sql & _
	  " ORDER BY ModoRegistra"
CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)
%>
<script language="JavaScript" type="text/javascript">
	function Selezione(ObjId, ObjNome){
		opener.form1.<%= request.querystring("field_id") %>.value = ObjId.value;
		opener.form1.<%= request.querystring("field_nome") %>.value = ObjNome.value;
		window.close();
	}
</script>
<div id="content_ridotto">
<form action="" method="post" id="ricerca" name="ricerca">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption>
		<table border="0" cellspacing="0" cellpadding="1" align="right">
			<tr>
				<td style="font-size: 1px; padding-right:1px;" nowrap>
					<input type="submit" name="cerca" value="CERCA" class="button">
					&nbsp;
					<input type="submit" name="tutti" value="VEDI TUTTI" class="button">
				</td>
			</tr>
		</table>
		Opzioni di ricerca
	</caption>
	<tr>
		<th>NOME CONTATTO</th>
		<th>LOGIN CONTATTO</th>
	</tr>
	<tr>
		<td class="content">
			<input type="text" name="search_nome" value="<%= TextEncode(session("cli_nome")) %>" style="width:100%;">
		</td>
		<td class="content">
			<input type="text" name="search_login" value="<%= TextEncode(session("cli_login")) %>" style="width:100%;">
		</td>
	</tr>
</table>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption class="border">Elenco clienti</caption>
	<tr>
		<td>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
				<tr>
					<td class="label_no_width" colspan="3">
						<% if rs.eof then %>
							Nessuna cliente trovato.
						<% else %>
							Trovati n&ordm; <%= Pager.recordcount %> clienti in n&ordm; <%= Pager.PageCount %> pagine
						<% end if %>
					</td>
				</tr>
				<% if not rs.eof then %>
					<tr>
						<th class="L2">SEL.</th>
						<th class="L2">CONTATTO</th>
					</tr>
					<%rs.AbsolutePage = Pager.PageNo
					while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
						<tr>
							<td width="4%" class="content_center">
								<input type="hidden" name="NAME_<%= rs("riv_id") %>" value="<%= ContactFullName(rs) %>">
								<input type="radio" name="seleziona" class="checkbox" value="<%= rs("riv_id") %>" <%= Chk(CInteger(request.querystring("selected")) = rs("riv_id")) %>
									   title="Click per selezionare il cliente"	
									   onclick="Selezione(this, ricerca.NAME_<%= rs("riv_id") %>)">
							</td>
							<td class="content">
								<a href="javascript:void(0);" title="apri scheda del cliente" <%= ACTIVE_STATUS %>
									onclick="OpenAutoPositionedScrollWindow('ClientiGestione.asp?ID=<%= rs("IDElencoIndirizzi") %>', 'cliente', 760, 400, true);">
									<%= ContactFullName(rs) %>
								</a>
							</td>
						</tr>
						<% rs.MoveNext
					wend%>
					<tr>
						<td colspan="3" class="footer">
							<table width="100%" cellpadding="0" cellspacing="0">
								<tr>
									<td><% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%></td>
									<td align="right">
										<a class="button" href="javascript:window.close();" title="chiudi la finestra" <%= ACTIVE_STATUS %>>
											CHIUDI</a>
									</td>
								</tr>
							</table>
							
						</td>
					</tr>
				<% end if %>
			</table>
		</td>
	</tr>
</table>
</form>
</div>
</body>
</html>
<% 
rs.close
conn.close
set rs = nothing
set rsr = nothing
set conn = nothing
%>