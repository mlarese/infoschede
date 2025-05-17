<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->

<%'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Titolo_sezione, Sezione, HREF, Action
Titolo_sezione = "Tipologie Newsletter - elenco"
HREF = "Comunicazioni.asp;NewsletterTipNew.asp"
Action = "INDIETRO;NUOVO TIPO DI NEWSLETTER"

%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************


dim conn, rs, sql

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT * FROM tb_newsletters ORDER BY nl_nome_it"
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco tipologie di newsletter - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<%%>
		<tr>
			<th class="center" style="width:5%;">ID</th>
			<th>NOME</th>
			<th class="center" style="width:10%;">dinamica</th>
			<th class="center" style="width:5%;">lingua</th>
			<th class="center" style="width:20%;" colspan="2">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("nl_id") %></td>
				<td class="content"><%= rs("nl_nome_IT") %></td>
				<td class="content_center">
					<input type="checkbox" disabled class="checkbox" <%= chk(rs("nl_gestione_dinamica_contenuti")) %>>
				</td>
				<td class="content_center" style="vertical-align:bottom;">
					<% if Trim(cString(rs("nl_lingua")))<>"" then %>
						<img src="../grafica/flag_mini_<%= rs("nl_lingua") %>.jpg">
					<% else %>
						&nbsp;
					<% end if %>
				</td>
				<td class="content_right" style="font-size:1px;" nowrap>
					<a class="button" href="NewsletterTipMod.asp?ID=<%= rs("nl_id") %>">
						MODIFICA
					</a>&nbsp;
					<% sql = "SELECT COUNT(*) FROM tb_newsletters_contents WHERE nlc_tipo_id = " & rs("nl_id") %>
					<% if cIntero(GetValueList(conn, NULL, sql)) > 0 then %>
						<a class="button_disabled" title="Impossibile eliminare il record: sono presenti delle newsletter di questa tipologia." href="javascript:void(0);">
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('NEWSLETTER_TIP','<%= rs("nl_id") %>');" >
							CANCELLA
						</a>
					<% end if %>
				</td>
			</tr>
			<%rs.movenext
		wend%>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set conn = nothing%>