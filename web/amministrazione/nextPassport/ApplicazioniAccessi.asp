<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	

dim i, conn, rs, rss, sql, lock, Pager
set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rss = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	Pager.Reset
	CALL GotoRecord(conn, rs, Session("SQL_APPLICAZIONI"), "id_sito", "ApplicazioniAccessi.asp")
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione applicazioni - log accessi"
if Session("PASS_ADMIN") = "" then
	dicitura.puls_new = "INDIETRO"
	dicitura.link_new = "Applicazioni.asp"
else
	dicitura.puls_new = "INDIETRO;DATI APPLICAZIONE;PARAMETRI;TABELLE DATI"
	dicitura.link_new = "Applicazioni.asp;ApplicazioniMod.asp?ID=" & request("ID") & ";ApplicazioniParamsModifica.asp?ID=" & request("ID") & ";ApplicazioniTabelle.asp?ID=" & request("ID")
end if
dicitura.scrivi_con_sottosez() 

sql = "SELECT * FROM tb_siti WHERE id_sito=" & cIntero(request("ID"))
rss.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

if rss("sito_amministrazione") then
    sql = "SELECT * FROM (log_admin INNER JOIN tb_admin ON log_admin.log_admin_id = tb_admin.id_admin) " & _
          " INNER JOIN tb_siti ON log_admin.log_sito_id = tb_siti.id_sito " & _
          " WHERE tb_siti.id_sito=" & cIntero(request("ID")) & _
          " ORDER BY log_data DESC "
else
    sql = "SELECT *, '' AS log_http_raw FROM ((tb_utenti INNER JOIN tb_Indirizzario ON tb_utenti.ut_NextCom_ID=tb_Indirizzario.IDElencoIndirizzi) " & _
	      " INNER JOIN log_utenti ON tb_utenti.ut_id=log_utenti.log_ut_id) " &_
	      " INNER JOIN tb_siti ON log_utenti.log_sito_id=tb_siti.id_sito " & _
          " WHERE tb_siti.id_sito=" & cIntero(request("ID")) & _
          " ORDER BY log_data DESC "
end if
CALL Pager.OpenSmartRecordset(conn, rs, sql, 25)
%>

<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">
                        Applicazione "<%= rss("sito_nome") %>": 
                        <span style="white-space:nowrap">Trovati n&ordm; <%= Pager.recordcount %> accessi in n&ordm; <%= Pager.PageCount %> pagine</span>
                    </td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="applicazione precedente">
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="applicazione successiva">
							SUCCESSIVA &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
        <% if not rs.eof then%>
			<tr>
				<th class="center" width="20%">DATA</th>
				<th>UTENTE</th>
				<th class="center" width="40%">APPLICAZIONE</th>
			</tr>
			<%rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo %>
				<tr>
					<td class="content_center"><%= DateTimeIta(rs("log_data")) %></td>
					<td class="content">
                        <% if rss("sito_amministrazione") then %>
                            <%= rs("admin_cognome") %>&nbsp;<%= rs("admin_nome") %>
                        <% else %>
                            <%= ContactFullName(rs) %>
                        <% end if %>
                    </td>
					<td class="content"><%= rs("sito_nome") %></td>
					<!--
					<%= rs("log_http_raw")%>
					-->
				</tr>
				<%rs.movenext
			wend%>
			
			<tr>
				<td colspan="3" class="footer" style="text-align:left;">
					<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
				</td>
			</tr>
		<% else %>
			<tr><td class="noRecords">Nessun accesso effettuato</th></tr>
		<% end if %>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% 
rs.close
rss.close
conn.close 
set rs = nothing
set rss = nothing
set conn = nothing
%>
