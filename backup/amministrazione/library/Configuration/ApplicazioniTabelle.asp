<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../Tools4Color.asp" -->
<!--#INCLUDE FILE="../../nextPassport/ToolsApplicazioni.asp" -->
<% 	

dim i, conn, rs, rsr, sql, value
set conn = Server.CreateObject("ADODB.Connection")
conn.open GetConfigurationConnectionstring()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("SQL_APPLICAZIONI"), "id_sito", "ApplicazioniTabelle.asp")
end if


dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione applicazioni - tabelle dati"
dicitura.puls_new = "INDIETRO;DATI APPLICAZIONE;PARAMETRI"
dicitura.link_new = "Applicazioni.asp;ApplicazioniMod.asp?ID=" & request("ID") & ";ApplicazioniParamsModifica.asp?ID=" & request("ID")
dicitura.scrivi_con_sottosez() 

sql = "SELECT * FROM tb_siti WHERE id_sito=" & cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dell'applicazione "<%= rs("sito_nome") %>"</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="applicazione precedente">
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="applicazione successiva">
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="9">TABELLE DI MEMORIZZAZIONE DATI PUBBLICATI</th></tr>
		<% rs.close
		sql = " SELECT *" + _
			  " FROM tb_siti_tabelle t WHERE tab_sito_id=" & cIntero(request("ID")) & " ORDER BY tab_titolo"
		Session("SQL_APPLICAZIONE_TABELLE") = sql
		rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText %>
		<tr>
			<td class="label" style="width:30%" colspan="2">
				<% if rs.eof then %>
					Nessuna tabella definita per l'applicazione
				<% else %>
					Trovati n&ordm; <%= rs.recordcount %> record
				<% end if %>
			</td>
			<td colspan="7" class="content_right" style="padding-right:0px;">
				<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra l'inserimento di una nuova tabella" <%= ACTIVE_STATUS %>
				   onclick="OpenAutoPositionedScrollWindow('ApplicazioniTabelleNew.asp?SITO_ID=<%= request("ID") %>', 'tabelle', 740, 400, true)">
					NUOVA TABELLA
				</a>
			</td>
		</tr>
		<% if not rs.eof then %>
			<tr>
				<th class="L2">nome</th>
				<th class="l2_center" width="5%">colore</th>
				<th class="L2">nome tabella</th>
				<th class="l2_center" width="24%" colspan="2">operazioni</th>
			</tr>
			<% while not rs.eof %>
				<tr>
					<td class="content"><%= rs("tab_titolo")%></td>
					<td class="content_center"><% WriteColor(rs("tab_colore"))%></td>
					<td class="content"><%= rs("tab_name")%></td>
					<td class="content_center">
						<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la modifica della tabella" <%= ACTIVE_STATUS %>
						   onclick="OpenAutoPositionedScrollWindow('ApplicazioniTabelleMod.asp?ID=<%= rs("tab_id") %>', 'tabella_<%= rs("tab_id") %>', 740, 250, true)">
							MODIFICA
						</a>
					</td>
					<td class="content_center">
						<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione della tabella" <%= ACTIVE_STATUS %>
						   onclick="OpenDeleteWindow('SITI_TABELLE','<%= rs("tab_id") %>');">
							CANCELLA
						</a>
					</td>
				</tr>
				<% rs.movenext
			wend
		end if
		rs.close %>
	</table>
	&nbsp;
</div>
</body>
</html>
<% conn.close
set conn = nothing%>
	