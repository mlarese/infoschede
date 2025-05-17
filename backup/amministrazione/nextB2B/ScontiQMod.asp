<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ScontiQSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione classi di sconto per quantit&agrave; - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "ScontiQ.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, rsd, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.Recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_SCONTIQ_SQL"), "scc_ID", "ScontiQMod.asp")
end if

sql = "SELECT * FROM gtb_scontiQ_classi WHERE scc_id="& cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica della classe di sconto</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="classe precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="classe successiva" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="2">DATI DELLA CLASSE</th></tr>
		<tr>
			<td class="label">nome:</td>
			<td class="content">
				<input type="text" class="text" name="tft_scc_nome" value="<%= rs("scc_nome") %>" maxlength="255" size="75">
				(*)
			</td>
		</tr>
		<tr><th colspan="4">DEFINIZIONE INTERVALLI DI SCONTO</th></tr>
		<tr>
			<td colspan="4">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<%sql = "SELECT * FROM gtb_scontiQ WHERE sco_classe_id=" & rs("scc_id") & " ORDER BY sco_qta_da" 
					rsd.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText%>
					<tr>
						<td class="label" colspan="4" style="width:80%;">
							<% if rsd.eof then %>
								nessun intervallo di sconto definito per la classe.
							<% else %>
								n&ordm; <%= rsd.recordcount %> intervallo di sconto definito per la classe.
							<% end if %>
						</td>
						<td colspan="2" class="content_right" style="padding-right:0px;">
							<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ScontiQ_dettagliNew.asp?EXTID=<%= request("ID") %>', 'ScontiQ', 510, 200, true)"
							   title="apre la finestra per l'inserimento di nuovo intervallo e relativo sconto" <%= ACTIVE_STATUS %>>
								NUOVO INTERVALLO
							</a>
						</td>
					</tr>
					<% if not rsd.eof then %>
						<tr>
							<th class="L2">A PARTIRE DA (n&ordm; unit&agrave;)</th>
							<th class="L2">FINO A (n&ordm; unit&agrave;)</th>
							<th class="L2">VARIAZIONE %</th>
							<th class="L2">PREZZO UNITARIO</th>
							<th class="l2_center" width="16%" colspan="2">OPERAZIONI</th>
						</tr>
						<% if rsd("sco_qta_da")>1 then%>
							<tr>
								<td class="content">1</td>
								<td class="content"><%= rsd("sco_qta_da") - 1 %></td>
								<td class="content">0 %</td>
								<td class="content"><i>(prezzo di listino)</i></td>
								<td class="content_center">&nbsp;</td>
								<td class="content_center">&nbsp;</td>
							</tr>
						<%end if
						while not rsd.eof %>
							<tr>
								<td class="content"><%= rsd("sco_qta_da") %></td>
								<td class="content">
									<% rsd.movenext
									if not rsd.eof then %>
										<%= rsd("sco_qta_da")-1 %>
									<% end if
									rsd.movePrevious %>
								</td>
								<% if cReal(rsd("sco_sconto"))<>0 then %>
									<td class="content"><%= rsd("sco_sconto") %> %</td>
								<% else %>
									<td class="content">-</td>
								<% end if %>
								<% if cReal(rsd("sco_prezzo"))>0 then %>
									<td class="content"><%= FormatPrice(rsd("sco_prezzo"), 2, true) %> &euro;</td>
								<% else %>
									<td class="content">-</td>
								<% end if %>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ScontiQ_dettagliMod.asp?EXTID=<%= request("ID") %>&ID=<%= rsd("sco_id") %>', 'ScontiQ', 510, 200, true)">
										MODIFICA
									</a>
								</td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('SCONTIQ_D','<%= rsd("sco_id") %>');">
										CANCELLA
									</a>
								</td>
							</tr>
							<%rsd.MoveNext
						wend
					end if
					rsd.close%>
				</table>
			</td>
		</tr>
		<tr><th colspan="2">NOTE</th></tr>
		<tr>
			<td class="content" colspan="2">
				<textarea style="width:100%;" rows="5" name="tft_scc_note"><%= rs("scc_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="2">
				(*) Campi obbligatori.
				<input type="submit" style="width:23%;" class="button" name="mod" value="SALVA & TORNA ALL'ELENCO">
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>

<%
set rs = nothing
set rsd = nothing
conn.Close
set conn = nothing
%>