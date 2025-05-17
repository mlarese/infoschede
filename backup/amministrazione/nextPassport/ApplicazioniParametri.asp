<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	

dim i, conn, rs, rsr, sql, lock
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("SQL_APPLICAZIONI"), "id_sito", "ApplicazioniParametri.asp")
end if


dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione applicazioni - parametri di funzionamento"
dicitura.puls_new = "INDIETRO;DATI APPLICAZIONE;ACCESSI;TABELLE DATI"
dicitura.link_new = "Applicazioni.asp;ApplicazioniMod.asp?ID=" & request("ID") & ";ApplicazioniAccessi.asp?ID=" & request("ID") & ";ApplicazioniTabelle.asp?ID=" & request("ID")
dicitura.scrivi_con_sottosez() 


if Request.ServerVariables("REQUEST_METHOD")="POST" then
	sql = "SELECT * FROM tb_siti_parametri WHERE par_id=" & cInteger(request("MOD_ID"))
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	if request("MOD_ID") = "" then
		'inserimento
		rs.AddNew
	end if
	rs("par_sito_id") = request("ID")
	rs("par_key") = request("tft_par_key")
	rs("par_value") = request("tft_par_value")
	rs.Update
	rs.close
	
	response.redirect "ApplicazioniParametri.asp?ID=" & request("ID")
end if

sql = "SELECT * FROM tb_siti WHERE id_sito=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica parametri dell'applicazione "<%= rs("sito_nome") %>"</td>
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
		<tr><th colspan="6">PARAMETRI DI GESTIONE / FUNZIONAMENTO</th></tr>
		<% rs.close
		sql = "SELECT * FROM tb_siti_parametri WHERE par_sito_id=" & cIntero(request("ID")) & " ORDER BY par_key"
		rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText%>
			<tr>
				<td class="label" style="width:30%">
					<% if rs.eof then %>
						Nessun parametro per questa applicazione
					<% else %>
						Trovati n&ordm; <%= rs.recordcount %> record
					<% end if %>
				</td>
				<td colspan="5" class="content_right" style="padding-right:0px;">
					<% if request("NEW")="" and request("MOD_ID")="" then %>
						<a class="button_L2" href="?ID=<%= request("ID") %>&NEW=1">
							NUOVO PARAMETRO
						</a>
					<% else %>
						&nbsp;
					<% end if %>
				</td>
			</tr>
			<% if not rs.eof OR request("NEW")<>"" then %>
				<tr>
					<th class="L2">key</th>
					<th class="L2">value</th>
					<th class="l2_center" width="15%" colspan="3">operazioni</th>
				</tr>
			<% end if 
			if not rs.eof then
				while not rs.eof 
					if cInteger(request("MOD_ID"))=rs("par_id") then %>
						<input type="hidden" name="MOD_ID" value="<%= request("MOD_ID") %>">
						<tr>
							<td class="content"><input type="text" class="text" name="tft_par_key" value="<%= rs("par_key") %>" maxlength="50" size="30"></td>
							<td class="content"><input type="text" class="text" name="tft_par_value" value="<%= rs("par_value") %>" maxlength="250" size="70"></td>
							<td class="content_center"><input type="submit" class="button" name="SALVA" value="SALVA"></td>
					<td class="content_center"><a class="button" href="?ID=<%= request("ID") %>" style="padding-top:1px;">ANNULLA</a></td>
						</tr>
					<% else %>
						<tr>
							<td class="content"><%= rs("par_key") %></td>
							<td class="content"><%= rs("par_value") %></td>
							<td class="content_center">
								<a class="button_L2" href="ApplicazioniParametri.asp?ID=<%= request("ID") %>&MOD_ID=<%= rs("par_id") %>">
									MODIFICA
								</a>
							</td>
							<td class="content_center">
								<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('APPLICAZIONI_PARAMETRI','<%= rs("par_id") %>');">
									CANCELLA
								</a>
							</td>
						</tr>
					<%end if
					rs.MoveNext
				 wend
			end if
			if request("NEW")<>"" and request("MOD_ID")="" then%>
				<tr>
					<td class="content"><input type="text" class="text" name="tft_par_key" value="<%= request("tft_par_key") %>" maxlength="50" size="30"></td>
					<td class="content"><input type="text" class="text" name="tft_par_value" value="<%= request("tft_par_value") %>" maxlength="250" size="70"></td>
					<td class="content_center"><input type="submit" class="button" name="SALVA" value="SALVA"></td>
					<td class="content_center"><a class="button" href="?ID=<%= request("ID") %>" style="padding-top:1px;">ANNULLA</a></td>
				</tr>
			<% end if %>
	</table>
	</form>
	&nbsp;
</div>
</body>
</html>
<% conn.close
set conn = nothing%>