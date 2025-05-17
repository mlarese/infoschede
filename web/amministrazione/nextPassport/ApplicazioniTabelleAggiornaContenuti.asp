<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1000000 %>
<% 
if request("campo_ordinamento")<>"" then
	response.buffer = false
end if
 %>
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp"-->

<%'--------------------------------------------------------
sezione_testata = "aggiorna contenuti tabella" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<% 
dim conn, rs, rss, rss1, rsi, sql, field, rssCount
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsi = Server.CreateObject("ADODB.Recordset")
set rss = Server.CreateObject("ADODB.Recordset")
set rss1 = Server.CreateObject("ADODB.Recordset")

sql = " SELECT *, " + _
	  " (SELECT COUNT(*) FROM tb_siti_tabelle_pubblicazioni WHERE pub_tabella_id=t.tab_id) AS N_PUBBLICAZIONI, " + _
	  " (SELECT COUNT(*) FROM v_indice WHERE co_F_table_id=t.tab_id) AS N_CONTENUTI " + _
	  " FROM tb_siti_tabelle t WHERE tab_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenDynamic, adLockOptimistic


sql = "SELECT COUNT(*) FROM " + rs("tab_from_sql")
rssCount = cIntero(GetValueList(conn, rss, sql))
sql = "SELECT TOP 1 * FROM " + rs("tab_from_sql")
rss1.open sql, conn, adOpenDynamic, adLockOptimistic

%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<caption>Aggiornamento contenuti della tabella</caption>
			<tr><th colspan="3">DATI DELLA TABELLA</th></tr>
			<tr>
				<td class="label" style="width:26%;">nome:</td>
				<td class="content">
					<%= rs("tab_titolo") %>
					<% WriteColor(rs("tab_colore")) %>
				</td>
			</tr>
			<tr>
				<td class="label">tabella:</td>
				<td class="content"><%= rs("tab_name") %></td>
			</tr>
			<tr>
				<td class="label">query sorgente:</td>
				<td class="content"><%= rs("tab_from_sql") %></td>
			</tr>
			<tr><th colspan="3">CONTEGGIO CONTENUTI</th></tr>
			<tr>
				<td class="label">numero pubblicazioni automatiche:</td>
				<td class="content"><%= rs("N_PUBBLICAZIONI") %></td>
			</tr>
			<tr>
				<td class="label">numero attuale contenuti:</td>
				<td class="content"><%= rs("N_CONTENUTI") %></td>
			</tr>
			<tr>
				<td class="label">numero righe sorgenti:</td>
				
				<td class="content"><%= rssCount %></td>
			</tr>
			<% if request.form("esegui")<>"" then %>
				<tr><th colspan="3">AGGIORNAMENTO CONTENUTI</th></tr>
			
				<% 
				dim aggiornabile
				
				sql = "SELECT * FROM " + rs("tab_from_sql")
				if cIntero(request("metodo_aggiornamento"))<>0 then
					'aggiorna solo record non presenti
					sql = sql + " WHERE " & rs("tab_field_chiave") & " NOT IN (SELECT co_f_key_id FROM v_indice_it WHERE co_f_table_id=" & rs("tab_id") & ") "
				end if
				if request("campo_ordinamento")<>"" then
					sql = sql + " ORDER BY " + ParseSQL(request("campo_ordinamento"), adChar)
					if lcase(request("campo_ordinamento")) <> lcase(rs("tab_field_chiave")) then
						sql = sql + ", " + rs("tab_field_chiave")
					end if
				end if
				rss.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
				while not rss.eof 
					
					index.conn.begintrans
				
					if cIntero(request("metodo_aggiornamento"))=0 then
						aggiornabile = true
					else
						sql = "SELECT top 1 idx_id FROM v_indice WHERE tab_id=" & rs("tab_id") & " AND co_f_key_id=" & rss(cString(rs("tab_field_chiave")))
						aggiornabile = cIntero(GetValueList(Index.conn, rsi, sql))<1
					end if
					%>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
						<tr>
							<td class="label_no_width" style="width:26%;">
								<% if aggiornabile then %>
									contenuto in aggiornamento:
								<% else %>
									contenuto saltato perchè già presente:
								<% end if %>
							</td>
							<td class="content"><%= rss(cString(rs("tab_field_chiave"))) %></td>
							<td class="content_right"><%= rss.AbsolutePosition %> di <%= rss.recordcount %></td>
						</tr>
					</table>
					<% if aggiornabile then
						CALL Index_UpdateItem(index.conn, rs("tab_name"), rss(cString(rs("tab_field_chiave"))), false)
					end if
					
					index.conn.committrans
				
					rss.movenext
				wend
				
				%>
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<tr>
						<td class="content ok">
							AGGIORNAMENTO ESEGUITO CORRETTAMENTE
						</td>
					</tr>
					<tr>
						<td class="footer" colspan="3">
							<input type="button" class="button" name="annulla" value="CHIUDI" onclick="window.close();">
						</td>
					</tr>
				</table>
			<% else %>
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<tr><th colspan="3">GESTIONE AGGIORNAMENTO</th></tr>
					<tr>
						<td class="label" rowspan="4">modalità di aggiornamento</td>
						<td class="content_center" rowspan="2" style="width:5%">
							<input type="radio" class="checkbox" name="metodo_aggiornamento" value="0" <%= chk(rssCount < 200)%>>
						</td>
						<td class="content">
							completo
						</td>
					</tr>
					<tr>
						<td class="content note">
							Esegue l'aggiornamento completo di tutti i contenuti.
						</td>
					</tr>
					<tr>
						<td class="content_center" rowspan="2" style="width:5%">
							<input type="radio" class="checkbox" name="metodo_aggiornamento" value="1" <%= chk(rssCount >= 200)%>>
						</td>
						<td class="content">
							solo mancanti
						</td>
					</tr>
					<tr>
						<td class="content note">
							Aggiorna solo i contenuti mancanti dall'indice.
						</td>
					</tr>
					<tr>
						<td class="label" style="width:26%;">ordinamento:</td>
						<td class="content" colspan="2">
							<select name="campo_ordinamento">
								<% for each field in rss1.fields %>
									<option value="<%= field.name %>" <%= IIF(lcase(field.name) = lcase(rs("tab_field_chiave")), "selected", "") %>><%= field.name %></option>
								<% next %>
							</select><br>
							<span class="note">
							Specifica il campo secondo il quale viene ordinato il set di righe di origine per l'esecuzione dell'aggiornamento.
							</span>
						</td>
					</tr>
					<tr>
						<td class="footer" colspan="3">
							<input type="submit" class="button" name="esegui" value="ESEGUI AGGIORNAMENTO">
							<input type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
						</td>
					</tr>
				</table>
			<% end if %>
	</form>
</div>
</body>
</html>
<% 
rs.close
rss1.close
conn.close
set rs = nothing
set rsi = nothing
set rss = nothing
set conn = nothing %>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>