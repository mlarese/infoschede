<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("MagazziniCarichiSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione magazzini carico merce - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "MagazziniCarichi.asp?IDMAG=" & IIF(request.form("tfn_car_magazzino_id")<>"",request.form("tfn_car_magazzino_id"),request("IDMAG"))
dicitura.scrivi_con_sottosez()

dim conn
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo carico merce</caption>
		<tr><th colspan="4">DATI DEL CARICO</th></tr>
		<tr>
			<td class="label">codice fornitore:</td>
			<td class="content">
				<input type="text" class="text" name="tft_car_fornitore_cod" value="<%= request("tft_car_fornitore_cod") %>" maxlength="50" size="10"> (*)
			</td>
			<td class="label">fornitore:</td>
			<td class="content">
				<input type="text" class="text" name="tft_car_fornitore" value="<%= request("tft_car_fornitore") %>" maxlength="50" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">stato:</td>
			<td class="content" colspan="3">
				<input type="checkbox" class="checkbox" name="chk_car_movimentato" <%= chk(request("chk_movimentato")<>"") %>>Movimentato
			</td>
		</tr>
		
		<tr>
			<td class="label">data di carico:</td>
			<td class="content">
				<table cellpadding="0" cellspacing="0">
				<tr>
					<td><% CALL WriteDataPicker_Input("form1", "tfd_car_data", Date(), "", "/", true, true, LINGUA_ITALIANO) %></td>
					<td>&nbsp;(*)</td>
				</tr>
				</table>
			</td>
			<td class="label">magazzino:</td>
			<td class="content">
				<% CALL dropDown(conn, "SELECT * FROM gtb_magazzini ORDER BY mag_nome", "mag_id", "mag_nome", "tfn_car_magazzino_id", IIF(request.form("tfn_car_magazzino_id")<>"",request.form("tfn_car_magazzino_id"),request("IDMAG")), true, "", LINGUA_ITALIANO) %> (*)
			</td>
		</tr>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="3" name="tft_car_note"><%= request("tft_car_note") %></textarea>
			</td>
		</tr>
		<tr><th colspan="4">ARTICOLI ARRIVATI</th></tr>
		<tr><td class="note" colspan="4">L'immissione degli articoli da caricare sar&agrave; disponibile dopo aver salvato.</td></tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA &gt;&gt;">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>