<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("MagazziniSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione magazzini - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Magazzini.asp"
dicitura.scrivi_con_sottosez() 


dim conn, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo magazzino</caption>
		<tr><th colspan="3">DATI DEL MAGAZZINO</th></tr>
		<tr>
			<td class="label">Codice magazzino:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_mag_codice" value="<%= request("tft_mag_codice") %>" maxlength="50" size="40">
			</td>
		</tr>
		<tr>
			<td class="label" style="width:22%;">Nome:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_mag_nome" value="<%= request("tft_mag_nome") %>" maxlength="255" size="75">(*)
			</td>
		</tr>
		<tr>
			<td class="label">magazzino ordini al pubblico</td>
			<td class="content" width="10%">
				<input type="checkbox" class="noBorder" name="chk_mag_vendita_pubblico" value="0" <%= chk(request("chk_mag_vendita_pubblico")<>"") %>>
			</td>
			<td class="note">
				Magazzino utilizzato per la registrazione e lo scarico degli ordini esterni in ingresso.<br>
				<% sql = "SELECT COUNT(*) FROM gtb_magazzini WHERE mag_vendita_pubblico=1" 
				if cInteger(GetValueList(Conn, NULL, sql))>0 then %>
					ATTENZIONE: impostando questa opzione verr&agrave; tolta l'impostazione al magazzino attualmente attivo.
				<% end if %>
			</td>
		</tr>
		<tr>
			<td class="label">disponibilit&agrave; in fase di vendita</td>
			<td class="content" width="10%">
				<input type="checkbox" class="noBorder" name="chk_mag_disponibilita" value="0" <%= chk(request("chk_mag_disponibilita")<>"") %>>
			</td>
			<td class="note">
				Considera il magazzino nel calcolo della quantit&agrave; disponibile in fase di vendita.
			</td>
		</tr>
		<tr><th colspan="3">NOTE</th></tr>
		<tr>
			<td class="content" colspan="3">
				<textarea style="width:100%;" rows="3" name="tft_mag_note"><%= request("tft_mag_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="3">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA &gt:&gt;">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% 
conn.close 
set conn = nothing
%>
