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
dicitura.sezione = "Gestione magazzini - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Magazzini.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_MAG_SQL"), "mag_id", "MagazziniMod.asp")
end if

sql = "SELECT * FROM gtb_magazzini WHERE mag_id="& cIntero(request("ID"))
set rs = conn.Execute(sql)
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati del magazzino</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="magazzino precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="magazzino successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="3">DATI DEL MAGAZZINO</th></tr>
		<tr>
			<td class="label">Codice magazzino:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_mag_codice" value="<%= rs("mag_codice") %>" maxlength="50" size="40">
			</td>
		</tr>
		<tr>
			<td class="label" style="width:22%;">Nome:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_mag_nome" value="<%= rs("mag_nome") %>" maxlength="255" size="75"> (*)
			</td>
		</tr>
		<tr>
			<td class="label">magazzino ordini al pubblico</td>
			<td class="content" width="10%">
				<input type="checkbox" class="noBorder" name="chk_mag_vendita_pubblico" value="0" <%= chk(rs("mag_vendita_pubblico")) %>>
			</td>
			<td class="note">
				Magazzino utilizzato per la registrazione e lo scarico degli ordini esterni in ingresso.<br>
				ATTENZIONE: Solo un magazzino pu&ograve; essere impostato come "magazzino ordini al pubblico".<br>
				<% sql = "SELECT COUNT(*) FROM gtb_magazzini WHERE mag_vendita_pubblico=1 AND mag_id<>" & cIntero(request("ID"))
				if cInteger(GetValueList(Conn, NULL, sql))>0 then %>
					Impostando questo magazzino come "magazzino ordini al pubblico" verr&agrave; tolta l'opzione al magazzino attualmente assegnato.
				<% end if %>
			</td>
		</tr>
		<tr>
			<td class="label">disponibilit&agrave; in fase di vendita</td>
			<td class="content" width="10%">
				<input type="checkbox" class="noBorder" name="chk_mag_disponibilita" value="0" <%= chk(rs("mag_disponibilita")) %>>
			</td>
			<td class="note">
				Considera il magazzino nel calcolo della quantit&agrave; disponibile in fase di vendita.
			</td>
		</tr>
		<tr><th colspan="3">NOTE</th></tr>
		<tr>
			<td class="content" colspan="3">
				<textarea style="width:100%;" rows="3" name="tft_mag_note"><%= rs("mag_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="3">
				(*) Campi obbligatori.
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
conn.Close
set conn = nothing
%>
