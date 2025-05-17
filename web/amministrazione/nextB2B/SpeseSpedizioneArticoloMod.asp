<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SpeseSpedizioneArticoloSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione modalit&agrave; di spedizione dell'articolo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SpeseSpedizioneArticolo.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_SPA_SQL"), "spa_id", "SpeseSpedizioneArticoloMod.asp")
end if

sql = "SELECT * FROM gtb_spese_spedizione_articolo WHERE spa_id="& cIntero(request("ID"))
set rs = conn.Execute(sql)
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica condizioni di applicazione</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="area precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="area successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="3">DATI DEL METODO DI SPEDIZIONE DELL'ARTICOLO</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 	if i = 0 then %>
			<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
		<% 	end if %>
			<td class="content" colspan="2">
				<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
				<input type="text" class="text" name="tft_spa_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("spa_nome_"& Application("LINGUE")(i)) %>" maxlength="50" size="75">
				<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
			</td>
		</tr>
		<%next %>
		
		<tr>
			<td class="label">importo spedizione:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_spa_importo_spese" value="<%= FormatPrice(rs("spa_importo_spese"), 2, false)%>" maxlength="20" size="10">
				&euro;&nbsp;(*)
			</td>
			<td class="note">
				Importo previsto per il tipo di spedizione. 
			</td>
		</tr>
		<tr><th colspan="3">CONDIZIONI DI APPLICAZIONE</th></tr>
		<tr>	
			<td class="label">quantit&agrave;:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_spa_annullamento_qta" value="<%= rs("spa_annullamento_qta")%>" maxlength="20" size="10">
				&nbsp;(*)
			</td>
			<td class="note">
				Quantit&agrave; entro la quale viene azzerato l'importo della spedizione. 
			</td>
		</tr>
		<tr>
			<td class="label">importo:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_spa_annullamento_importo" value="<%= FormatPrice(rs("spa_annullamento_importo"), 2, false)%>" maxlength="20" size="10">
				&euro;&nbsp;(*)
			</td>
			<td class="note">
				Prezzo entro il quale viene azzerato l'importo della spedizione. 
			</td>
		</tr>
		<tr><th colspan="3">DESCRIZIONE</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="3">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="5" name="tft_spa_condizioni_<%= Application("LINGUE")(i) %>"><%= rs("spa_condizioni_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		<%next %>
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