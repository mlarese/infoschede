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
dicitura.sezione = "Gestione classi di sconto per quantit&agrave; - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "ScontiQ.asp"
dicitura.scrivi_con_sottosez() 
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova classe di sconto</caption>
		<tr><th colspan="2">DATI DELLA CLASSE</th></tr>
		<tr>
			<td class="label">nome:</td>
			<td class="content">
				<input type="text" class="text" name="tft_scc_nome" value="<%= request("tft_scc_nome") %>" maxlength="255" size="75">
				(*)
			</td>
		</tr>
		<tr><th colspan="2">DEFINIZIONE DEGLI INTERVALLI E RELATIVI SCONTI</th></tr>
		<tr><td colspan="2" class="note">La definizione degli intervalli ed i relaviti sconti sar&agrave; disponibile dopo aver salvato.</td></tr>
		<tr><th colspan="2">NOTE</th></tr>
		<tr>
			<td class="content" colspan="2">
				<textarea style="width:100%;" rows="5" name="tft_scc_note"><%= request("tft_scc_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="2">
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