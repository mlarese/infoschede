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

dim i
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo Metodo di spedizione articolo </caption>
		<tr><th colspan="3">DATI DEL METODO DI SPEDIZIONE DELL' ARTICOLO</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 	if i = 0 then %>
			<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome spedizione:</td>
		<% 	end if %>
			<td class="content" colspan="2">
				<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
				<input type="text" class="text" name="tft_spa_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_spa_nome_"& Application("LINGUE")(i)) %>" maxlength="50" size="75">
				<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
			</td>
		</tr>
		<%next %>
		<tr>
			<td class="label">importo spedizione:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_spa_importo_spese" value="<%= request("tfn_spa_importo_spese") %>" maxlength="20" size="10">
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
				<input type="text" class="text" name="tfn_spa_annullamento_qta" value="<%= request("tfn_spa_annullamento_qta") %>" maxlength="20" size="10">
				&nbsp;(*)
			</td>
			<td class="note">
				Quantit&agrave; entro la quale viene azzerato l'importo della spedizione.  
			</td>
		</tr>
		<tr>
			<td class="label">importo:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_spa_annullamento_importo" value="<%= request("tfn_spa_annullamento_importo") %>" maxlength="20" size="10">
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
							<td><textarea style="width:100%;" rows="5" name="tft_spa_condizioni_<%= Application("LINGUE")(i) %>"><%= request("tft_spa_condizioni_" & Application("LINGUE")(i)) %></textarea></td>
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