<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" AND request("salva")<>"" then
	Server.Execute("MarchiSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura, conn, sql
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione marchi / produttori - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Marchi.asp"
dicitura.scrivi_con_sottosez() 

dim i
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_mar_generica" value="0">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo marchio</caption>
		<tr><th colspan="2">DATI DEL MARCHIO</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 	if i = 0 then %>
			<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
		<% 	end if %>
			<td class="content">
				<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
				<input type="text" class="text" name="tft_mar_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_mar_nome_"& Application("LINGUE")(i)) %>" maxlength="50" size="75">
				<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
			</td>
		</tr>
		<%next %>
		<tr>
			<td class="label">codice aziendale:</td>
			<td class="content">
				<input type="text" class="text" name="tft_mar_codice" value="<%= request("tft_mar_codice") %>" maxlength="20" size="15">
			</td>
		</tr>
		<tr>
			<td class="label">sito internet:</td>
			<td class="content">
				<input type="text" class="text" name="tft_mar_sito" value="<%= request("tft_mar_sito") %>" maxlength="255" size="75">
			</td>
		</tr>
		<tr>
			<td class="label">logo:</td>
			<td class="content">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_mar_logo", request.form("tft_mar_logo"), " width:311px; ", FALSE) %>
			</td>
		</tr>
		<tr>
			<td class="label">immagine:</td>
			<td class="content">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_mar_img", request.form("tft_mar_img"), " width:311px; ", FALSE) %>
			</td>
		</tr>
		<tr>
			<td class="label">costruttore:</td>
			<td class="content" colspan="3">
				<table cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td  width="54%">
						<input type="hidden" name="tfn_mar_anagrafica_id" value="<%= request.form("tfn_mar_anagrafica_id") %>">
						<input READONLY type="text" name="cliente" style="padding-left:3px; width:100%" value="<%= request.form("cliente") %>" 
							   onclick="OpenAutoPositionedScrollWindow('ClientiSelezione.asp?field_nome=cliente&field_id=tfn_mar_anagrafica_id&selected=' + tfn_mar_anagrafica_id.value + '&filtro_profilo=<%=COSTRUTTORI%>', 'SelezioneCliente', 450, 480, true)" title="Click per aprire la finestra per la selezione del cliente">
					</td>
					<td>
						<a class="button_input" href="javascript:void(0)" onclick="form1.cliente.onclick();" 
							 title="Apre la filnestra per la selezione del cliente" <%= ACTIVE_STATUS %>>
							SELEZIONA CLIENTE
						</a>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr><th colspan="2">DESCRIZIONE</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="2">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="5" name="tft_mar_descr_<%= Application("LINGUE")(i) %>"><%= request("tft_mar_descr_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="footer" colspan="2">
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