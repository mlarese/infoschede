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
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione marchi / produttori - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Marchi.asp"
CALL dicitura.InitializeIndex(Index, "gtb_marche", request("ID"))
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_MARCHI_SQL"), "mar_id", "MarchiMod.asp")
end if

sql = " SELECT * FROM gtb_marche LEFT JOIN " & _
      "			gv_rivenditori ON gtb_marche.mar_anagrafica_id = gv_rivenditori.riv_id WHERE mar_id="& cIntero(request("ID"))
set rs = conn.Execute(sql)
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati del marchio</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= cIntero(request("ID")) %>&goto=PREVIOUS" title="marchio precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= cIntero(request("ID")) %>&goto=NEXT" title="marchio successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="2">DATI DELLA MARCA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 	if i = 0 then %>
			<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
		<% 	end if %>
			<td class="content">
				<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
				<input type="text" class="text" name="tft_mar_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("mar_nome_"& Application("LINGUE")(i)) %>" maxlength="50" size="75">
				<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
			</td>
		</tr>
		<%next %>
		<tr>
			<td class="label">codice aziendale:</td>
			<td class="content">
				<input type="text" class="text" name="tft_mar_codice" value="<%= rs("mar_codice") %>" maxlength="20" size="15">
			</td>
		</tr>
		<tr>
			<td class="label">sito internet:</td>
			<td class="content">
				<input type="text" class="text" name="tft_mar_sito" value="<%= rs("mar_sito") %>" maxlength="255" size="75">
			</td>
		</tr>
		<tr>
			<td class="label">logo:</td>
			<td class="content">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_mar_logo", rs("mar_logo"), " width:311px; ", FALSE) %>
			</td>
		</tr>
		<tr>
			<td class="label">immagine:</td>
			<td class="content">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_mar_img", rs("mar_img"), " width:311px; ", FALSE) %>
			</td>
		</tr>
		<tr>
			<td class="label">costruttore:</td>
			<td class="content" colspan="3">
				<table cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td width="54%">
						<input type="hidden" name="tfn_mar_anagrafica_id" value="<%= rs("mar_anagrafica_id") %>">
						<input READONLY type="text" name="cliente" style="padding-left:3px; width:100%" value="<%= ContactFullName(rs) %>" 
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
							<td><textarea style="width:100%;" rows="5" name="tft_mar_descr_<%= Application("LINGUE")(i) %>"><%= rs("mar_descr_" & Application("LINGUE")(i)) %></textarea></td>
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

<%
set rs = nothing
conn.Close
set conn = nothing
%>