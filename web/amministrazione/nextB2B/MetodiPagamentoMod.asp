<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("MetodiPagamentoSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione modalità di pagamento / produttori - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "MetodiPagamento.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_MODPAGA_SQL"), "mosp_id", "MetodiPagamentoMod.asp")
end if

sql = "SELECT * FROM gtb_modipagamento WHERE mosp_id="& cIntero(request("ID"))
set rs = conn.Execute(sql)
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre_intermedia">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica modalit&agrave; di pagamento</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="modalit&agrave; precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="modalit&agrave; successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="2">DATI DELLA MODALIT&Agrave;</th></tr>
		<tr>
			<td class="label" style="width:20%;">modalit&agrave; abilitata:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_mosp_se_abilitato" <%= chk(rs("mosp_se_abilitato")) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_mosp_se_abilitato" <%= chk(not rs("mosp_se_abilitato")) %>>
				no
			</td>
		</tr>
		<tr>
			<td class="label" style="width:20%;">modalit&agrave; di default:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_mosp_default" <%= chk(rs("mosp_default")) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_mosp_default" <%= chk(not rs("mosp_default")) %>>
				no
			</td>
		</tr>
		<tr>
			<td class="label">modalit&agrave; personalizzata:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_mosp_personalizzato"<%= chk(rs("mosp_personalizzato")) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_mosp_personalizzato"<%= chk(not rs("mosp_personalizzato")) %>>
				no
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 	if i = 0 then %>
			<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
		<% 	end if %>
			<td class="content">
				<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
				<input type="text" class="text" name="tft_mosp_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("mosp_nome_"& Application("LINGUE")(i)) %>" maxlength="50" size="94">
				<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
			</td>
		</tr>
		<%next %>
		<tr>
			<td class="label">codice:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_mosp_codice" value="<%= rs("mosp_codice") %>" maxlength="200" size="40">
			</td>
		</tr>
		<tr>
			<td class="label">icona / logo del servizio:</td>
			<td class="content">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_mosp_url_logo_servizio_modo", rs("mosp_url_logo_servizio_modo") , "width:424px;", false) %>
			</td>
		</tr>
		
	</table>
	
	<script type="text/javascript" language="JavaScript">
	
		function ApplicaVariazioniPrezzo(tag){
				tag.value = FormatNumber(tag.value, 2);
		}
		
		function SetControlsState(){
			tfn_mosp_se_esterno_true = document.getElementById("tfn_mosp_se_esterno_true");
			EnableIfChecked(tfn_mosp_se_esterno_true, form1.tft_mosp_url_servizio_esterno);
			DisableIfChecked(tfn_mosp_se_esterno_true, form1.tfn_mosp_id_pagina_startup);
			
			tfn_mosp_se_spesespedizione_true = document.getElementById("tfn_mosp_se_spesespedizione_true");
			EnableIfChecked(tfn_mosp_se_spesespedizione_true, form1.tfn_mosp_ammontare_spsp);
		}
	</script>
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre_intermedia">
		<tr><th colspan="5">PARAMETRI</th></tr>
		<tr>
			<td class="label_no_width">stato ordine:</td>
			<td class="content" colspan="2">
					<% sql = "SELECT * FROM gtb_stati_ordine ORDER BY so_stato_ordini, so_ordine"
					CALL dropDown(conn, sql, "so_id", "so_nome_it", "tfn_mosp_stato_ordine_id", rs("mosp_stato_ordine_id"), false, " style=""width=80%""", LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label_no_width" rowspan="4" style="width:18%;">gestito esternamente:</td>
			<td class="content" rowspan="2">
				<input type="radio" class="checkbox" value="1" name="tfn_mosp_se_esterno" id="tfn_mosp_se_esterno_true" <%= chk(rs("mosp_se_esterno")) %> onclick="SetControlsState()">
				si
			</td>
			<td class="label_no_width" colspan="2">indirizzo del servizio (es. paypal):</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_mosp_url_servizio_esterno" value="<%= rs("mosp_url_servizio_esterno") %>" maxlength="200" size="73">
			</td>
		</tr>
		<tr>
			<td class="content" rowspan="2">
				<input type="radio" class="checkbox" value="0" name="tfn_mosp_se_esterno" id="tfn_mosp_se_esterno_false" <%= chk(not rs("mosp_se_esterno")) %> onclick="SetControlsState()">
				no
			</td>
			<td class="label_no_width" colspan="2">pagina di avvio procedura guidata: </td>
		</tr>
		<tr>
			<td class="content">
				<% CALL DropDownPages(NULL, "form1", 400, 0, "tfn_mosp_id_pagina_startup", rs("mosp_id_pagina_startup"), false, false) %>
			</td>
		</tr>
		<tr>
			<td class="label_no_width" rowspan="2">con spese di incasso:</td>
			<td class="content" colspan="2">
				<input type="radio" class="checkbox" value="1" name="tfn_mosp_se_spesespedizione" id="tfn_mosp_se_spesespedizione_true" <%= chk(rs("mosp_se_spesespedizione")) %> onclick="SetControlsState()">
				si
				<input type="radio" class="checkbox" value="0" name="tfn_mosp_se_spesespedizione" id="tfn_mosp_se_spesespedizione_false" <%= chk(not rs("mosp_se_spesespedizione")) %> onclick="SetControlsState()">
				no
			</td>
		</tr>
		<tr>
			<td class="label_no_width" style="width:12%;">importo spese:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_mosp_ammontare_spsp" value="<%= rs("mosp_ammontare_spsp") %>" maxlength="6" size="6" onchange="ApplicaVariazioniPrezzo(this)">
				&nbsp;&euro;
			</td>
		</tr>
		<tr>
			<td class="label_no_width" rowspan="2">pagamento immediato</td>
			<td class="content" colspan="2">
				<input type="radio" class="checkbox" value="1" name="tfn_mosp_pag_immediato" id="tfn_mosp_pag_immediato_true" <%= chk(rs("mosp_pag_immediato")) %> onclick="SetControlsState()">
				si
				<input type="radio" class="checkbox" value="0" name="tfn_mosp_pag_immediato" id="tfn_mosp_pag_immediato_false" <%= chk(not rs("mosp_pag_immediato")) %> onclick="SetControlsState()">
				no
			</td>
		</tr>
	</table>
	
	<script type="text/javascript" language="JavaScript">
		SetControlsState();
	</script>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
	<tr><th colspan="2">ISTRUZIONI</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="2">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="5" name="tft_mosp_istruzioni_<%= Application("LINGUE")(i) %>"><%= rs("mosp_istruzioni_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		<%next %>

		<tr><th colspan="2">DESCRIZIONE PER IL CLIENTE</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="2">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="5" name="tft_mosp_label_spsp_<%= Application("LINGUE")(i) %>"><%= rs("mosp_label_spsp_" & Application("LINGUE")(i)) %></textarea></td>
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