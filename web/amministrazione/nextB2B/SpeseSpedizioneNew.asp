<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SpeseSpedizioneSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione spese di spedizione / modalit&agrave; spedizione ordine"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SpeseSpedizione.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo Metodo di Consegna Ordine</caption>
		<tr><th colspan="3">DATI DEL METODO DI CONSEGNA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">area geografica:</td>
			<% 	end if %>
				<td class="content" colspan="2">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_sp_area_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_sp_area_nome_"& Application("LINGUE")(i)) %>" maxlength="50" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label" style="width:20%;">codice:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_sp_codice" value="<%= request("tft_sp_codice") %>" maxlength="500" size="20">
			</td>
		</tr>
		<script language="JavaScript" type="text/javascript">
			
			function setState(){
				var spese_spedizione = document.getElementById("spese_spedizione_true");
				EnableIfChecked(spese_spedizione, document.getElementById("tfn_sp_importo_euro"));
				DisableIfChecked(spese_spedizione, document.getElementById("tfn_sp_percentuale"));
				DisableIfChecked(spese_spedizione, document.getElementById("minimo_spese"));
				
				if (spese_spedizione.checked){
					document.getElementById('percent').innerHTML='';
					document.getElementById('minim').innerHTML='';
					document.getElementById('costo').innerHTML='(*)';
				}
				else{
					document.getElementById('percent').innerHTML='(*)';
					document.getElementById('minim').innerHTML='(*)';
					document.getElementById('costo').innerHTML='';
				}
			}
				
		</script>
		<tr>
			<td class="label" rowspan="2">spese spedizione:</td>
			<td class="label_no_width" colspan="2">
				<table style="width:100%;">
					<tr>
						<td class="label" style="width:19%;">
							<input type="radio" class="checkbox" value="1" name="spese_spedizione" id="spese_spedizione_true" value="1" <%= chk(request("spese_spedizione")="1" OR request("spese_spedizione")="") %> onClick="setState()">
							costo
						</td>
						<td class="content" style="width:24%;">
							&euro;
							<input type="text" class="text" name="tfn_sp_importo_euro" id="tfn_sp_importo_euro" value="<%= request("tfn_sp_importo_euro") %>" maxlength="20" size="10">
							&nbsp;<span id="costo"></span>
						</td>
						<td class="note">
							Spese di spezione previste per l'area geografica. 
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td class="label_no_width" colspan="2">
				<table style="width:100%;">
					<tr>
						<td class="label" style="width:19%;">
							<input type="radio" class="checkbox" value="0" name="spese_spedizione" value="" <%= chk(request("spese_spedizione")="0") %> onClick="setState()">
							percentuale
						</td>
						<td class="content" style="width:24%;">
							&nbsp;&nbsp;
							<input type="text" class="text" name="tfn_sp_percentuale" id="tfn_sp_percentuale" value="<%= request("tfn_sp_percentuale") %>" maxlength="20" size="10">
							%&nbsp;<span id="percent"></span>
						</td>
						<td class="note">
							Spese spedizione colcolate in percentuale sul totale dell'ordine. 
						</td>
					</tr>
					<tr>
						<td class="label">
							<span style="width:19px;">&nbsp;</span>
							costo minimo
						</td>
						<td class="content">
							&euro;
							<input type="text" class="text" name="tfn_sp_importo_euro" id="minimo_spese" value="<%= request("tfn_sp_importo_euro") %>" maxlength="20" size="10">
							&nbsp;<span id="minim"></span>
						</td>
						<td class="note">
							Importo minimo per le spese di spedizione. 
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td class="label">categoria i.v.a.:</td>
			<td class="content" colspan="5">
				<% sql = "SELECT * FROM gtb_iva ORDER BY iva_ordine"
				CALL dropDown(conn, sql, "iva_id", "iva_nome", "tfn_sp_iva_id", request("tfn_sp_iva_id"), true, "", LINGUA_ITALIANO)%>
			</td>
		</tr>
		<tr><th colspan="3">CONDIZIONI DI APPLICAZIONE</th></tr>
		<tr>
			<td class="label">importo:</td>
			<td class="content">
				&euro;
				<input type="text" class="text" name="tfn_sp_annullamento_importo" value="<%= request("tfn_sp_annullamento_importo") %>" maxlength="20" size="10">
				&nbsp;(*)
			</td>
			<td class="note">
				Prezzo oltre il quale viene azzerato l'importo della spedizione.
			</td>
		</tr>
		<tr><th colspan="3">CONDIZIONI GENERALI</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="3">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="5" name="tft_sp_condizioni_<%= Application("LINGUE")(i) %>"><%= request("tft_sp_condizioni_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		<%next %>
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
<script language="JavaScript" type="text/javascript">
	setState()
</script>
</html>