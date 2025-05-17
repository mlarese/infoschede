<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("CaratteristicheSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(1)
dicitura.sottosezioni(1) = "GRUPPI"
dicitura.links(1) = "CaratteristicheGruppi.asp"
dicitura.sezione = "Gestione caratteristiche tecniche - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Caratteristiche.asp"
dicitura.scrivi_con_sottosez() 

dim conn, i, rsc, rs, sql, value
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rsc = server.CreateObject("ADODB.recordset")
set rs = server.CreateObject("ADODB.recordset")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova caratteristica tecnica</caption>
		<tr><th colspan="2">DATI DELLA CARATTERISTICA TECNICA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% 	if i = 0 then %>
					<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
				<% 	end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_ct_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_ct_nome_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">tipo di dato:</td>
			<td class="content">
				<% DesDropTipi "tfn_ct_tipo", "", request.Form("tfn_ct_tipo") %>
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% 	if i = 0 then %>
					<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">unit&agrave; di misura:</td>
				<% 	end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_ct_unita_<%= Application("LINGUE")(i) %>" value="<%= request("tft_ct_unita_"& Application("LINGUE")(i)) %>" maxlength="50" size="50">
				</td>
			</tr>
		<%next %>
		<tr><th colspan="2">GESTIONE DELLA CARATTERISTICA TECNICA</th></tr>
		<tr>
			<td class="label" >codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_ct_codice" value="<%= request("tft_ct_codice") %>" maxlength="250" size="26">
			</td>
		</tr>
		<% if cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM gtb_carattech_raggruppamenti"))>0 then %>
			<tr>
				<td class="label">gruppo:</td>
				<td class="content" colspan="3">
					<% sql = "SELECT * FROM gtb_carattech_raggruppamenti ORDER BY ctr_titolo_it"
	                CALL dropDown(conn, sql, "ctr_id", "ctr_titolo_it", "tfn_ct_raggruppamento_id", request.form("tfn_ct_raggruppamento_id"), false, "", LINGUA_ITALIANO) %>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label">immagine:</td>
			<td class="content">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_ct_img", request.form("tft_ct_img"), "width:430px;", FALSE) %>
			</td>
		</tr>
		<tr>
			<td class="label">confrontabile:</td>
			<td class="content"><input type="checkbox" class="checkbox" name="chk_ct_per_confronto" <%= chk(request("chk_ct_per_confronto")<>"") %>></td>
		</tr>
		<tr>
			<td class="label">ricercabile:</td>
			<td class="content"><input type="checkbox" class="checkbox" name="chk_ct_per_ricerca" <%= chk(request("chk_ct_per_ricerca")<>"") %>></td>
		</tr>
		<tr><th colspan="2">CATEGORIE DI PRODOTTI A CUI &Egrave; ASSOCIATA</th></tr>
		<tr>
			<td colspan="2">
				<%sql = ", (NULL) AS ORDINE " + _
						 ", (0) AS N_ARTICOLI " + _
						 " FROM "%>
				<% sql = replace(categorie.QueryElenco(true, ""), " FROM ", sql) 
				rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<% if rsc.eof then %>
						<tr>
							<td class="label_no_width">
								Nessuna categoria di prodotti definita.
							</td>
						</tr>
					<% else %>
						<tr>
							<th class="l2_center" width="6%">associa</th>
							<th class="l2_center" width="7%">ordine</th>
							<th class="L2">categoria</th>
						</tr>
						<% while not rsc.eof %>
							<tr>
								<td class="content_center">
									<% if cInteger(rsc("N_ARTICOLI"))>0 then 
										value = true%>
										<input type="checkbox" checked class="checked" id="categorie_associate_<%= rsc("tip_id") %>" disabled onclick="set_state_<%= rsc("tip_id") %>(this)" title="Sono presenti valori negli articoli di questa categoria.">
										<input type="hidden" name="categorie_associate" value=" <%= rsc("tip_id") %> ">
									<% else 
										value = instr(1, request("categorie_associate"), " " & rsc("tip_id") & " ", vbTextCompare)>0%>
										<input type="checkbox" name="categorie_associate" id="categorie_associate_<%= rsc("tip_id") %>" value=" <%= rsc("tip_id") %> " <%= chk(value) %> class="<%= IIF(value, "checked", "checkbox") %>" onclick="set_state_<%= rsc("tip_id") %>(this)">
									<% end if %>
								</td>
								<td class="content_center"><input <%= disable(not value) %> type="text" class="<%= IIF(not value, "text_disabled", "text") %>" size="2" name="rel_ordine_<%= rsc("tip_id") %>" value="<%= request("rel_ordine_" & rsc("tip_id")) %>"></td>
								<td class="content"><%= rsc("NAME") %></td>
							</tr>
							<script language="JavaScript" type="text/javascript">
								function set_state_<%= rsc("tip_id") %>(chk){
									EnableIfChecked(chk, form1.rel_ordine_<%= rsc("tip_id") %>);
									if (chk.checked){
										form1.rel_ordine_<%= rsc("tip_id") %>.title = "Inserisci l'ordine di visualizzazione nella scheda dell'articolo";
									}
									else{
										form1.rel_ordine_<%= rsc("tip_id") %>.title = "Selezionare il flag di associazione prima di inserire l'ordine di visualizzazione nella scheda articolo";
									}
								}
							</script>
							<% rsc.movenext
						wend %>
					<% end if %>
				</table>
				<% rsc.close %>
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
<%
set rsc = nothing
conn.Close
set conn = nothing
%>