<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ApplicazioniTabelleSalva.asp")
end if

dim i, conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")

%>

<%'--------------------------------------------------------
sezione_testata = "inserimento nuova tabella" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_tab_sito_id" value="<%= request("SITO_ID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<caption>Inserimento nuova tabella</caption>
			<tr><th colspan="3">DATI DELLA TABELLA</th></tr>
			<tr>
				<td class="label_no_width" colspan="2">titolo</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_titolo" value="<%= request("tft_tab_titolo") %>" maxlength="255" style="width:50%;">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">tabella</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_name" value="<%= request("tft_tab_name") %>" maxlength="255" style="width:50%;">
					(*)
				</td>
			</tr>
			
			<tr>
				<td class="label_no_width" colspan="2">colore dei contentuti</td>
				<td class="content">
					<% CALL WriteColorPicker_Input("form1", "tft_tab_colore", request("tft_tab_colore"), "", true, false, "") %>
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">sorgente tabella</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_from_sql" value="<%= request("tft_tab_from_sql") %>" maxlength="255" style="width:95%;">
					(*)
				</td>
			</tr>
			
			<tr><th colspan="3">OPZIONI TABELLA</th></tr>
			<tr>
				<td class="label_no_width" colspan="2">ricercabile</td>
				<td class="content">
						<input type="radio" class="checkbox" value="1" name="tfn_tab_ricercabile" <%= chk(cIntero(request("tfn_tab_ricercabile"))=1 OR request("tfn_tab_ricercabile") = "")%>>
						si
						<input type="radio" class="checkbox" value="0" name="tfn_tab_ricercabile" <%= chk(cIntero(request("tfn_tab_ricercabile"))=0 AND request("tfn_tab_ricercabile") <> "")%>>
						no
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">sitemap</td>
				<td class="content">
						<input type="radio" class="checkbox" value="1" name="tfn_tab_per_sitemap" <%= chk(cIntero(request("tfn_tab_per_sitemap"))=1 OR request("tfn_tab_per_sitemap") = "")%>>
						si
						<input type="radio" class="checkbox" value="0" name="tfn_tab_per_sitemap" <%= chk(cIntero(request("tfn_tab_per_sitemap"))=0 AND request("tfn_tab_per_sitemap") <> "")%>>
						no
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">priorit&agrave;</td>
				<td class="content">
						<input type="text" class="text" name="tfn_tab_priorita_base" value="<%= request("tfn_tab_priorita_base") %>" maxlength="5" size="4">
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">pagina di default</td>
				<td class="content">
					<% CALL DropDownPages(conn, "form1", "455", 0, "tfn_tab_pagina_default_id", request("tfn_tab_pagina_default_id"), false, false) %>
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">indicizza in base alla visibilit&agrave;</td>
				<td class="content">
					<input type="radio" class="checkbox" value="1" name="tfn_tab_indicizza_per_visibilita" <%= chk(cIntero(request("tfn_tab_indicizza_per_visibilita"))=1 OR request("tfn_tab_indicizza_per_visibilita") = "")%>>
					si
					<input type="radio" class="checkbox" value="0" name="tfn_tab_indicizza_per_visibilita" <%= chk(cIntero(request("tfn_tab_indicizza_per_visibilita"))=0 AND request("tfn_tab_indicizza_per_visibilita") <> "")%>>
					no
				</td>
			</tr>
				
			<tr><th colspan="3">CAMPI DI LETTURA DEI DATI PER L'INDICE DEI CONTENUTI</th></tr>
			<tr>
				<td class="label_no_width" colspan="2">chiave primaria</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_field_chiave" value="<%= request("tft_tab_field_chiave") %>" maxlength="255" style="width:95%;">
					(*)
				</td>
			</tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label_no_width" colspan="2" rowspan="<%= ubound(Application("LINGUE"))+1 %>">codice</td>
				<% end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_tab_field_codice_<%= Application("LINGUE")(i) %>" value="<%= request("tft_tab_field_codice_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
				</td>
			</tr>
			<%next %>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label_no_width" colspan="2" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo</td>
				<% end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_tab_field_titolo_<%= Application("LINGUE")(i) %>" value="<%= request("tft_tab_field_titolo_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
			<%next %>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label_no_width" colspan="2" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo alternativo</td>
				<% end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0" style="vertical-align: top;">
					<input type="text" class="text" name="tft_tab_field_titolo_alt_<%= Application("LINGUE")(i) %>" value="<%= request("tft_tab_field_titolo_alt_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
				</td>
			</tr>
			<%next %>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label_no_width" style="width:11%;" rowspan="<%= ubound(Application("LINGUE"))+2 %>">gestione link</td>
					<td class="label_no_width" style="width:9%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">url</td>
				<% end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_tab_field_url_<%= Application("LINGUE")(i) %>" value="<%= request("tft_tab_field_url_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
				</td>
			</tr>
			<%next %>
			<tr>
				<td class="label_no_width">parametro</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_parametro" value="<%= request("tft_tab_parametro") %>" maxlength="255" style="width:95%;">
				</td>
			</tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label_no_width" colspan="2" rowspan="<%= ubound(Application("LINGUE"))+1 %>">descrizione</td>
				<% end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_tab_field_descrizione_<%= Application("LINGUE")(i) %>" value="<%= request("tft_tab_field_descrizione_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
				</td>
			</tr>
			<%next %>
		</table>	
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<tr>
				<td class="label_no_width" style="width:11%;" rowspan="4">foto</td>
				<td class="label_no_width" style="width:9%;" rowspan="2">thumbnail</td>
				<td class="content" colspan="2">
					<input type="text" class="text" name="tft_tab_field_foto_thumb" value="<%= request("tft_tab_field_foto_thumb") %>" maxlength="255" style="width:100%;">
				</td>
			</tr>
			<tr>
				<td class="label_no_width">immagine di default</td>
				<td class="content">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_tab_default_foto_thumb", request("tft_tab_default_foto_thumb") , "width:320px;", false) %>
				</td>
			</tr>
			<tr>
				<td class="label_no_width" rowspan="2">zoom</td>
				<td class="content" colspan="2">
					<input type="text" class="text" name="tft_tab_field_foto_zoom" value="<%= request("tft_tab_field_foto_zoom") %>" maxlength="255" style="width:100%;">
				</td>
			</tr>
			<tr>
				<td class="label_no_width" style="width:15%;">immagine di default</td>
				<td class="content">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_tab_default_foto_zoom", request("tft_tab_default_foto_zoom") , "width:320px;", false) %>
				</td>
			</tr>
		</table>
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<tr>
				<td class="label_no_width" colspan="2">visibilit&agrave;</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_field_visibile" value="<%= request("tft_tab_field_visibile") %>" maxlength="255" style="width:100%;">
				</td>
			</tr>	
			<tr>
				<td class="label_no_width" colspan="2">ordine</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_field_ordine" value="<%= request("tft_tab_field_ordine") %>" maxlength="255" style="width:100%;">
				</td>
			</tr>
			<tr>
				<td class="label_no_width" style="width:11%;" rowspan="2">validit&agrave;</td>
				<td class="label_no_width" style="width:9%;">pubblicazione</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_field_data_pubblicazione" value="<%= request("tft_tab_field_data_pubblicazione") %>" maxlength="255" style="width:100%;">
				</td>
			</tr>
			<tr>
				<td class="label_no_width">scadenza</td>
				<td class="label_no_width">
					<input type="text" class="text" name="tft_tab_field_data_scadenza" value="<%= request("tft_tab_field_data_scadenza") %>" maxlength="255" style="width:100%;">
				</td>
			</tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label_no_width" rowspan="<%= (ubound(Application("LINGUE"))+1)*2 %>">meta tag</td>
					<td class="label_no_width" rowspan="<%= ubound(Application("LINGUE"))+1 %>">keywords</td>
				<% end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0" style="vertical-align: top;">
					<textarea name="tft_tab_field_meta_keywords_<%= Application("LINGUE")(i) %>" style="width:94.5%;" rows="4"><%= request("tft_tab_field_meta_keywords_"& Application("LINGUE")(i)) %></textarea>
				</td>
			</tr>
			<%next %>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label_no_width" rowspan="<%= ubound(Application("LINGUE"))+1 %>">description</td>
				<% end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0" style="vertical-align: top;">
					<textarea name="tft_tab_field_meta_description_<%= Application("LINGUE")(i) %>" style="width:94.5%;" rows="4"><%= request("tft_tab_field_meta_description_"& Application("LINGUE")(i)) %></textarea>
				</td>
			</tr>
			<%next %>
			
			<tr><th colspan="3">CAMPI DI RITORNO DATI DALL'INDICE DEI CONTENUTI</th></tr>
			<tr>
				<td class="label_no_width" colspan="2">nome tabella campi URL</td>
				<td class="content" colspan="1">
					<input type="text" class="text" name="tft_tab_return_url_name" value="<%= request("tft_tab_return_url_name") %>" maxlength="255" style="width:95%;">
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2" rowspan="<%= ubound(Application("LINGUE"))+1 %>">URL del contenuto</td>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<td class="content" colspan="1">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0" style="vertical-align: top;">
					<input type="text" class="text" name="tft_tab_field_return_url_<%= Application("LINGUE")(i) %>" value="<%= request("tft_tab_field_return_url_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
				</td>
			</tr>
			<%next %>
			<tr>
				<td class="label_no_width" colspan="2">campo foto thumb</td>
				<td class="content" colspan="1">
					<input type="text" class="text" name="tft_tab_field_return_foto_thumb" value="<%= request("tft_tab_field_return_foto_thumb") %>" maxlength="255" style="width:95%;">
				</td>
			</tr>
			
			<% if TagAbilitati(conn) then %>
				<tr><th colspan="3">GESTIONE TAG</th></tr>
					<tr>
						<td class="label_no_width" colspan="2">tagging abilitato</td>
						<td class="content">
							<input type="radio" class="checkbox" value="1" name="tfn_tab_tags_abilitati" <%= chk(cIntero(request("tfn_tab_tags_abilitati"))=1) %>>
							si
							<input type="radio" class="checkbox" value="0" name="tfn_tab_tags_abilitati" <%= chk(cIntero(request("tfn_tab_tags_abilitati"))=0) %>>
							no
						</td>
					</tr>			
				<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					<tr>
						<% if i=0 then %>
							<td class="label_no_width" colspan="2" rowspan="<%= ubound(Application("LINGUE"))+1 %>">campi utilizzati come tag</td>
						<% end if %>
						<td class="content" colspan="1">
							<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0" style="vertical-align: top;">
							<input type="text" class="text" name="tft_tab_tags_fields_csv_<%= Application("LINGUE")(i) %>" value="<%= request("tft_tab_tags_fields_csv_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
						</td>
					</tr>
				<%next %>
				<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					<tr>
						<% if i=0 then %>
							<td class="label_no_width" colspan="2" rowspan="<%= ubound(Application("LINGUE"))+1 %>">campi utilizzati come tag</td>
						<% end if %>
						<td class="content" colspan="1">
							<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0" style="vertical-align: top;">
							<input type="text" class="text" name="tft_tab_tags_fields_ssv_<%= Application("LINGUE")(i) %>" value="<%= request("tft_tab_tags_fields_ssv_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
						</td>
					</tr>
				<%next %>
			<% end if %>
			
			<tr>
				<td class="footer" colspan="3">
					(*) Campi obbligatori.
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
					<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
				</td>
			</tr>
		</table>
	</form>
</div>
</body>
</html>
<%
set rs = nothing
%>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>