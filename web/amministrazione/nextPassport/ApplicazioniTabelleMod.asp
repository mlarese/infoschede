<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ApplicazioniTabelleSalva.asp")
end if
%>
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->

<%'--------------------------------------------------------
sezione_testata = "modifica dati tabella" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<% 
dim conn, rs, rsr, rsp, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsr = Server.CreateObject("ADODB.Recordset")
set rsp = Server.CreateObject("ADODB.Recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("SQL_APPLICAZIONE_TABELLE"), "tab_id", "ApplicazioniTabelleMod.asp")
end if

sql = "SELECT * FROM tb_siti_tabelle WHERE tab_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_tab_sito_id" value="<%= rs("tab_sito_id") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<caption>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td class="caption">Modifica della tabella</td>
						<td align="right" style="font-size: 1px;">
							<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="tabella precedente">
								&lt;&lt; PRECEDENTE
							</a>
							&nbsp;
							<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="tabella successiva">
								SUCCESSIVA &gt;&gt;
							</a>
						</td>
					</tr>
				</table>
			</caption>
			<tr><th colspan="3">DATI DELLA TABELLA</th></tr>
			<tr>
				<td class="label_no_width" colspan="2">titolo</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_titolo" value="<%= rs("tab_titolo") %>" maxlength="255" style="width:50%;">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">nome tabella</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_name" value="<%= rs("tab_name") %>" maxlength="255" style="width:50%;">
					(*)
				</td>
			</tr>
			
			<tr>
				<td class="label_no_width" colspan="2">colore dei contentuti</td>
				<td class="content">
					<% CALL WriteColorPicker_Input("form1", "tft_tab_colore", rs("tab_colore"), "", true, false, "") %>
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">sorgente tabella</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_from_sql" value="<%= rs("tab_from_sql") %>" maxlength="255" style="width:95%;">
					(*)
				</td>
			</tr>
			
			<tr><th colspan="3">OPZIONI TABELLA</th></tr>
			<tr>
				<td class="label_no_width" colspan="2">ricercabile</td>
				<td class="content">
						<input type="radio" class="checkbox" value="1" name="tfn_tab_ricercabile" <%= chk(rs("tab_ricercabile"))%>>
						si
						<input type="radio" class="checkbox" value="0" name="tfn_tab_ricercabile" <%= chk(not rs("tab_ricercabile"))%>>
						no
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">sitemap</td>
				<td class="content">
						<input type="radio" class="checkbox" value="1" name="tfn_tab_per_sitemap" <%= chk(rs("tab_per_sitemap"))%>>
						si
						<input type="radio" class="checkbox" value="0" name="tfn_tab_per_sitemap" <%= chk(not rs("tab_per_sitemap"))%>>
						no
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">priorit&agrave;</td>
				<td class="content">
						<input type="text" class="text" name="tfn_tab_priorita_base" value="<%= rs("tab_priorita_base") %>" maxlength="5" size="4">
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">pagina di default</td>
				<td class="content">
					<% CALL DropDownPages(conn, "form1", "455", 0, "tfn_tab_pagina_default_id", rs("tab_pagina_default_id"), false, false) %>
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">indicizza in base alla visibilit&agrave;</td>
				<td class="content">
					<input type="radio" class="checkbox" value="1" name="tfn_tab_indicizza_per_visibilita" <%= chk(rs("tab_indicizza_per_visibilita"))%>>
					si
					<input type="radio" class="checkbox" value="0" name="tfn_tab_indicizza_per_visibilita" <%= chk(not rs("tab_indicizza_per_visibilita"))%>>
					no
				</td>
			</tr>
			
			<tr><th colspan="3">CAMPI DI LETTURA DEI DATI PER L'INDICE DEI CONTENUTI</th></tr>
			<tr>
				<td class="label_no_width" colspan="2">chiave primaria</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_field_chiave" value="<%= rs("tab_field_chiave") %>" maxlength="255" style="width:95%;">
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
					<input type="text" class="text" name="tft_tab_field_codice_<%= Application("LINGUE")(i) %>" value="<%= rs("tab_field_codice_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
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
					<input type="text" class="text" name="tft_tab_field_titolo_<%= Application("LINGUE")(i) %>" value="<%= rs("tab_field_titolo_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
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
					<input type="text" class="text" name="tft_tab_field_titolo_alt_<%= Application("LINGUE")(i) %>" value="<%= rs("tab_field_titolo_alt_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
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
					<input type="text" class="text" name="tft_tab_field_url_<%= Application("LINGUE")(i) %>" value="<%= rs("tab_field_url_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
				</td>
			</tr>
			<%next %>
			<tr>
				<td class="label_no_width">parametro</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_parametro" value="<%= rs("tab_parametro") %>" maxlength="255" style="width:95%;">
				</td>
			</tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label_no_width" colspan="2" rowspan="<%= ubound(Application("LINGUE"))+1 %>">descrizione</td>
				<% end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_tab_field_descrizione_<%= Application("LINGUE")(i) %>" value="<%= rs("tab_field_descrizione_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
				</td>
			</tr>
			<%next %>
		</table>	
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<tr>
				<td class="label_no_width" style="width:11%;" rowspan="4">foto</td>
				<td class="label_no_width" style="width:9%;" rowspan="2">thumbnail</td>
				<td class="content" colspan="2">
					<input type="text" class="text" name="tft_tab_field_foto_thumb" value="<%= rs("tab_field_foto_thumb") %>" maxlength="255" style="width:100%;">
				</td>
			</tr>
			<tr>
				<td class="label_no_width" style="width:15%;">immagine di default</td>
				<td class="content">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_tab_default_foto_thumb", rs("tab_default_foto_thumb") , "width:320px;", false) %>
				</td>
			</tr>
			<tr>
				<td class="label_no_width" rowspan="2">zoom</td>
				<td class="content" colspan="2">
					<input type="text" class="text" name="tft_tab_field_foto_zoom" value="<%= rs("tab_field_foto_zoom") %>" maxlength="255" style="width:100%;">
				</td>
			</tr>
			<tr>
				<td class="label_no_width">immagine di default</td>
				<td class="content">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_tab_default_foto_zoom", rs("tab_default_foto_zoom") , "width:320px;", false) %>
				</td>
			</tr>
		</table>
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<tr>
				<td class="label_no_width" colspan="2">visibilit&agrave;</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_field_visibile" value="<%= rs("tab_field_visibile") %>" maxlength="255" style="width:100%;">
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">ordine</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_field_ordine" value="<%= rs("tab_field_ordine") %>" maxlength="255" style="width:100%;">
				</td>
			</tr>
			<tr>
				<td class="label_no_width" style="width:11%;" rowspan="2">validit&agrave;</td>
				<td class="label_no_width" style="width:9%;">pubblicazione</td>
				<td class="content">
					<input type="text" class="text" name="tft_tab_field_data_pubblicazione" value="<%= rs("tab_field_data_pubblicazione") %>" maxlength="255" style="width:100%;">
				</td>
			</tr>
			<tr>
				<td class="label_no_width">scadenza</td>
				<td class="label_no_width">
					<input type="text" class="text" name="tft_tab_field_data_scadenza" value="<%= rs("tab_field_data_scadenza") %>" maxlength="255" style="width:100%;">
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
					<textarea type="text" class="text" name="tft_tab_field_meta_keywords_<%= Application("LINGUE")(i) %>" style="width:94.5%;" rows="4"><%= rs("tab_field_meta_keywords_"& Application("LINGUE")(i)) %></textarea>
				</td>
			</tr>
			<%next
			for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label_no_width" rowspan="<%= ubound(Application("LINGUE"))+1 %>">description</td>
				<% end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0" style="vertical-align: top;">
					<textarea name="tft_tab_field_meta_description_<%= Application("LINGUE")(i) %>" style="width:94.5%;" rows="4"><%= rs("tab_field_meta_description_"& Application("LINGUE")(i)) %></textarea>
				</td>
			</tr>
			<%next %>
			
			<tr><th colspan="3">CAMPI DI RITORNO DATI DALL'INDICE DEI CONTENUTI</th></tr>
			<tr>
				<td class="label_no_width" colspan="2">nome tabella campi URL</td>
				<td class="content" colspan="1">
					<input type="text" class="text" name="tft_tab_return_url_name" value="<%= rs("tab_return_url_name") %>" maxlength="255" style="width:95%;">
				</td>
			</tr>
			<tr>
				<td class="label_no_width"  colspan="2" rowspan="<%= ubound(Application("LINGUE"))+1 %>">URL del contenuto</td>
			<% for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE")) %>
				<td class="content" colspan="1">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0" style="vertical-align: top;">
					<input type="text" class="text" name="tft_tab_field_return_url_<%= Application("LINGUE")(i) %>" value="<%= rs("tab_field_return_url_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
				</td>
			</tr>
			<% next %>
			<tr>
				<td class="label_no_width" colspan="2">campo foto thumb</td>
				<td class="content" colspan="1">
					<input type="text" class="text" name="tft_tab_field_return_foto_thumb" value="<%= rs("tab_field_return_foto_thumb") %>" maxlength="255" style="width:95%;">
				</td>
			</tr>
			
			<% if TagAbilitati(conn) then %>
				<tr><th colspan="3">GESTIONE TAG</th></tr>			
					<tr>
						<td class="label_no_width" colspan="2">tagging abilitato</td>
						<td class="content">
							<input type="radio" class="checkbox" value="1" name="tfn_tab_tags_abilitati" <%= chk(rs("tab_tags_abilitati")) %>>
							si
							<input type="radio" class="checkbox" value="0" name="tfn_tab_tags_abilitati" <%= chk(not rs("tab_tags_abilitati")) %>>
							no
						</td>
					</tr>
					
				<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					<tr>
						<% if i=0 then %>
							<td class="label_no_width"  colspan="2" rowspan="<%= ubound(Application("LINGUE"))+1 %>">campi utilizzati come tag - separati da virgola</td>
						<% end if %>
						<td class="content" colspan="1">
							<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0" style="vertical-align: top;">
							<input type="text" class="text" name="tft_tab_tags_fields_csv_<%= Application("LINGUE")(i) %>" value="<%= rs("tab_tags_fields_csv_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
						</td>
					</tr>
				<%next %>
				<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					<tr>
						<% if i=0 then %>
							<td class="label_no_width"  colspan="2" rowspan="<%= ubound(Application("LINGUE"))+1 %>">campi utilizzati come tag - separati da spazio</td>
						<% end if %>
						<td class="content" colspan="1">
							<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0" style="vertical-align: top;">
							<input type="text" class="text" name="tft_tab_tags_fields_ssv_<%= Application("LINGUE")(i) %>" value="<%= rs("tab_tags_fields_ssv_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:89.8%;">
						</td>
					</tr>
				<%next %>
				<tr><th colspan="3" class="l2">QUERY DI GESTIONE TAG</th></tr>			
				<tr>
					<td colspan="4">
						<% sql = " SELECT * FROM tb_siti_tabelle_tag_query WHERE tq_tab_id=" & cIntero(request("ID")) & _
								 " ORDER BY tq_nome "
						rsp.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch %>
						<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
							<tr>
								<td class="label_no_width" width="30%">
									<% if rsp.eof then %>
										Nessuna query inserita.
									<% else %>
										Trovati n&ordm; <%= rsp.recordcount %> record
									<% end if %>
								</td>
								<td colspan="4" class="content_right" style="padding-right:0px;">
									<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra l'inserimento di un metatag aggiuntivo" <%= ACTIVE_STATUS %>
									   onclick="OpenAutoPositionedWindow('ApplicazioniTabelleQueryNew.asp?TAB_ID=<%= request("ID") %>', 'Query_nuova', 640, 430)">
										NUOVA QUERY
									</a>
								</td>
							</tr>
							<% if not rsp.eof then %>
								<tr>
									<th class="L3">NOME</th>
									<th class="L3">SEPARATORE</th>
									<th class="L3">QUERY</th>
									<th class="L3_center" width="16%" colspan="2">OPERAZIONI</th>
								</tr>
								<%while not rsp.eof %>
									<tr>
										<td class="content" width="15%"><%= rsp("tq_nome") %></td>
										<td class="content" width="15%"><%= rsp("tq_separatore") %></td>
										<td class="content" width="50%"><%= server.HtmlEncode(sintesi(rsp("tq_query"),100,"...")) %></td>
										<td class="content_center">
											<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la modifica dei dati della query." <%= ACTIVE_STATUS %>
											   onclick="OpenAutoPositionedScrollWindow('ApplicazioniTabelleQueryMod.asp?ID=<%= rsp("tq_id") %>&TAB_ID=<%= request("ID") %>', 'Url_<%= rsp("tq_id") %>', 640, 430, true)">
												MODIFICA
											</a>
										</td>
										<td class="content_center">
											<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione della query"
											   onclick="OpenDeleteWindow('QUERY','<%= rsp("tq_id") %>');">
												CANCELLA
											</a>
										</td>
									</tr>
									<%rsp.MoveNext
								wend
							end if%>
						</table>
					<% rsp.close %>
					</td>
				</tr>
			<% end if %>
			
			<% 	sql = " SELECT MIN(tab_id) FROM tb_siti_tabelle WHERE tab_name LIKE '"& rs("tab_name") &"'"
				sql = " SELECT MIN(tab_id) FROM tb_siti_tabelle WHERE tab_id = " & request("ID")
				if CIntero(GetValueList(conn, rsr, sql)) = rs("tab_id") then %>
			<tr><th colspan="4">FORMATI IMMAGINI</th></tr>
			<tr>
				<td colspan="4">
				<% 	sql = " SELECT imf_id, imf_nome, rif_tab_id "& _
		  				  " FROM tb_immaginiFormati"& _
						  " LEFT JOIN rel_immaginiFormati ON (tb_immaginiFormati.imf_id = rel_immaginiFormati.rif_imf_id "& _
		  				  " AND rel_immaginiFormati.rif_tab_id = " & cIntero(request("ID")) & ")"& _
		  				  " ORDER BY imf_nome"
					rsr.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch %>
					<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
						<tr>
							<th class="l2">Formato</th>
							<th class="l2">SELEZIONA</th>
							<th class="l2">THUMB</th>
							<th class="l2">ZOOM</th>
						</tr>
				<% 	while not rsr.eof %>
						<tr>
							<td class="label" style="width: auto;"><%= rsr("imf_nome") %></td>
							<td class="content" style="width: 100px;">
								<input type="checkbox" class="noborder" name="immagini" id="immagini_<%= rsr("imf_id") %>" value="<%= rsr("imf_id") %>" <%= Chk(CIntero(rsr("rif_tab_id")) > 0) %> onclick="ImmaginiEnable(<%= rsr("imf_id") %>)">
								seleziona
							</td>
							<td class="content" style="width: 100px;">
								<input type="radio" class="noborder" name="tfn_tab_thumb" id="tfn_tab_thumb_<%= rsr("imf_id") %>" value="<%= rsr("imf_id") %>" <%= Chk(rsr("imf_id") = rs("tab_thumb")) %>>
								thumb
							</td>
							<td class="content" style="width: 100px;">
								<input type="radio" class="noborder" name="tfn_tab_zoom" id="tfn_tab_zoom_<%= rsr("imf_id") %>" value="<%= rsr("imf_id") %>" <%= Chk(rsr("imf_id") = rs("tab_zoom")) %>> 
								zoom
								<script type="text/javascript">
									function ImmaginiEnable(id) {
										EnableIfChecked(document.getElementById('immagini_' + id), document.getElementById('tfn_tab_thumb_' + id))
										EnableIfChecked(document.getElementById('immagini_' + id), document.getElementById('tfn_tab_zoom_' + id))
									}
									ImmaginiEnable(<%= rsr("imf_id") %>)
								</script>
							</td>
						</tr>
				<%		rsr.moveNext
					wend %>
					</table>
				<%	rsr.close %>
				</td>
			</tr>
			<% 	end if %>
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
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>
</body>
</html>
<% rs.close
conn.close
set rs = nothing
set rsr = nothing
set conn = nothing %>