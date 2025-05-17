<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#include file="Intestazione.asp"-->
<%

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Duplicazione sito"
dicitura.scrivi_con_sottosez()
%>

<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
	<form action="" method="post" id="form1" name="form1">
		<caption>Duplicazione sito: selezione database</caption>
        <tr><th colspan="2">Database da cui importare</th></tr>
		<% if request("conn_import")="" then %>
			<tr>
				<td class="label" style="width:18%;">connessione da cui importare:</td>
				<td class="content">
					<input type="text" name="conn_import" value="<%=Application("DATA_ConnectionString")%>" class="text" style="width:100%;">
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="3">
					<input style="width:20%;" type="submit" class="button" name="importa" value="AVANTI &gt;&gt;">
				</td>
			</tr>
		<% else 
			dim sql
			dim dconn, sconn, drs, srs, lingua
			
			set drs = Server.CreateObject("ADODB.RecordSet")
			set srs = Server.CreateObject("ADODB.RecordSet")
			set dconn = Server.CreateObject("ADODB.Connection")
			set sconn = Server.CreateObject("ADODB.Connection")
			dconn.open Application("DATA_ConnectionString")
			sconn.open request("conn_import")
			%>
			<tr>
				<td class="label">database da cui importare:</td>
				<td class="content">
					<%= sconn.connectionString %>
				</td>
			</tr>
			<tr><th colspan="2">Siti destinazione e sorgente</th></tr>
			<input type="hidden" name="conn_import" value="<%=request("conn_import")%>">
			<% if cIntero(request("SOURCE_WEB_ID"))= cintero(request("DEST_WEB_ID")) OR _	
				  cIntero(request("SOURCE_WEB_ID"))=0 OR cintero(request("DEST_WEB_ID"))=0 then 
				'richiede sito sorgente e destinazione 
				%>
				<tr>
					<td class="label">sito di origine:</td>
					<td class="content">
						<% sql = "SELECT * FROM tb_webs"
						CALL dropDown(sconn, sql, "id_webs", "nome_webs", "SOURCE_WEB_ID", request("SOURCE_WEB_ID"), true, "", LINGUA_ITALIANO)%>
					</td>
				</tr>
				<tr>
					<td class="label">sito di destinazione:</td>
					<td class="content">
						<% sql = "SELECT * FROM tb_webs"
						CALL dropDown(dconn, sql, "id_webs", "nome_webs", "DEST_WEB_ID", request("DEST_WEB_ID"), true, "", LINGUA_ITALIANO)%>
					</td>
				</tr>
				
				<tr>
					<td class="footer" colspan="3">
						<input style="width:20%;" type="submit" class="button" name="importa" value="AVANTI &gt;&gt;">
					</td>
				</tr>
			<%else 
				'esegue duplicazione
				dim dwebs_id, swebs_id
				dwebs_id = cIntero(request("DEST_WEB_ID"))
				swebs_id = cIntero(request("SOURCE_WEB_ID"))
				%>
				<input type="hidden" name="SOURCE_WEB_ID" value="<%=request("SOURCE_WEB_ID")%>">
				<input type="hidden" name="DEST_WEB_ID" value="<%=request("DEST_WEB_ID")%>">
				
				<% sql = "SELECT *, " + _
						 "(SELECT COUNT(*) FROM tb_paginesito WHERE id_web=tb_webs.id_webs) AS N_PAGINE, " + _
						 "(SELECT COUNT(*) FROM tb_objects WHERE id_webs=tb_webs.id_webs) AS N_OBJECTS, " + _
						 "(SELECT COUNT(*) FROM tb_pages WHERE id_webs=tb_webs.id_webs) AS N_PAGES, " + _
						 "(SELECT COUNT(*) from tb_layers INNER JOIN tb_pages ON tb_layers.id_pag = tb_pages.id_page WHERE id_webs=tb_webs.id_webs) AS N_LAYERS " + _
					     "from tb_webs WHERE id_webs="
				
				srs.open sql & swebs_id, sconn, adOpenStatic, adLockOptimistic, adCmdText
				drs.open sql & dwebs_id, dconn, adOpenStatic, adLockOptimistic, adCmdText
				%>
				<tr>
					<td class="label">sito di origine:</td>
					<td class="content">
						<%=srs("nome_webs")%>
					</td>
				</tr>
				<tr>
					<td class="label">sito di destinazione:</td>
					<td class="content">
						<%=drs("nome_webs")%>
					</td>
				</tr>
		</table>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
				<tr><th colspan="3">Dati da copiare</th></tr>
				<tr>
					<td class="note" colspan="3">Riassunto informazioni:</td>
				</tr>
				
				<tr>
					<th class="L2">informazione:</td>
					<th class="L2_center">sorgente</td>
					<th class="L2_center">destinazione</td>
				</tr>
				<tr>
					<td class="label">Plugin:</td>
					<td class="content_center">
						<%=srs("N_OBJECTS")%>
					</td>
					<td class="content_center">
						<%=drs("N_OBJECTS")%>
					</td>
				</tr>
				<tr>
					<td class="label">Paginesito:</td>
					<td class="content_center">
						<%=srs("N_PAGINE")%>
					</td>
					<td class="content_center">
						<%=drs("N_PAGINE")%>
					</td>
				</tr>
				<tr>
					<td class="label">Pagine:</td>
					<td class="content_center">
						<%=srs("N_PAGES")%>
					</td>
					<td class="content_center">
						<%=drs("N_PAGES")%>
					</td>
				</tr>
				<tr>
					<td class="label">Layers:</td>
					<td class="content_center">
						<%=srs("N_LAYERS")%>
					</td>
					<td class="content_center">
						<%=drs("N_LAYERS")%>
					</td>
				</tr>
				<% if cIntero(drs("N_OBJECTS"))=0 AND cIntero(drs("N_PAGINE"))=0 AND cIntero(drs("N_PAGES"))=0 AND cIntero(drs("N_LAYERS"))=0 then %>
					<% if request("ESEGUI")="" then %>
							<tr>	
								<td class="footer" colspan="3">
									<input style="width:20%;" type="submit" class="button" name="esegui" value="ESEGUI!!">
								</td>
							</tr>
					<% else 
						
						dconn.begintrans
						
						'aggiunge colonne per il transito dati
						sql = "ALTER TABLE tb_paginesito ADD old_ps_id INT NULL" + vbCrLf + _
							  "ALTER TABLE tb_pages ADD old_page_id INT" + vbCrLf + _
							  "ALTER TABLE tb_objects ADD old_obj_id INT" + vbCrLf
						CALL dconn.execute(sql)
						%>
						</table>
						<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
							<tr><th colspan="3">Esecuzione</th></tr>
							<tr>
								<td class="content">
									copia Pagine Sito
								</td>
								<%
								sql = "INSERT INTO tb_pagineSito (old_ps_id, id_web,archiviata,riservata,id_pagDyn_IT,id_pagDyn_EN,id_pagDyn_FR,id_pagDyn_DE,id_pagDyn_ES,id_pagStage_IT,id_pagStage_EN," +_
									  " id_pagStage_FR,id_pagStage_DE,id_pagStage_ES,nome_ps_IT,nome_ps_EN,nome_ps_FR,nome_ps_DE,nome_ps_ES,PAGE_keywords_IT,PAGE_keywords_EN," +_
									  " PAGE_keywords_FR,PAGE_keywords_DE,PAGE_keywords_ES,PAGE_description_IT,PAGE_description_EN,PAGE_description_FR,PAGE_description_DE," +_
									  " PAGE_description_ES,ps_insData,ps_insAdmin_id,ps_modData,ps_modAdmin_id,nome_ps_interno,id_pagDyn_ru,id_pagStage_ru,nome_ps_ru,PAGE_keywords_ru," +_
									  " PAGE_description_ru,id_pagDyn_cn,id_pagStage_cn,nome_ps_cn,PAGE_keywords_cn,PAGE_description_cn,id_pagDyn_pt,id_pagStage_pt,nome_ps_pt,PAGE_keywords_pt,PAGE_description_pt,indicizzabile) " + _
									  "SELECT id_paginesito, " & dwebs_id & ", archiviata, riservata, id_pagDyn_IT, id_pagDyn_EN, id_pagDyn_FR, id_pagDyn_DE, id_pagDyn_ES, id_pagStage_IT, id_pagStage_EN, id_pagStage_FR, " + _
									  " id_pagStage_DE, id_pagStage_ES, nome_ps_IT, nome_ps_EN, nome_ps_FR, nome_ps_DE, nome_ps_ES, PAGE_keywords_IT, PAGE_keywords_EN, PAGE_keywords_FR, PAGE_keywords_DE, " + _
									  " PAGE_keywords_ES, PAGE_description_IT, PAGE_description_EN, PAGE_description_FR, PAGE_description_DE, PAGE_description_ES, ps_insData, ps_insAdmin_id, ps_modData, " + _
									  " ps_modAdmin_id, nome_ps_interno, id_pagDyn_ru, id_pagStage_ru, nome_ps_ru, PAGE_keywords_ru, PAGE_description_ru, id_pagDyn_cn, id_pagStage_cn, nome_ps_cn, " + _
									  " PAGE_keywords_cn, PAGE_description_cn, id_pagDyn_pt, id_pagStage_pt, nome_ps_pt, PAGE_keywords_pt, PAGE_description_pt, indicizzabile " + _
									  " FROM tb_pagineSito " + _
									  " WHERE id_web = " & swebs_id
								CALL dconn.execute(SQL)
								%>
								<td colspan="2" class="content ok">
									Pagine Sito Importate
								</td>
							</tr>
							<tr>
								<td class="content">
									copia Oggetti
								</td>
								<%
								sql = " INSERT INTO tb_objects " + _
									  " (old_obj_id, id_webs, name_objects, identif_objects, param_list, obj_insData, obj_insAdmin_id, obj_modData, obj_modAdmin_id, obj_type, obj_html_it, obj_html_en, obj_html_fr, obj_html_de, obj_html_es, obj_html_ru, obj_html_cn, obj_html_pt) " + _
									  " SELECT id_objects, " & dwebs_id & ", name_objects, identif_objects, param_list, obj_insData, obj_insAdmin_id, obj_modData, obj_modAdmin_id, obj_type, obj_html_it, obj_html_en, obj_html_fr, obj_html_de, obj_html_es, obj_html_ru, obj_html_cn, obj_html_pt " + _
									  " FROM tb_objects " + _
									  " WHERE id_webs = " & swebs_id
								CALL dconn.execute(SQL)
								%>
								<td colspan="2" class="content ok">
									Oggetti importati
								</td>
							</tr>
							<tr>
								<td class="content">
									copia Pagine
								</td>
								<%
								sql = " INSERT INTO tb_pages (id_PaginaSito, old_page_id, id_webs, id_template, nomepage, template, SfondoColore, SfondoImmagine, Contatore, ContRes, contUtenti, " + _
									  " contCrawler, contAltro, page_insData, page_insAdmin_id, page_modData, page_modAdmin_id, lingua_tmp, lingua, semplificata) " + _
									  " SELECT " + _
									  " (SELECT id_paginesito FROM tb_pagineSito WHERE old_ps_id= tb_pages.id_PaginaSito), " + _
									  " id_page, " & dwebs_id & ", id_template, nomepage, template, SfondoColore, SfondoImmagine, Contatore, ContRes, contUtenti, contCrawler, contAltro, " + _
									  " page_insData, page_insAdmin_id, page_modData, page_modAdmin_id, lingua_tmp, lingua, semplificata " + _
									  " FROM tb_pages " + _
									  " WHERE id_webs = " & swebs_id
								CALL dconn.execute(SQL)
								%>
								<td colspan="2" class="content ok">
									pagine importate
								</td>
							</tr>
							<tr>
								<td class="content">
									corregge template
								</td>
								<%
								sql = " UPDATE tb_pages SET id_template = (SELECT id_page FROM tb_pages tb_template WHERE old_page_id= tb_pages.id_template) " + _ 
									  " WHERE id_webs=" & dwebs_id
								CALL dconn.execute(SQL)
								%>
								<td colspan="2" class="content ok">
									template corretti
								</td>
							</tr>
							<tr>
								<td class="content">
									copia Layers
								</td>
								<%
								sql = " INSERT INTO tb_layers " + _
									  " (id_pag, id_objects, " + _
									  " id_tipo, tipo_contenuto, z_order, nome, visibile, x, y, largo, alto, em_x, em_y, em_largo, em_alto, html, format, testo, aspcode, RTF, CHECKSUM_STILI) " + _
									  " SELECT " + _
									  " (SELECT id_page FROM tb_pages WHERE old_page_id= tb_layers.id_pag), " + _
									  " (SELECT id_objects FROM tb_objects WHERE old_obj_id= tb_layers.id_objects), " + _
									  " id_tipo, tipo_contenuto, z_order, nome, visibile, x, y, largo, alto, em_x, em_y, em_largo, em_alto, html, format, testo, aspcode, RTF, CHECKSUM_STILI " + _
									  " FROM dbo.tb_layers " + _
									  " WHERE id_pag IN (SELECT id_page FROM tb_pages WHERE id_webs=" & swebs_id & ")"
								CALL dconn.execute(SQL)
								%>
								<td colspan="2" class="content ok">
									layers importati
								</td>
							</tr>
							<tr>
								<td class="content">
									corregge collegamento paginesito - pagine
								</td>
								<%
								sql = " UPDATE tb_paginesito SET "
								for each lingua in LINGUE_CODICI
									sql = sql + "id_pagDyn_" + lingua + " = (SELECT id_page FROM tb_pages WHERE old_page_id= tb_paginesito.id_pagDyn_" + lingua + "), " + _ 
												"id_pagStage_" + lingua + " = (SELECT id_page FROM tb_pages WHERE old_page_id= tb_paginesito.id_pagStage_" + lingua + "), " 
								next
								sql = left(sql, len(sql)-2) + " WHERE id_web=" & dwebs_id
								CALL dconn.execute(SQL)
								%>
								<td colspan="2" class="content ok">
									collegamento corretto
								</td>
							</tr>
						<% 
						
						
						'aggiunge colonne per il transito dati
						sql = "ALTER TABLE tb_paginesito DROP COLUMN old_ps_id " + vbCrLf + _
							  "ALTER TABLE tb_pages DROP COLUMN old_page_id" + vbCrLf + _
							  "ALTER TABLE tb_objects DROP COLUMN old_obj_id" + vbCrLf
						CALL dconn.execute(sql)
						
						'response.end
						dconn.committrans%>
						<tr>
							<td class="conten ok" colspan=3">
								Duplicazione sito completata
							</td>
						</tr>
						<tr>
							<td class="footer" colspan="3">
								<a class="button" href="default.asp">FINE</a>
							</td>
						</tr>
					<% end if
				else %>
				
					<tr>	
						<td class="footer alert" colspan="3">
							Il sito di destinazione non &egrave; vuoto!
						</td>
					</tr>
				<% end if 
				srs.close
				drs.close
			end if
		end if %>
	</table>
	</form>
</div>
</body>
</html>
