<% 


'scrive la porzione di HTML che visualizza le statistiche generali del sito.
sub WRITE_StatisticheGenerali(conn, rs, id_webs)
	dim sql
	
	sql = " SELECT * FROM tb_webs WHERE id_webs=" & id_webs
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adAsyncFetch %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption class="border">Statistiche generali di accesso al sito "<%= rs("nome_webs") %>":</caption>
		<tr>
			<td class="label" style="width:30%;" rowspan="4">N&ordm; visitatori dal "<%= rs("contRes") %>"</td>
			<td class="label" style="width:10%;">utenti:</td>
			<td class="content"><%= cIntero(rs("contUtenti")) %></td>
			<td class="content_center" style="width:24%;" rowspan="4">
				<%CALL WriteButton_StatisticheArchiviaAzzera(id_webs) %>
			</td>
		</tr>
		<tr>
			<td class="label">motori di ricerca:</td>
			<td class="content"><%= cIntero(rs("contCrawler")) %></td>
		</tr>
			<td class="label">altri visitatori:</td>
			<td class="content"><%= cIntero(rs("contAltro")) %></td>
		</tr>
		<tr>
			<td class="label">totale:</td>
			<td class="content"><%= cIntero(rs("contatore")) %></td>
		</tr>
	</table>
	<% rs.close
end sub


'genera il link per aprire la finestra di archiviazione / azzeramento
sub WriteButton_StatisticheArchiviaAzzera(id_webs) %>
	<a HREF="javascript:void(0);" onClick="OpenAutoPositionedWindow('SitoAnalisiStatReset.asp?WEB_ID=<%= id_webs %>', 'reset_contatori', 500, 250)" class="button_L2_block">
		ARCHIVIA STATISTICHE<br>
		ED<br>
		AZZERA CONTATORI
	</a>
<% end sub


sub StatisticheArchiviaAzzera(conn, id_webs)
	dim sql
	
	'salva testata del sito nello storico
	sql = " INSERT INTO tb_storico_webs(sw_insData, sw_insAdmin_id, sw_modData, sw_modAdmin_id, sw_webs_id, sw_data, sw_contatore, sw_contUtenti, sw_contCrawler, sw_contAltro) " & _
		  " SELECT "& SQL_Now(conn) &", "& session("ID_ADMIN") &", "& SQL_Now(conn) &", "& session("ID_ADMIN") & _
				   ", id_webs, "& SQL_Now(conn) &", contatore, contUtenti, contCrawler, contAltro FROM tb_webs WHERE id_webs=" & id_webs
	CALL conn.execute(sql, , adExecuteNoRecords)
		
	'salva righe con storico delle pagine
	sql = " INSERT INTO tb_storico_pages(sp_page_id, sp_pagineSito_id, sp_nomepage, sp_lingua, sp_contatore, sp_contUtenti, sp_contCrawler, sp_contAltro, sp_sw_id) "& _
  		  " SELECT id_page, ( SELECT TOP 1 id_pagineSito FROM tb_pagineSito "
		  if DB_Type(conn) = DB_ACCESS then
			sql = sql + _
					" WHERE id_pagDyn_it=id_page OR id_pagDyn_en=id_page OR id_pagDyn_fr=id_page " + _
					" OR id_pagDyn_es=id_page OR id_pagDyn_de=id_page), "
		  else
			sql = sql + _
					" WHERE id_pagDyn_it=id_page OR id_pagDyn_en=id_page OR id_pagDyn_fr=id_page " + _
					" OR id_pagDyn_es=id_page OR id_pagDyn_de=id_page OR id_pagDyn_ru=id_page OR id_pagDyn_cn=id_page OR id_pagDyn_pt=id_page), "
		  end if
		  sql = sql + " nomepage, lingua, contatore, contUtenti, contCrawler, contAltro, " + _
					  " (SELECT TOP 1 sw_id FROM tb_storico_webs WHERE sw_webs_id=id_webs ORDER BY sw_id DESC) " + _
					  " FROM tb_pages WHERE (SELECT TOP 1 id_pagineSito FROM tb_pagineSito "
		  if DB_Type(conn) = DB_ACCESS then		  
			sql = sql + _
					" WHERE id_pagDyn_it=id_page OR id_pagDyn_en=id_page OR id_pagDyn_fr=id_page " + _
					" OR id_pagDyn_es=id_page OR id_pagDyn_de=id_page) "
		  else
			sql = sql + _
					" WHERE id_pagDyn_it=id_page OR id_pagDyn_en=id_page OR id_pagDyn_fr=id_page " + _
					" OR id_pagDyn_es=id_page OR id_pagDyn_de=id_page OR id_pagDyn_ru=id_page OR id_pagDyn_cn=id_page OR id_pagDyn_pt=id_page) "
		  end if
		  sql = sql + "> 0 AND id_webs=" & id_webs
	CALL conn.execute(sql, , adExecuteNoRecords)
		
	'salva righe storico dell'indice
	sql = " INSERT INTO tb_storico_index(si_sw_id, si_idx_id, si_idx_padre_id, si_idx_ordine_assoluto, si_co_F_key_id,"& _
			  						   " si_idx_foglia, si_idx_livello, si_tab_id, si_tab_name,"& _
									   " si_titolo_it, si_titolo_en, si_titolo_fr, si_titolo_de, si_titolo_es, "
    if DB_Type(conn) = DB_SQL then
		sql = sql + "si_titolo_ru, si_titolo_cn, si_titolo_pt," + _
					"si_link_ru, si_link_cn, si_link_pt,"
    end if
	sql = sql + " si_link_pagina_id, si_link_it, si_link_en, si_link_fr, si_link_de, si_link_es, "& _
				" si_contatore, si_contUtenti, si_contCrawler, si_contAltro)"& _
		  	" SELECT (SELECT TOP 1 sw_id FROM tb_storico_webs WHERE sw_webs_id = "& id_webs &" ORDER BY sw_id DESC),"& _
				   " idx_id, idx_padre_id, idx_ordine_assoluto, co_F_key_id, idx_foglia, idx_livello, tab_id, tab_name,"& _
				   " co_titolo_it, co_titolo_en, co_titolo_fr, co_titolo_de, co_titolo_es, "
				    if DB_Type(conn) = DB_SQL then
						sql = sql + "co_titolo_ru, co_titolo_cn, co_titolo_pt, "& _
								  "idx_link_url_ru, idx_link_url_cn, idx_link_url_pt,"
					end if
				   sql = sql + " idx_link_pagina_id, idx_link_url_it, idx_link_url_en, idx_link_url_fr, idx_link_url_de, idx_link_url_es, "& _
				   " idx_contatore, idx_contUtenti, idx_contCrawler, idx_contAltro"& _
			" FROM (tb_contents_index i INNER JOIN tb_contents c ON i.idx_content_id = c.co_id)"& _
			" INNER JOIN tb_siti_tabelle t ON c.co_F_table_id = t.tab_id "
	CALL conn.execute(sql, , adExecuteNoRecords)
	
	'azzera contatori del sito
	sql = " UPDATE tb_webs SET contRes = "& SQL_Now(conn) &", contatore = 0, contUtenti=0, contCrawler=0, contAltro=0, " & _
		  SetUpdateParamsSQL(conn, "webs_", false) & _
		  " WHERE id_webs=" & id_webs
	CALL conn.execute(sql, , adExecuteNoRecords)
	
	'azzera contatori della pagina
	sql = " UPDATE tb_pages SET contRes = "& SQL_Now(conn) &", contatore = 0, contUtenti=0, contCrawler=0, contAltro=0, " & _
		  SetUpdateParamsSQL(conn, "page_", false) & _
		  " WHERE id_webs=" & id_webs
	CALL conn.execute(sql, , adExecuteNoRecords)
	
	'azzera righe storico dell'indice
	sql = " UPDATE tb_contents_index SET idx_contRes = "& SQL_Now(conn) &", idx_contatore = 0, idx_contUtenti=0, idx_contCrawler=0, idx_contAltro=0, " & _
		  SetUpdateParamsSQL(conn, "idx_", false) & _
		  " WHERE idx_webs_id =" & id_webs
	CALL conn.execute(sql, , adExecuteNoRecords)
end sub


'restituisce la tabella che visualizza le statistiche
sub GetStatTable()
	if request.querystring("AllSats") = "1" then
		session("AllSats") = "1"		'visualizza tutte le statistiche
	elseif request.querystring("AllSats") = "0" then
		session("AllSats") = ""
	end if %>
	
	<style type="text/css">
		#legenda {
			<% if session("AllSats") <> "" then %>
				top: 155px !important;
			<% else %>
				top: 185px !important;
			<% end if %>
		}
		#td_IndexAlbero_ {
			height: 100px;
			vertical-align: top;
		}
		
		table.stat {
			position: absolute;
			left: 540px;
			top: 118px;
			width: 180px;
		}
		
		<% if session("AllSats") <> "" then %>
		
			table.stat_mini {
				width:300px;
				border-top-width:1px;
				margin-left: 10px;
				
			}
			
			td.stat_mini_td{
				width:45px;
			}
			
		<% end if %>
	</style>
		
	<script type="text/javascript">
		function Stat(contUtenti, contCrawler, contAltro, contatore) {
			if (document.getElementById('contUtenti')) {
				var tdContatore;
				tdContatore = document.getElementById('contUtenti');
				tdContatore.innerHTML = contUtenti;
				if (contUtenti != 0)
					tdContatore.className = 'content_center';
				else
					tdContatore.className = 'content_center notes';
				
				tdContatore = document.getElementById('contCrawler');
				tdContatore.innerHTML = contCrawler;
				if (contCrawler != 0)
					tdContatore.className = 'content_center';
				else
					tdContatore.className = 'content_center notes';
					
				tdContatore = document.getElementById('contAltro');
				tdContatore.innerHTML = contAltro;
				if (contAltro != 0)
					tdContatore.className = 'content_center';
				else
					tdContatore.className = 'content_center notes';
					
				tdContatore = document.getElementById('contatore');
				tdContatore.innerHTML = contatore;
				if (contatore != 0)
					tdContatore.className = 'content_center';
				else
					tdContatore.className = 'content_center notes';
				
			}
		}
	</script>
		
	<table cellpadding="0" cellspacing="1" class="tabella_madre stat">
		<% if session("AllSats") <> "" then %>
			<caption class="border">STATISTICHE DI ACCESSO</caption>
			<tr>
				<td class="content_center" colspan="4">
					<a href="?AllSats=0" class="button_L2_block">NASCONDI TUTTE</a>
				</td>
			</tr>
		<% else %>
			<caption>STATISTICHE DI ACCESSO</caption>
			<tr>
				<th class="center">utenti</th>
				<th class="center">crawler</th>
				<th class="center">altro</th>
				<th class="center">totale</th>
			</tr>
			<tr>
				<td id="contUtenti" class="content_center notes" style="height:14px;">0</td>
				<td id="contCrawler" class="content_center notes">0</td>
				<td id="contAltro" class="content_center notes">0</td>
				<td id="contatore" class="content_center notes">0</td>
			</tr>
			<tr>
				<td class="content_center" colspan="4">
					<a href="?AllSats=1" class="button_L2_block">VISUALIZZA TUTTE</a>
				</td>
			</tr>
		<% end if %>
	</table>
<% End sub


'restituisce la tabella che visualizza le statistiche per riga
Function GetStatNome(nome, tab_name, idx_id, idx_contUtenti, idx_contCrawler, idx_contAltro, idx_contatore)
	idx_contUtenti = CIntero(idx_contUtenti)
	idx_contCrawler = CIntero(idx_contCrawler)
	idx_contAltro = CIntero(idx_contAltro)
	idx_contatore = CIntero(idx_contatore)
	
	if session("AllSats") <> "" then
		GetStatNome = nome & _
			"<table id='tab_" & idx_id & "' cellpadding='0' cellspacing='1' class='tabella_madre stat_mini'>" & _
				"<tr>"
		if (idx_contUtenti + idx_contCrawler + idx_contAltro + idx_contatore) = 0 then
			GetStatNome = replace(GetStatNome, "tabella_madre", "tabella_madre_disabled")
			if instr(1, tab_name, "tb_contents", vbTextCompare)>0 then
				GetStatNome = GetStatNome + "<td class='notes_disabled' style='text-align:center;' title='I raggruppamenti non sono &quot;visitabili&quot; dagli utenti nella navigazione perch&egrave; non hanno un proprio indirizzo.\nRappresentano solo dei punti di aggregazione dei contenuti.'>nessuna visita (raggruppamento)</td>"
			else
				GetStatNome = GetStatNome + "<td class='content_center notes_disabled'>nessuna visita</td>"
			end if
		else
			GetStatNome = GetStatNome + _
				"<td class='label_no_width'>utenti: </td>"& _
				"<td class='content_b stat_mini_td" & IIF(idx_contUtenti=0, " notes", "") & "'>"& idx_contUtenti &"</td>"& _
				"<td class='label_no_width'>crawler: </td>"& _
				"<td class='content_b stat_mini_td" & IIF(idx_contCrawler=0, " notes", "") & "'>"& idx_contCrawler &"</td>"& _
				"<td class='label_no_width'>altro: </td>"& _
				"<td class='content_b stat_mini_td" & IIF(idx_contAltro=0, " notes", "") & "'>"& idx_contAltro &"</td>"& _
				"<td class='label_no_width'>totale: </td>"& _
				"<td class='content_b stat_mini_td" & IIF(idx_contatore=0, " notes", "") & "'>"& idx_contatore &"</td>"
		end if
		GetStatNome = GetStatNome & _
				"</tr>" & _
			"</table>"
	else
		GetStatNome = _
			"<a onmouseover=\""Stat("& idx_contUtenti &", "& idx_contCrawler &", "& idx_contAltro &", "& idx_contatore &")\"" onmouseout=\""Stat(0, 0, 0, 0)\"">" & _
			nome & _
			"</a>"
	end if
End Function


 %>