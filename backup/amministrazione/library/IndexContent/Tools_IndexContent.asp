<!--#INCLUDE FILE="ClassIndex.asp"-->
<!--#INCLUDE FILE="ClassContent.asp"-->

<%
'file che contiene le funzioni per l'inclusione in tutti gli applicativi.


'.............................................................................................................................
'funzione che permette l'aggiornamento del contenuto collegato al record corrente,
'o l'inserimento automatico del contenuto e la sua indicizzazzione se sono attive ed applicabili delle pubblicazioni automatiche.
'restituisce l'id del contenuto
'table				nome tabella sorgente del contenuto da aggiornare
'keyvalue			valore della chiave del record di origine del contenuto da cui prelevare i dati
'forzaContenuto		se true: inserisce(ed aggiorna) sempre e comunque la proiezione del contenuto sorgente nella tabella tb_contents
'					indipendentemente se questo e' pubblicato o meno.
'.............................................................................................................................
Function Index_UpdateItem(conn, table, KeyValue, forzaContenuto)
	Index_UpdateItem = Index_UpdateItemTransaction(conn, table, KeyValue, forzaContenuto, true)
End Function

'.............................................................................................................................
'funzione che permette l'aggiornamento del contenuto collegato al record corrente,
'o l'inserimento automatico del contenuto e la sua indicizzazzione se sono attive ed applicabili delle pubblicazioni automatiche.
'restituisce l'id del contenuto
'table				nome tabella sorgente del contenuto da aggiornare
'keyvalue			valore della chiave del record di origine del contenuto da cui prelevare i dati
'forzaContenuto		se true: inserisce(ed aggiorna) sempre e comunque la proiezione del contenuto sorgente nella tabella tb_contents
'					indipendentemente se questo e' pubblicato o meno.
'beAtomic			se true: mantiene la transazione esterna aperta, senza modificarla.
'					se false: esegue il commit prima di iniziare l'aggiornamento dell'indice.
'.............................................................................................................................
Function Index_UpdateItemTransaction(conn, table, KeyValue, forzaContenuto, beAtomic)
	dim sql, ContentId, Published, lingua
	dim rst, rsp, rspr, rsc, rsCat, rsx, rsf , rsTag, value, RicalcolaAlbero

	conn.CommandTimeout = 10000000

	set Index.conn = conn
	set Index.content.conn = conn
	set index.dizionario = server.createobject("Scripting.Dictionary")
	
	'resetta la transazione se richiesto per salvare i dati immessi dall'utente prima delle elaborazioni dell'indice.
	if not beAtomic then
		conn.committrans
		conn.begintrans
	end if
	
	set rst = Server.CreateObject("ADODB.recordset")
	set rsp = Server.CreateObject("ADODB.recordset")
	set rspr = Server.CreateObject("ADODB.recordset")
	set rsc = Server.CreateObject("ADODB.recordset")
	set rsCat = Server.CreateObject("ADODB.recordset")
    set rsx = Server.CreateObject("ADODB.recordset")
    set rsf = Server.CreateObject("ADODB.recordset")
	set rsTag = Server.CreateObject("ADODB.recordset")
	
	'verifica se la tabella e' descritta ed indicizzabile
	sql = "SELECT * FROM tb_siti_tabelle WHERE"
	if CIntero(table) > 0 then
		sql = sql &" tab_id = "& cIntero(table)
	else
		sql = sql &" tab_name LIKE '" & ParseSql(table, adChar) & "'"
	end if
	rst.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

    while not rst.eof
        ContentId = 0
		'tabella trovata nell'indice
		
		'verifica se il record e' gia' presente nei contenuti
		sql = "SELECT co_id FROM tb_contents WHERE co_F_table_id=" & cIntero(rst("tab_id")) & " AND co_F_Key_id=" & cIntero(KeyValue)
		rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

		if not rsc.eof OR forzaContenuto then
			if not rsc.eof then
				ContentId = rsc("co_id")
			end if

			'aggiorna i dati del contenuto
			CALL Index.Content.GeneraDaTabella(ContentId, rst, KeyValue)

			'aggiorno il link dell'indice se collegato al contenuto
			if CString(rst("tab_field_url_it")) <> "" then
				CALL index.content.SalvaIndexLink(contentId, rsx, rst("tab_parametro"))
			end if
		end if

		if session("ERRORE") = "" then
			'verifica se esistono pubblicazioni automatiche per il tipo contenuto
			sql = ""
			for each lingua in Application("LINGUE")
				sql = sql & " tab_field_url_"& lingua & ","
			next
			sql = " SELECT "&sql&" pub_id, pub_pagina_id, tab_parametro, pub_categoria_field, pub_padre_index_id, pub_categoria_tabella_id, pub_field_principale, " & _
				  " 		tab_from_sql, pub_filtro_pubblicazione, tab_field_chiave, " & _
				  "			   (" & SQL_IF(conn, "( SELECT COUNT(idx_id) FROM tb_contents_index " + _
	                                              " INNER JOIN rel_index_pubblicazioni ON tb_contents_index.idx_id = rel_index_pubblicazioni.rip_idx_id " + _
	                                              " WHERE rel_index_pubblicazioni.rip_pub_id = tb_siti_tabelle_pubblicazioni.pub_id AND idx_content_id=" & cIntero(ContentId) & ") > 0", "0", "1") & ") AS CON_CONTENUTI" & _
	              " FROM tb_siti_tabelle_pubblicazioni " & _
				  " INNER JOIN tb_siti_tabelle ON tb_siti_tabelle_pubblicazioni.pub_tabella_id = tb_siti_tabelle.tab_id " +_
				  " WHERE pub_tabella_id=" & cIntero(rst("tab_id"))
	        rsp.CursorLocation = adUseClient
			rsp.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	        rsp.sort = " CON_CONTENUTI "
			while not rsp.eof			
	            RicalcolaAlbero = 0

				'verifica se sono applicabili per mezzo dei filtri al record corrente
				sql = Pubblicazione_GetQuery(conn, rsp, 0, KeyValue)
				rspr.open sql, conn, adOpenStatic, adLockOptimistic
				if not rspr.eof then        'FILTRI OK: record da pubblicare automaticamente
	                
					'pubblicazione valida per il record corrente: esegue pubblicazione sull'indice
					if cInteger(ContentId)=0 then
						'inserisce anche un nuovo contenuto
						CALL Index.Content.GeneraDaTabella(ContentId, rst, KeyValue)
					end if
					
					if session("ERRORE") = "" then
						index.dizionario("idx_content_id") = contentId
						
						'gestione link
						value = false
						'response.write TypeName(index.Content.dizionario)
						for each lingua in Application("LINGUE")
							'response.write "tab_field_url= """ & rsp("tab_field_url_" + lingua) & """<br>"
							if CString(rsp("tab_field_url_" + lingua).value) <> "" AND _
							   index.Content.dizionario.Exists("co_link_url_" + lingua) then
								if CString(index.Content.dizionario("co_link_url_" + lingua)) <> "" then
									value = true
									exit for
								end if
							end if
						next
						
						'se non ho il valore impostato dalla tabella provo ad impostarlo dalla pubblicazione
						if NOT value AND CIntero(rsp("pub_pagina_id")) > 0 then
							CALL LinkCalculate(conn, "co", index.content.dizionario, rsp, "pub_pagina_id", "", rsp("tab_parametro"))
						end if
		                
						'pubblicazione tramite categorizzazione
						if cString(rsp("pub_categoria_field"))<>"" then
							if cInteger(rspr("pub_categoria_field"))>0 then
		                        if cInteger(rsp("pub_padre_index_id"))>0 then
								    sql = " SELECT idx_id FROM (tb_contents_index i"& _
									      " INNER JOIN tb_contents c ON i.idx_content_id = c.co_id)"& _
		    							  " WHERE co_F_table_id = "& cIntero(rsp("pub_categoria_tabella_id")) & _
			    						  " AND co_F_key_id = "& cIntero(rspr("pub_categoria_field")) & _
		                                  " AND " & SQL_IdListSearch(conn, "idx_tipologie_padre_lista", cIntero(rsp("pub_padre_index_id")))
		                            rsCat.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		    						if not rsCat.eof then
		    							'pubblicazione del padre trovata
		    							index.dizionario("idx_padre_id") = rsCat("idx_id")
										if cString(rspr("principale"))<>"" then
											index.dizionario("idx_principale") = CBoolean(rspr("principale"), false)
										end if
		    						    index.SetIndexFromContent()
		    							
		                                dim IdxIdOld, IdxIdDest
		                                'verifica se esiste gia' il record pubblicato per questa pubblicazione o un record autopubblicato per un'altra pubblicazione non piu' valida.
		                                sql = " SELECT TOP 1 idx_id FROM  tb_contents_index LEFT JOIN rel_index_pubblicazioni ON tb_contents_index.idx_id = rel_index_pubblicazioni.rip_idx_id " + _
		                                      " WHERE idx_content_id=" & cIntero(contentId) & _
		                                            " AND ( rip_pub_id =" & cIntero(rsp("pub_id")) & " OR " & _
		                                                  " (" & SQL_IsTrue(conn, "idx_autopubblicato") & " AND idx_id NOT IN (SELECT rip_idx_id FROM rel_index_pubblicazioni)) )" & _
		                                      " ORDER BY rip_id DESC "
		                                IdxIdOld = cIntero(GetValueList(conn, rsx, sql))
		                                
		                                'verifica se esiste gia' il contenuto pubblicato nella nuova posizione.
		                                sql = " SELECT TOP 1 idx_id FROM tb_contents_index " & _
		                                      " WHERE idx_content_id = " & cIntero(contentId) & _
		    								  " AND idx_padre_id = " & cIntero(rsCat("idx_id"))
		                                IdxIdDest = cIntero(GetValueList(conn, rsx, sql))
										
										
		                                if IdxIdDest > 0 AND IdxIdOld>0 AND IdxIdOld <> IdxIdDest then
		                                    'conflitto: la pubblicazione esite gia' e deve essere spostata dove ce n'e' un'altra.
		                                    'verifico se la voce automatica precedente ha figli
		                                    sql = "SELECT COUNT(*) FROM tb_contents_index WHERE idx_padre_id=" & cIntero(IdxIdOld)
		                                    if cIntero(GetValueList(conn, rsx, sql))=0 then
		                                        'nessun figlio: uso quella di destinazione e sblocco quella vecchia.
		                                        value = IdxIdDest
		                                        
		                                        sql = " DELETE FROM rel_index_pubblicazioni " & _
		                                              " WHERE rip_pub_id = "& cIntero(rsp("pub_id")) & _
		                                              " AND rip_idx_id=" & cIntero(IdxIdOld)
		                                        CALL conn.execute(sql, ,adExecuteNoRecords)
		                                    else
		                                        'trovati figli: nella vecchia voce: verifica quella di "destinazione"
		                                        sql = "SELECT COUNT(*) FROM tb_contents_index WHERE idx_padre_id=" & cIntero(IdxIdDest)
		                                        if cIntero(GetValueList(conn, rsx, sql))=0 then
		                                            'nessun figlio trovato: usa la voce autopubblicata.
		                                            value = IdxIdOld
		                                            RicalcolaAlbero = IdxIdOld
		                                            
		                                            'marchia voce esistente come da cancellare.
		                                            sql = "UPDATE tb_contents_index SET idx_autopubblicato=1 WHERE idx_id=" & cIntero(IdxIdDest)
		                                            CALL conn.execute(sql, ,adexecuteNoRecords)
		                                            sql = " DELETE FROM rel_index_pubblicazioni WHERE rip_idx_id=" & cIntero(IdxIdDest)
		                                            CALL conn.execute(sql, ,adexecuteNoRecords)
		                                            
		                                        else
		                                            'entrambe hanno figli: considera valida quella di destinazione: 
		                                            'sposta i figli della Old nella nuova: solo figli che non sono gia' presenti nella destinazione
		                                            sql = " SELECT * FROM tb_contents_index " & _
		                                                  " WHERE idx_content_id NOT IN (SELECT idx_content_id FROM tb_contents_index ci WHERE ci.idx_padre_id=" & cIntero(IdxIdDest) & ")" & _
		                                                  " AND idx_padre_id = " & cIntero(IdxIdOld)
		                                            rsf.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		                                            
		                                            while not rsf.eof
		                                                'sposta ogni figlio
		                                                rsf("idx_padre_id") = IdxIdDest
		                                                rsf("idx_autopubblicato") = false
		                                                rsf.update
		                                                
		                                                'rimuove figlio da blocco pubblicazioni
		                                                sql = "DELETE FROM rel_index_pubblicazioni WHERE rip_idx_id = "& cIntero(rsf("idx_id"))
		                                                CALL conn.execute(sql, ,adexecuteNoRecords)
		                                                
		                                                rsf.movenext
		                                            wend
		                                            rsf.close
		                                            
		                                            value = IdxIdDest
		                                            RicalcolaAlbero = IdxIdDest
		                                            
		                                            'marchia voce precedente come da cancellare.
		                                            sql = "UPDATE tb_contents_index SET idx_autopubblicato=1 WHERE idx_id=" & cIntero(IdxIdOld)
		                                            CALL conn.execute(sql, ,adexecuteNoRecords)
		                                            sql = " DELETE FROM rel_index_pubblicazioni WHERE rip_idx_id=" & cIntero(IdxIdOld)
		                                            CALL conn.execute(sql, ,adexecuteNoRecords)
		                                            
		                                        end if
		                                    end if
		                                
		                                elseif IdxIdDest > 0 then
		                                    'esiste solo la destinazione: usa quella da bloccare.
		                                    value = IdxIdDest
		                                elseif IdxIdOld>0 then
		                                    'esiste gia' la pubblicazione: deve aggiornarla.
		                                    value = IdxIdOld
		                                end if
										CALL index.SalvaPubblicazione(value, rsp("pub_id"))
										
		    							Published = true
		    						else
		    							Published = false
		    						end if
		    						rsCat.close
		                            
		                        else
		                            Published = false
		                        end if
							else
								Published = false
							end if
						else
							Published = false
						end if

						'pubblicazione diretta se non gia' pubblicato tramite padre
						if not Published then
							if CIntero(rsp("pub_padre_index_id")) > 0 then			'pubblicazione diretta
								index.dizionario("idx_padre_id") = rsp("pub_padre_index_id")
								if cString(rspr("principale"))<>"" then
									index.dizionario("idx_principale") = CBoolean(rspr("principale"), false)
								end if
								index.SetIndexFromContent()
								
								sql = " SELECT idx_id FROM tb_contents_index"& _
									  " WHERE idx_content_id = "& cIntero(contentId) & _
									  " AND idx_padre_id = "& cIntero(rsp("pub_padre_index_id"))
								CALL index.SalvaPubblicazione(GetValueList(conn, rsx, sql), rsp("pub_id"))
		                        
		    					Published = true
							end if
						end if
		                
		                if RicalcolaAlbero > 0 then
		                    'ricalcola ed aggiorna l'albero a partire dal livello modificato
		                    Index.operazioni_ricorsive_tipologia(RicalcolaAlbero)
		                end if
					end if		'fine errore su generazione contenuto
	                
				else        'FILTRI FALLITI: record non pubblicato automaticamente
	                Published = false
				end if		'fine filtri
				rspr.close
			
	            if not Published then
	                'rimuove eventuale collegamento tra pubblicazione e contenuto perche' non piu' valida
	                sql = " DELETE FROM rel_index_pubblicazioni " & _
	                      " WHERE rip_pub_id = "& cIntero(rsp("pub_id")) & _
	                      "     AND rip_idx_id IN (SELECT idx_id FROM tb_contents_index WHERE idx_content_id=" & cIntero(ContentId) & ")"
	                CALL conn.execute(sql,,adexecuteNoRecords)
	            end if
	            
				rsp.movenext
			wend		'fine pubblicazioni presenti
			rsp.close
		end if			'fine errore su generazione contenuto
		
		rsc.close
		
		
		
		'inserimento automatico tag
		if rst("tab_tags_abilitati") then
			sql = " SELECT "
			for each lingua in Application("LINGUE")
			
				'codice tags separati da virgola
				if CString(rst("tab_tags_fields_csv_" & lingua))<>"" then
					sql = sql & "(" & rst("tab_tags_fields_csv_" & lingua) & ") AS tags_csv_" & lingua & ", "
				else
					sql = sql & " '' AS tags_csv_" & lingua & ", "
				end if
				
				'codice tags separati da spazio
				if CString(rst("tab_tags_fields_ssv_" & lingua))<>"" then
					sql = sql & "(" & rst("tab_tags_fields_ssv_" & lingua) & ") AS tags_ssv_" & lingua & ", "
				else
					sql = sql & " '' AS tags_ssv_" & lingua & ", "
				end if
			next
			sql = Left(sql, (len(sql)-2))
			sql = sql & " FROM " & rst("tab_from_sql") & " WHERE " & rst("tab_field_chiave") & " = " & KeyValue
			
			dim rs
			set rs = Server.CreateObject("ADODB.recordset")
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

			if not rs.eof then
				CALL Index.Content.RemoveTags(CIntero(contentId), true)
				for each lingua in Application("LINGUE")
				
					'inserisce codice tags separati da virgola
					if CString(rs("tags_csv_" & lingua))<>"" then
						CALL Index.Content.SaveTags(CIntero(contentId), rs("tags_csv_" & lingua), lingua, true, ",")
					end if
					
					'inserire codice tags separati da spazio
					if CString(rs("tags_ssv_" & lingua))<>"" then
						CALL Index.Content.SaveTags(CIntero(contentId), rs("tags_ssv_" & lingua), lingua, true, " ")
					end if
				next
			end if
			rs.close	
			
			sql="select * from tb_siti_tabelle_tag_query where tq_tab_id=" & cIntero(table)
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			
			dim query,campo
			while not rs.eof		
				query=rs("tq_query")					
				query=Replace(query,"<id>",request("co_F_key_id"))						
				for each lingua in Application("LINGUE")
					sql=Replace(query,"<lingua>",lingua)				
					rsTag.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText					
					while not rsTag.eof						
						for each campo in rsTag.Fields
							if cString(campo)<>"" then
								CALL Index.Content.SaveTags(CIntero(contentId), campo, lingua, true, rs("tq_separatore"))
							end if	
						next
						rsTag.movenext
					wend					
					rsTag.close									
				next				
				rs.movenext								
			wend
			rs.close
			
		end if
		
		
		'aggiorna dati di ritorno del contenuto sulla tabella originale
		if cString(rst("tab_return_url_name")) <> "" then
			dim url, sqlUrl
			sqlUrl = ""
			for each lingua in Application("LINGUE")
				if cString(rst("tab_field_return_url_" & lingua)) <>"" then
					
					'c'è l'url di ritorno da impostare: recupera l'url principale del contenuto
					sql = " SELECT TOP 1 idx_link_url_rw_" & lingua & _
						  " FROM v_indice " & _
						  " WHERE co_id = " & contentId & _
						  " ORDER BY idx_principale DESC, visibile_assoluto DESC, idx_id" 'aggiunto idx_id per problema url di ritorno nel caso 2 o più contenuti "principali" (in tal caso ha la precedenza il più vecchio), Giacomo 30/05/2013
					url = cString(GetValueList(conn, rsCat, sql))
					if url <> "" then
						'imposta l'url di ritorno dall'indice al conutenuto originale.
						sqlUrl = sqlUrl & IIF(sqlUrl<>"", ", ", "") & rst("tab_field_return_url_" & lingua) & "='" & ParseSQL(url, adChar) & "'"
					end if
				end if
			next
			
			if cString(rst("tab_field_return_foto_thumb")) <>"" then	
				'c'è foto thumb di ritorno da impostare: recupera il percorso principale del contenuto
				sql = " SELECT TOP 1 co_foto_thumb " & _
					  " FROM tb_contents " & _
					  " WHERE co_id = " & contentId & _
					  " ORDER BY co_data_pubblicazione DESC "
				url = cString(GetValueList(conn, rsCat, sql))

				if url <> "" then
					'imposta foto thumb di ritorno dall'indice al conutenuto originale.
					sqlUrl = sqlUrl & IIF(sqlUrl<>"", ", ", "") & rst("tab_field_return_foto_thumb") & "='" & ParseSQL(url, adChar) & "'"
				end if
			end if
			
			if sqlUrl <> "" then
				'aggiorna campi url di ritorno
				sql = " UPDATE " & rst("tab_return_url_name") & _
					  " SET " &  sqlUrl &  _
					  " WHERE " & rst("tab_field_chiave") & "=" & KeyValue
				CALL conn.execute(sql)
			end if
			
		end if
        rst.movenext
    wend
	rst.close
	
    'ripulisce pubblicazioni dell'indice "autopubblicate" ma non relazionate ad alcuna pubblicazione automatica
    sql = " SELECT idx_id FROM tb_contents_index WHERE " & SQL_IsTrue(conn, "idx_autopubblicato") & _
          " AND idx_id NOT IN (SELECT rip_idx_id FROM rel_index_pubblicazioni) " & _
		  " ORDER BY idx_livello DESC "
    rsCat.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	while not rsCat.eof
        'cancella voce e figli
		CALL Index.Delete(rsCat("idx_id"))
        rsCat.movenext
    wend
    rsCat.close
	
	
	'cancello i contenuti che non hanno più ragione di esistere
	sql = " SELECT co_id, tab_field_chiave, tab_from_sql " & _
		  " FROM tb_contents INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " & _
		  " WHERE co_F_key_id = " & KeyValue & " AND tb_siti_tabelle.tab_name LIKE "
		  if CIntero(table) > 0 then
			sql = sql & " (SELECT tab_name FROM tb_siti_tabelle WHERE tab_id = " & cIntero(table) & ") "
		  else
			sql = sql & " '" & ParseSql(table, adChar) & "'"
		  end if
	rsCat.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	while not rsCat.eof
        sql = " SELECT " & rsCat("tab_field_chiave") & " FROM " & rsCat("tab_from_sql")
		if inStr(sql, "WHERE") > 0 then
			sql = sql & " AND " & rsCat("tab_field_chiave") & " = " & KeyValue
		else
			sql = sql & " WHERE " & rsCat("tab_field_chiave") & " = " & KeyValue
		end if	
			  
		if cIntero(GetValueList(conn, NULL, sql)) = 0 then ' se il contenuto non dovrebbe esistere, secondo l'attuale definizione della tabella, ...
			sql = "SELECT TOP 1 idx_id FROM tb_contents_index WHERE NOT" & SQL_isTrue(conn, "idx_autopubblicato") & " AND idx_content_id = " & rsCat("co_id")
			if cIntero(GetValueList(conn, NULL, sql)) = 0 then '...controllo se è già pubblicato sull'indice:
				CALL Index.Content.Delete(rsCat("co_id")) ' se non è presente sull'indice, oppure è presente attraverso pubblicazioni automatiche, cancello il contenuto.
			end if
		end if
        rsCat.movenext
    wend
    rsCat.close	
	
	'aggiorno i dati dataModifica e IDModifica
	if DB_Type(conn) <> DB_Access AND _
	   table <> "tb_webs" then
	   'modifica fatta il 06/02/2012 da Nicola
	   
		sql = "UPDATE tb_contents_index SET " & SetUpdateParamsSQL(conn, "idx_", false) & " WHERE idx_content_id = "& cIntero(contentId)
		conn.execute(sql)
	end if
	
    Index_UpdateItemTransaction = CIntero(contentId)

	set rst = nothing
	set rsp = nothing
	set rspr = nothing
	set rsc = nothing
	set rsx = nothing
	set rsf = nothing
	set rsCat = nothing

End Function



'.............................................................................................................................
'Funzione che costruisce la query di lettura di una pubblicazione dalla sua definizione.
'.............................................................................................................................
function Pubblicazione_GetQuery(conn, rs, PubId, KeyValue)
	dim sql, rsql, field, i, filters, fieldValue
	if cInteger(PubId) > 0 then
		sql = " SELECT * FROM tb_siti_tabelle_pubblicazioni INNER JOIN tb_siti_tabelle ON tb_siti_tabelle_pubblicazioni.pub_tabella_id = tb_siti_tabelle.tab_id " +_
			  " WHERE pub_id=" & cIntero(PubId)
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	end if
	if not rs.eof then
		rsql = ""
		for each field in rs.fields
			if instr(1, field.name, "_field", vbTextCompare)>0 then
				fieldValue = field.value
				
				if cString(fieldValue)<>"" then
					rsql = rsql + IIF(rsql<>"", ", ", "") + "("
					if left(trim(fieldValue), 1) = "'" AND DB_Type(conn) = DB_sql then
						'il campo contiene una stringa: antepone la codifica nazionale alla stringa
						rsql = rsql + "N" + fieldValue
					else
						rsql = rsql + fieldValue
					end if
					rsql = rsql + ") AS [" + field.name + "]"
				end if
			end if
		next

		rsql = "SELECT " & IIF(cIntero(KeyValue) = 0, " TOP 1 ", "") + _
			   rsql + ", ("& IIF(CString(rs("pub_field_principale")) = "", "0", rs("pub_field_principale")) &") AS principale" + _
			   " FROM " + rs("tab_from_sql")
		filters = ""
		if cString(rs("pub_filtro_pubblicazione"))<>"" then
			filters = SQL_AddOperator(rs("tab_from_sql"), "AND") + rs("pub_filtro_pubblicazione")
		end if
		if cIntero(KeyValue)>0 then
			filters = filters + SQL_AddOperator(rs("tab_from_sql") & filters, "AND") + rs("tab_field_chiave") & "=" & cIntero(KeyValue)
		end if
		rsql = rsql + filters
	else
		rsql = ""
	end if
	if cInteger(PubId) > 0 then
		rs.close
	end if
	
	Pubblicazione_GetQuery = rsql
end function


'calcola i link del contenuto / indice
'	prefisso:		idx, co
'	dest:			dizionario destinazione contenente i campi del link
'	urlSorg:		eventuale dizionario contenente i dati dell'url
'	urlCampoPagina:	eventuale nome del campo contenente l'ID della pagina
'	urlPrefisso:	eventuale prefisso  dei campi contenenti gli url in urlSorg
'	parametro:		eventuale nome del parametro da passare alla pagina
Sub LinkCalculate(conn, prefisso, dest, urlSorg, urlCampoPagina, urlPrefisso, parametro)

'response.write "<h2>LinkCalculate</h2>"
'response.write "prefisso: " & prefisso & "<br>"
'response.write "TypeName(dest): " & TypeName(dest) & "<br>"
'response.write "TypeName(urlSorg): " & TypeName(urlSorg) & "<br>"
'response.write "urlCampoPagina: " & urlCampoPagina & "<br>"
'response.write "urlPrefisso: " & urlPrefisso & "<br>"
'response.write "parametro: " & parametro & "<br>"


	dim rs, sql, recordset
	dim lingua, campo, pagina, FKeyId, isPagina
	pagina = CIntero(urlSorg(urlCampoPagina))
	recordset = InStr(1, TypeName(urlSorg), "recordset", vbTextCompare)
	
	if recordset then		'se sorgente recordset
		if FieldExists(urlSorg, prefisso &"_link_tipo") then
			isPagina = CIntero(urlSorg(prefisso &"_link_tipo")) = lnk_interno
		else
			isPagina = true
		end if
		isPagina = isPagina AND pagina > 0
	else
		isPagina = CIntero(urlSorg(prefisso &"_link_tipo")) = lnk_interno AND pagina > 0
	end if
	
	'recupera codice chiave esterna
	if CString(parametro) <> "" then
		if instr(1, prefisso, "co", vbTextCompare)>0 then
			'e' presente il codice della chiave esterna corretto
			FKeyId = dest("co_F_key_id")
		else
			'recupera codice da codice index
			sql = "SELECT co_F_key_id FROM tb_contents WHERE co_id=" & cIntero(dest("idx_content_id"))
			set rs = conn.Execute(sql)
			FKeyId = rs("co_F_key_id")
		end if
	end if
	
	if isPagina then
		sql = "SELECT * FROM tb_pagineSito WHERE id_pagineSito = "& cIntero(pagina)
		set rs = conn.Execute(sql)
		
		dest(prefisso &"_link_tipo") = lnk_interno
		dest(prefisso &"_link_pagina_id") = pagina
	else
		dest(prefisso &"_link_tipo") = lnk_esterno
        dest(prefisso &"_link_pagina_id") = NULL
	end if
	

'response.write "isPagina: " & isPagina & "<br>"
	
	for each lingua in Application("LINGUE")
		campo = prefisso &"_link_url_"& lingua
		
		if isPagina then		'pagina
			dest(campo) = "?PAGINA="& rs("id_pagDyn_"& lingua)
			
			if CString(parametro) <> "" then
				if InStr(dest(campo), "?") then
					dest(campo) = dest(campo) &"&"& parametro &"="& FKeyId
				else
					dest(campo) = dest(campo) &"?"& parametro &"="& FKeyId
				end if
			end if
		else					'url esterno
			if recordset then
				if FieldExists(urlSorg, urlPrefisso + lingua) then
					dest(campo) = CString(urlSorg(urlPrefisso + lingua))
				end if
			else
				dest(campo) = CString(urlSorg(urlPrefisso + lingua))
			end if
		end if
'response.write "risultato dest(" & campo & ") = " & dest(campo)  &"<br>"
	next

'response.end

End Sub



'richiama la pagina che permette di SPOSTARE gli indirizzi alternativi (e COPIARE quelli principali) da una voce dell'indice ad un'altra
'	IDX:			id della voce dell'indice dalla quale copiare gli URL
'	text:			testo
'	button_text:	testo del pulsante
sub WriteCopiaIndirizziAlternativi(IDX,text,button_text)
	%>
	<table>
		<tr>
			<% if text <> "" then %>
				<td class="note"><%=text%></td>
			<% end if %>
			<td class="content">
				<a class="button_L2" href="javascript:void(0);" onclick="OpenAutoPositionedScrollWindow('<%= GetLibraryPath() %>IndexContent/CopiaIndirizziAlternativi.asp?IDX=<%=IDX%>', 'copy', 500, 300, true);" >
					<%=button_text%>
				</a>
			</td>
		</tr>
	</table>
	<%
end sub


%>