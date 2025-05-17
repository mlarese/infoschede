<%
'.................................................................................................
'.................................................................................................
'.................................................................................................
'FUNZIONI PER LA GESTIONE DEI DATI DEL B2B
'.................................................................................................
'.................................................................................................



class GestioneVariante
    public conn
    public rs
    
    private sql


'DEFINIZIONE COSTRUTTORI: *************************************************************************************************************************************

    Private Sub Class_Initialize()
		set rs = Server.CreateObject("adodb.recordset")
	End Sub
	
	Private Sub Class_Terminate()
        set rs = nothing
	End Sub


'DEFINIZIONE METODI PER INSERIMENTO E MODIFICA DATI: **********************************************************************************************************
    
	public function InsertUpdate(ArtId, RelId, ListaIdValori, _
                                 CodInt, CodPro, CodAlt, _
                                 Prezzo, IsPrezzoIndipendente, PrezzoVarEuro, PrezzoVarSconto, ScontoQId, _
                                 Disabilitato, GiacenzaMin, LottoRiordino, QtaMinOrd)
		InsertUpdate = InsertUpdateComplete(ArtId, RelId, ListaIdValori, _
                                 CodInt, CodPro, CodAlt, _
                                 Prezzo, IsPrezzoIndipendente, PrezzoVarEuro, PrezzoVarSconto, ScontoQId, _
                                 Disabilitato, GiacenzaMin, LottoRiordino, QtaMinOrd, _
								 0, 0, 0, 0, 0, 0, 0, 0)
	end function
		
    public function InsertUpdateComplete(ArtId, RelId, ListaIdValori, _
                                 CodInt, CodPro, CodAlt, _
                                 Prezzo, IsPrezzoIndipendente, PrezzoVarEuro, PrezzoVarSconto, ScontoQId, _
                                 Disabilitato, GiacenzaMin, LottoRiordino, QtaMinOrd, _
								 pesoNetto, pesoLordo, colliNum, colloPezziPer, colloWidth, colloHeight, colloLenght, colloVolume)
        dim rst
        dim IdValore, IsInserting
        set rst = Server.CreateObject("adodb.recordset")
        
        'verifica univocita' codice
        if CodeIsUnique(RelId, CodInt) OR not cBoolean(session("ART_COD_INT_UNIVOCO"), false) then
    
            sql = "SELECT * FROM grel_art_valori WHERE rel_art_id=" & cIntero(ArtId)
            if cIntero(RelId)>0 then
                'variante esistente
                sql = sql & " AND rel_id=" & cIntero(RelId)
            elseif ListaIdValori <> "" then
                'articolo con variante: cerca di identificare variante da valori
                sql = sql & " AND rel_id IN (SELECT rvv_art_var_id FROM grel_art_vv WHERE rvv_val_id IN (" & ListaIdValori & ")) " & _
                            " AND rel_id NOT IN (SELECT rvv_art_var_id FROM grel_art_vv WHERE rvv_val_id NOT IN (" & ListaIdValori & ")) "
            end if
            'inserisce/aggiorna record variante
            rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
    		if rs.eof then
    			rs.addNew       'non trovato: lo inserisce
                IsInserting = true
            else
                IsInserting = false
    		end if
            rs("rel_art_id") = ArtId
            rs("rel_cod_int") = CodInt
    		rs("rel_cod_pro") = CodPro
    		rs("rel_cod_alt") = CodAlt
            if not IsNull(prezzo) OR IsInserting then
        		rs("rel_prezzo") = cReal(Prezzo)
    	    	rs("rel_var_euro") = PrezzoVarEuro
    		    rs("rel_var_sconto") = PrezzoVarSconto
        		rs("rel_prezzo_indipendente") = IsPrezzoIndipendente
            end if
            rs("rel_scontoq_id") = IIF(cIntero(ScontoQId)=0, NULL, ScontoQId)
            rs("rel_disabilitato") = Disabilitato
    		rs("rel_giacenza_min") = cIntero(GiacenzaMin)
    		rs("rel_lotto_riordino") = cIntero(LottoRiordino)
    		rs("rel_qta_min_ord") = cIntero(QtaMinOrd)
            CALL SetUpdateParamsRS(rs, "rel_", IsInserting)
			
			rs("rel_peso_netto") = cReal(pesoNetto)
			rs("rel_peso_lordo") = cReal(pesoLordo)
			rs("rel_colli_num") = cIntero(colliNum)
			rs("rel_collo_pezzi_per") = cIntero(colloPezziPer)
			rs("rel_collo_width") = cReal(colloWidth)
			rs("rel_collo_height") = cReal(colloHeight)
			rs("rel_collo_lenght") = cReal(colloLenght)
			rs("rel_collo_volume") = cReal(colloVolume)
			
            rs.update
            
            RelId = rs("rel_id")
            
            rs.close
            
            'gestione valori variante collegati
            if ListaIdValori <> "" then
				
                If IsInserting then
					CALL InserisciValoriVariante(RelId, ListaIdValori)
                end if
				
				CALL ImpostaOrdineVariante(RelId)
			
			end if
            
            CALL InsertDefaultRows(RelId)
            
            'aggiorna date modifica
            CALL UpdateParams(conn, "gtb_articoli", "art_", "art_id", ArtId, false)
        else
            if ListaIdValori = "" then
                Session("ERRORE") = "Codice dell'articolo non univoco: Esiste gi&agrave; un articolo con codice &quot;" & CodInt & "&quot;."
            else
                Session("ERRORE") = "Nell'inserimento della variante &egrave; stato generato il codice &quot;" & CodInt & "&quot; che risulta gi&agrave; utilzzato. Verificare i codici dell'articolo e relative varianti."
            end if
        end if
    		
        set rst = nothing
        
        InsertUpdateComplete = RelId
    end function
    
    
    '.................................................................................................
    '	funzione inserisce le righe della variante nei listini e nei magazzini
    '	conn:					connessione aperta a database
    '	rel_id                  id della variante
    '.................................................................................................
    public sub InsertDefaultRows(RelId)
        'inserisce righe variante nei listini base
    	sql = " INSERT INTO gtb_prezzi (prz_prezzo, prz_iva_id, prz_var_sconto, prz_var_euro, prz_visibile, prz_promozione, prz_scontoQ_id, prz_listino_id, prz_variante_id )" + _
              " SELECT CASE WHEN ISNULL(listino_default_var_euro,0)<>0 THEN rel_prezzo + listino_default_var_euro " + _
						  " WHEN ISNULL(listino_default_var_sconto,0)<>0 THEN rel_prezzo + ((listino_default_var_sconto/100) * rel_prezzo) " + _
						  " ELSE rel_prezzo " + _
						  " END, " + _
					 " art_iva_id, ISNULL(listino_default_var_sconto,0), ISNULL(listino_default_var_euro,0), "& SQL_If(conn, "listino_offerte=1", "0", "1") &", 0, rel_scontoQ_id, listino_id, rel_id " + _
              "     FROM (grel_art_valori INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id), gtb_listini " + _
              "     WHERE rel_id=" & cIntero(RelId) & " AND (listino_base=1 OR ISNULL(listino_default_var_euro,0)<>0 OR ISNULL(listino_default_var_sconto,0)<>0)" + _
              "           AND rel_id NOT IN (SELECT prz_variante_id FROM gtb_prezzi WHERE prz_listino_id = gtb_listini.listino_id)"
        CALL conn.execute(sql, , adExecuteNoRecords)
    		
    	'inserisce righe variante nella gestione dei magazzini
    	sql = "INSERT INTO grel_giacenze (gia_magazzino_id, gia_art_var_id, gia_qta, gia_impegnato, gia_ordinato, gia_iniziale) " + _
    	      " SELECT mag_id, rel_id, 0, 0, 0, 0 " + _
              " FROM grel_art_valori, gtb_magazzini " + _
              " WHERE rel_id = " & cIntero(RelId) & _
              "       AND rel_id NOT IN (SELECT gia_art_var_id FROM grel_giacenze WHERE gia_magazzino_id = gtb_magazzini.mag_id) "
        CALL conn.execute(sql, , adExecuteNoRecords)
    end sub
    
    
    '.................................................................................................
    '   funzione che aggiona i dati dell'articolo
    '.................................................................................................
    public sub UpdateParamsArticolo(RelId)
        sql = "SELECT rel_art_id FROM grel_art_valori WHERE rel_id=" & cIntero(RelId)
        CALL UpdateParams(conn, "gtb_articoli", "art_", "art_id", GetValueList(conn, rs, sql), false)
    end sub
	
	
    '.................................................................................................
    '   funzione che restituisce l'id della variante dato l'atricolo ed i valori che la compongono
    '.................................................................................................
	public function GetVarianteIdByValori(ArtId, ListaIdValori)
		
		dim sql
		sql = "SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & cIntero(ArtId) & _
			  " AND rel_id IN (SELECT rvv_art_var_id FROM grel_art_vv WHERE rvv_val_id IN (" & ListaIdValori & ")) " & _
			  " AND rel_id NOT IN (SELECT rvv_art_var_id FROM grel_art_vv WHERE rvv_val_id NOT IN (" & ListaIdValori & ")) "
		GetVarianteIdByValori = GetValueList(conn, NULL, sql)
		
	end function

'DEFINIZIONE METODI PER INSERIMENTO E MODIFICA DATI: **********************************************************************************************************

    public sub Delete(RelId)
        
        'esegue operazioni propedeutiche alla cancellazione della variante
        CALL BeforeDelete(RelId)
        
        'cancella variante
        sql = "DELETE FROM grel_art_valori WHERE rel_id=" & cIntero(RelId)
        CALL conn.execute(sql, , adExecuteNoRecords)
        
    end sub
    
    
    public sub BeforeDelete(RelId)
		dim descrizione_IT, descrizione_EN, descrizione_FR, descrizione_DE, descrizione_ES
        
        'disabilta l'articolo padre se questo non ha piu' varianti
        sql = " SELECT COUNT(*) FROM grel_art_valori WHERE rel_art_id IN " + _
		      " (SELECT rel_Art_id FROM grel_art_valori WHERE rel_id=" & cIntero(RelId) & ") " + _
			  " AND rel_id<>" & cIntero(RelId)
        
		if cInteger(GetValueList(conn, rs, sql)) = 0 then	'l'articolo non ha piu' varianti
            sql = " UPDATE gtb_articoli SET art_disabilitato=1 WHERE art_id IN " + _
                  " (SELECT rel_Art_id FROM grel_art_valori WHERE rel_id=" & cIntero(RelId) & ")"
            CALL conn.execute(sql, , adExecuteNoRecords)
        end if
        
        'rimuove le righe in ordine nelle shopping cart
        sql = "DELETE FROM gtb_dett_cart WHERE dett_art_var_id=" & cIntero(RelId)
        CALL conn.execute(sql, , adExecuteNoRecords)
        
        'converte le righe degli ordini in righe di descrizione libera
        sql = "SELECT COUNT(*) FROM gtb_dettagli_ord WHERE det_art_var_id=" & cIntero(RelId)
        if cIntero(GetValueList(conn, rs, sql))>0 then
			CALL GetDescrizioneVariante(RelId, descrizione_IT, descrizione_EN, descrizione_FR, descrizione_DE, descrizione_ES)
			
			sql = " UPDATE gtb_dettagli_ord SET det_art_var_id=NULL, " + _
                  " det_descr_IT = '" & ParseSql(descrizione_IT, adChar) & "', " + _
                  " det_descr_EN = '" & ParseSql(descrizione_EN, adChar) & "', " + _
                  " det_descr_FR = '" & ParseSql(descrizione_FR, adChar) & "', " + _
                  " det_descr_DE = '" & ParseSql(descrizione_DE, adChar) & "', " + _
                  " det_descr_ES = '" & ParseSql(descrizione_ES, adChar) & "' " + _
                  " WHERE det_art_var_id=" & RelId

			CALL conn.execute(sql, , adExecuteNoRecords)
            
        end if
    end sub


'DEFINIZIONE METODI PER GESTIONE DATI: ************************************************************************************************************************
    
    '.................................................................................................
    '	funzione che restituisce l'ordine per la riga di esploso richiesta
    '	grel_art_valori_id:		id della riga di esploso di cui calcolare l'ordine
    '.................................................................................................
    public function GetOrdineVariante(rs, grel_art_valori_id)
    	dim Ordine
    	sql = " SELECT var_ordine, val_ordine FROM grel_art_vv INNER JOIN gtb_valori ON grel_art_vv.rvv_val_id = gtb_valori.val_id " + _
    		  " INNER JOIN gtb_varianti ON gtb_valori.val_var_id = gtb_varianti.var_id " + _
    		  " WHERE grel_art_vv.rvv_art_var_id=" & cIntero(grel_art_valori_id) & _
    		  " ORDER BY var_ordine, val_ordine "
    	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    	Ordine = ""
    	while not rs.eof
    		Ordine = Ordine + String(3 - len(cString(rs("var_ordine"))), "0") + cString(rs("var_ordine")) + _
    				 		  String(3 - len(cString(rs("val_ordine"))), "0") + cString(rs("val_ordine"))
    		rs.movenext
    	wend
    	rs.close
    	GetOrdineVariante = Ordine
    end function
    
    
    
    '.................................................................................................
    'funzione verifica se il codice dell'articolo &egrave; univoco
    '		rel_id 		id della combinazione articolo/varianti
    '		cod			codice da verificare
    '.................................................................................................
    public function CodeIsUnique(rel_id, cod)
    	dim sql
    	sql = " SELECT COUNT(*) FROM grel_art_valori WHERE rel_cod_int LIKE '" & ParseSQL(cod, adChar) & "' "
        if cIntero(rel_id)>0 then
            sql = sql & " AND rel_id<>" & rel_id
        end if
    	CodeIsUnique = (GetValueList(conn, rs, sql)=0)
    end function
    
    
    '.................................................................................................
    '	funzione che restituisce la denominazione completa della variante (nome articolo + valori varianti)
    '	rel_id                  :       id della variante dell'articolo di cui ottenere la descrzione
    '   descrizione_<lingua>    :       valori delle descrizioni ritornate
    '.................................................................................................
    public sub GetDescrizioneVariante(RelId, byref descrizione_IT, byref descrizione_EN, byref descrizione_FR, byref descrizione_DE, byref descrizione_ES)
        
        sql = "SELECT art_nome_it, art_nome_en, art_nome_fr, art_nome_de, art_nome_es, art_varianti FROM gv_articoli WHERE rel_id=" & cIntero(RelId)
        rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
        if not rs.eof then
            descrizione_IT = CBLL(rs, "art_nome", LINGUA_ITALIANO)
            descrizione_EN = CBLL(rs, "art_nome", LINGUA_INGLESE)
            descrizione_FR = CBLL(rs, "art_nome", LINGUA_FRANCESE)
            descrizione_DE = CBLL(rs, "art_nome", LINGUA_TEDESCO)
            descrizione_ES = CBLL(rs, "art_nome", LINGUA_SPAGNOLO)
            
            if rs("art_varianti") then
                rs.close
                
                'recupera lista varianti
                sql = " SELECT * FROM gv_articoli_varianti WHERE rvv_art_var_id=" & cIntero(RelId)
	            rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
                while not rs.eof
                    descrizione_IT = descrizione_IT & IIF(rs.AbsolutePosition=1, " ( ", " - ") & CBLL(rs, "var_nome", LINGUA_ITALIANO) & ": " & CBLL(rs, "val_nome", LINGUA_ITALIANO) & IIF(rs.AbsolutePosition = rs.recordcount, " )", "")
                    descrizione_EN = descrizione_EN & IIF(rs.AbsolutePosition=1, " ( ", " - ") & CBLL(rs, "var_nome", LINGUA_INGLESE) & ": " & CBLL(rs, "val_nome", LINGUA_INGLESE) & IIF(rs.AbsolutePosition = rs.recordcount, " )", "")
                    descrizione_FR = descrizione_FR & IIF(rs.AbsolutePosition=1, " ( ", " - ") & CBLL(rs, "var_nome", LINGUA_FRANCESE) & ": " & CBLL(rs, "val_nome", LINGUA_FRANCESE) & IIF(rs.AbsolutePosition = rs.recordcount, " )", "")
                    descrizione_DE = descrizione_DE & IIF(rs.AbsolutePosition=1, " ( ", " - ") & CBLL(rs, "var_nome", LINGUA_TEDESCO) & ": " & CBLL(rs, "val_nome", LINGUA_TEDESCO) & IIF(rs.AbsolutePosition = rs.recordcount, " )", "")
                    descrizione_ES = descrizione_ES & IIF(rs.AbsolutePosition=1, " ( ", " - ") & CBLL(rs, "var_nome", LINGUA_SPAGNOLO) & ": " & CBLL(rs, "val_nome", LINGUA_SPAGNOLO) & IIF(rs.AbsolutePosition = rs.recordcount, " )", "")
                    
                    rs.movenext
                wend
            end if
        else
            descrizione_IT = ""
            descrizione_EN = ""
            descrizione_FR = ""
            descrizione_DE = ""
            descrizione_ES = ""
        end if
        rs.close
        
    end sub
    
    

    'funzione che recupera l'id del valore variante, se non c'e' lo inserisce.
    function GetIdValoreVariante(VarianteId, Valore)
        sql = "SELECT * FROM gtb_valori WHERE val_var_id=" & cIntero(VarianteId) & " AND val_nome_it LIKE '" & ParseSql(Valore, adChar) & "'"
        rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
        if rs.eof then
            rs.AddNew
            rs("val_var_id") = VarianteId
            rs("val_nome_it") = Valore
            rs("val_cod_int") = Valore
            rs.update
        end if
    
        GetIdValoreVariante = rs("val_id")
    
        rs.close
    end function
	
    
	'inserisce i valori della variante indicata
    sub InserisciValoriVariante(relId, listaIdValori)
		dim idValore, sql
		
		for each idValore in split(replace(listaIdValori, " ", ""), ",")
			sql = "INSERT INTO grel_art_vv (rvv_art_var_id, rvv_val_id) " + _
				  "VALUES(" & relId & ", " & idValore & ")"
			CALL conn.execute(sql, , adExecuteNoRecords)
        next
		
	end sub
	
	
	'ricalcola ed imposta l'ordine della variante indicata
    sub ImpostaOrdineVariante(relId)
	
		dim sql
		sql = "UPDATE grel_art_valori SET rel_ordine = '" & GetOrdineVariante(rs, relId) & "' WHERE rel_id=" & cIntero(relId)
		CALL conn.execute(sql)
	
    end sub


end class
%>