<%
'.................................................................................................................
'
'  TOOLS PER IMPORT DATI DA PORTALI APT
'
'.................................................................................................................



'****************************************************************************************************
'COSTANTI PER SYNCRO ED IMPORT DA APT
'****************************************************************************************************

'.................................................................................................................
'Definizione delle costanti per i codici di collegamento

'radici dei codici per ogni APT e per l'assessorato
const CODE_ASSESSORATO = "ASS00"
const CODE_APTBC = "APT04"
const CODE_APTJE = "APT05"
const CODE_APTVE = "APT06"
const CODE_APTCH = "APT07"

'radici dei codici per ogni entita'
const CODE_LUOGHI = "L"
const CODE_SPIAGGE = "S"
const CODE_NOTIZIE = "N"        'notizie utili e relativi sottotipi
const CODE_NOTIZIE_T = "U"      'tipi notizie utili
const CODE_LOCALI = "T"         'locali e servizi e relativi sottotipi
const CODE_LOCALI_T = "D"       'tipi di locali e servizi
const CODE_RICETTIVITA = "R"    'ricettivita'
const CODE_ASSOCIAZIONI = "C"   'asssociazioni di categoria dell'assessorato
const CODE_ZONE = "Z"
const CODE_SUBZONE = "S"
const CODE_EVENTI = ""


'*****************************************************************************************************************
'DEFINIZIONE COSTANTI E FUNZIONI COMUNI
'*****************************************************************************************************************



'.................................................................................................................
'funzione che ritorna la parte di filtro da aggiungere alla query per ricercare la categoria con codice indicato
'nei codici delle categorie o nella lista dei codici provenienti dalle APT.
'oCat                       oggetto categorie da filtrare
'Codice                     codice o porzione del codice da utilizzare nel filtro
'PerCategoriaPrincipale     cerca tra i dati di collegamento con le categorie APT
'PerCategoriaAlternativa    cerca nei codici delle categorie importate direttamente dai portali delle APT
'.................................................................................................................
function import_QueryCategorieByCodice(oCat, Codice, PerCategoriaPrincipale, PerCategoriaAlternativa)
    dim sql
    if PerCategoriaPrincipale OR PerCategoriaAlternativa then
        sql = " ( "
        if PerCategoriaPrincipale then
            'ricerca le categorie del portale collegate alla categoria di origine dell'APT
            sql = sql + oCat.prefisso + "_ListaCategorieApt_EXT LIKE '%" & Codice & "%' OR "
        end if
        if PerCategoriaAlternativa then
            'rierca le categorie che arrivano direttamente dalle APT
            sql = sql + oCat.prefisso + "_codice LIKE '" & Codice & IIF(len(Codice)<7, "%", "") & "' OR "
        end if
        import_QueryCategorieByCodice = left(sql, len(sql)-3) & ") "
    else
        import_QueryCategorieByCodice = ""
    end if
end function


'.................................................................................................................
'funzione che recupera la categoria pricipale collegata (tramite campi esterni di collegamento con APT) alla
'categoria di codice indicato
'oCat           oggetto categorie da cui recuperare la categoria interessata
'codice:        codice della categoria da ricercare nel campo lista di collegamento.
'.................................................................................................................
function import_GetCategoriaPrincipale(oCat, conn, rs, codice)
    dim sql
    sql = "SELECT " + oCat.prefisso + "_id FROM " + oCat.Tabella + " WHERE " + import_QueryCategorieByCodice(oCat, codice, true, false)
    import_GetCategoriaPrincipale = GetValueList(conn, rs, sql)
end function


'.................................................................................................................
'funzione che recupera la categoria alternativa (importate da APT) con codice uguale a quello richiesto
'codice:        codice della categoria da ricercare
'.................................................................................................................
function import_GetCategoriaAlternativa(oCat, conn, rs, codice)
    dim sql
    sql = "SELECT " + oCat.prefisso + "_id FROM " + oCat.Tabella + " WHERE " + import_QueryCategorieByCodice(oCat, codice, false, true)
    import_GetCategoriaAlternativa = GetValueList(conn, rs, sql)
end function


'.................................................................................................................
'funzione che ritorna l'area corrispondente alla ricerca indicata
'AptCode        codice dell'apt di competenza
'ZonaId         id della zona di cui cercare l'area di competenza
'SubZonaId      id della subzona di competenza
'.................................................................................................................
function import_GetAptArea(AptCode, ZonaId, SubZonaId)
    dim sql, Area
    ZonaId = cInteger(ZonaId)
    SubZonaId = cInteger(SubZonaId)
    Area = 0
    
    if SubZonaId>0 then
        Area = import_GetCategoriaByCode(iAree, GetCodice(AptCode, CODE_SUBZONE, SubZonaId))
    end if
    
    if Area=0 then
        Area = import_GetCategoriaByCode(iAree, GetCodice(AptCode, CODE_ZONE, ZonaId))
    end if
    
	if Area=0 then
        Area = import_GetCategoriaByCode(iAree, GetCodice(AptCode, "", ""))
	end if
	
    if Area = 0 then
        response.write "ERRORE: AREA NON TROVATA<br>"
        response.end
    else
        import_GetAptArea = Area
    end if
end function


'.................................................................................................................
'funzione che inizializza oggetti e recupera configurazione per inserimento dati delle anagrafiche
'.................................................................................................................
sub import_CaricaConfigurazione(conn, rs, byref NEXTAIM_ADMIN_ID, byref RUBRICA_ANAGRAFICHE, byref oContatto)
	dim sql
	sql = "SELECT id_Admin FROM tb_admin WHERE admin_login LIKE 'NEXTAIM'"
	NEXTAIM_ADMIN_ID = GetValueList(conn, rs, sql)
	
	sql = "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'RUBRICA_ANAGRAFICHE' AND par_sito_id=" & NEXTINFO
	RUBRICA_ANAGRAFICHE = GetValueList(conn, rs, sql)
	
	'se non c'e la crea ed associa tutte le anagrafiche corrispondenti
	sql = " SELECT * FROM tb_rubriche WHERE id_Rubrica=" & cInteger(RUBRICA_ANAGRAFICHE)
	rs.open sql, conn, adOpenKeySet, adLockOptimistic
	if rs.eof then
		rs.addnew
		rs("nome_rubrica") = "Anagrafiche - elenco completo"
		rs("locked_rubrica") = true
		rs("rubrica_esterna") = true
		rs("SyncroTable") = "itb_anagrafiche"
		rs.update
		
		RUBRICA_ANAGRAFICHE = rs("id_rubrica")
		rs.close
		
		sql = "SELECT * FROM tb_siti_parametri WHERE par_key LIKE 'RUBRICA_ANAGRAFICHE' AND par_sito_id=" & NEXTINFO
		rs.open sql, conn, adOpenKeySet, adLockOptimistic
		if rs.eof then
			rs.AddNew
			rs("par_key") = "RUBRICA_ANAGRAFICHE"
			rs("par_sito_id") = NEXTINFO
		end if
		rs("par_value") = cString(RUBRICA_ANAGRAFICHE)
		rs.update
	else
		RUBRICA_ANAGRAFICHE = rs("id_rubrica")
	end if
	rs.close
	
	set oContatto = new IndirizzarioLock
	set oContatto.conn = conn
end sub


'.................................................................................................................
'funzione che inserisce una categoria date le informazioni indicate
'conn:              connessione al database aperta e sotto transazione
'rs:                oggetto recordset usato
'oCat:              oggetto categorie creato ed impostato per descrivere l'albero di categorizzazione in cui inserire il record
'CategoriaPadre:    Categoria padre della categoria da inserire
'codice;            codifica secondo costanti del codice del record proveniente da portali APT
'nome_..            nome in lingua della categoria
'.................................................................................................................
function import_CATEGORIA(conn, rs, oCat, CategoriaPadre, codice, nome_it, nome_en, nome_fr, nome_de, nome_es, ordine)
    import_CATEGORIA = import_syncro_CATEGORIA(conn, rs, oCat, CategoriaPadre, codice, NULL, NULL, nome_it, nome_en, nome_fr, nome_de, nome_es, ordine)
end function

function import_syncro_CATEGORIA(conn, rs, oCat, CategoriaPadre, codice, ExternalSource, ExternalId, nome_it, nome_en, nome_fr, nome_de, nome_es, ordine)
    dim sql
    
    'recupera record categoria se recordset chiuso
    if rs.state = adStateClosed then
        sql = " SELECT *  FROM " + oCat.Tabella
        if IsNull(ExternalSource) then
            'ricerca la categoria in base al codice
            sql = sql + " WHERE " + oCat.prefisso + "_codice LIKE '" & codice & "' " & _
                        " AND " + oCat.prefisso + "_padre_id=" & CategoriaPadre
        else
            'ricerca la categoria in base alle categorie sorgenti
            sql = sql + " WHERE " + oCat.prefisso + "_external_id LIKE '%" & IdList_ID(ExternalId) & "%' " & _
                        " AND " + oCat.prefisso + "_external_source LIKE '" & ExternalSource & "' "
        end if
        rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    end if
    if rs.eof then
        'se non c'e' la inserisce
		rs.AddNew
	end if
    if not IsNull(ExternalSource) then
        rs(oCat.prefisso + "_external_id") = IdList_ADD(rs(oCat.prefisso + "_external_id"), ExternalId)
	    rs(oCat.prefisso + "_external_source") = ExternalSource
    end if
    rs(oCat.prefisso + "_codice") = codice
	rs(oCat.prefisso + "_padre_id") = CategoriaPadre
	rs(oCat.prefisso + "_nome_it") = nome_it
	rs(oCat.prefisso + "_nome_en") = nome_en
	rs(oCat.prefisso + "_nome_fr") = nome_fr
	rs(oCat.prefisso + "_nome_de") = nome_de
	rs(oCat.prefisso + "_nome_es") = nome_es
	rs(oCat.prefisso + "_visibile") = true
    rs(oCat.prefisso + "_ordine") = ordine
    rs.Update
    sql = rs.source
	rs.close
	
    rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    import_syncro_CATEGORIA = cInteger(rs(oCat.prefisso + "_id"))
    rs.close
	
	'esegue operazioni di gestione categoria
	oCat.operazioni_ricorsive_tipologia(CategoriaPadre)
end function


'.................................................................................................................
'funzione che restituisce l'id della categoria con codice indicato
'.................................................................................................................
function import_GetCategoriaByCode(oCat, CodiceRecord)
    dim sql
    sql = "SELECT * FROM " & oCat.tabella & " WHERE " & oCat.prefisso & "_codice LIKE '" & CodiceRecord & "'"
    import_GetCategoriaByCode = cInteger(GetValueList(oCat.conn, NULL, sql))
end function


'.................................................................................................................
'funzione che ritorna il gruppo dei descrittori delle anagrafiche indicato da Sorgente e chiave - se non c'e' lo inserisce
'.................................................................................................................
function import_GruppoDescrittoreAnagrafiche(conn, rs, ExternalSource, ExternalId, Nome_it, Nome_en, Nome_fr, Nome_de, Nome_es, Ordine)
	dim sql
	
	sql = " SELECT * FROM itb_anagrafiche_descrRag " + _
		  " WHERE adr_external_id LIKE '%(" & ExternalId & ")%' AND adr_external_source LIKE '" & ExternalSource & "' "
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	if rs.eof then
		rs.AddNew
	end if
	rs("adr_external_id") = IdList_ADD(rs("adr_external_id"), ExternalId)
	rs("adr_external_source") = ExternalSource
	rs("adr_titolo_it") = nome_it
	rs("adr_titolo_en") = nome_en
	rs("adr_titolo_fr") = nome_fr
	rs("adr_titolo_de") = nome_de
	rs("adr_titolo_es") = nome_es
	rs("adr_ordine") = ordine
	rs.update
    rs.close
    
    rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    import_GruppoDescrittoreAnagrafiche = cInteger(rs("adr_id"))
    rs.close
end function


'.................................................................................................................
'funzione che ritorna lid del descrittore indicato da Sorgente e Chiave, se non c'e' lo inserisce
'ex funzione syncro_descrittore anagrafiche
'.................................................................................................................
function import_DescrittoreAnagrafiche(conn, rs, ExternalSource, ExternalId, Nome_it, Nome_en, Nome_fr, Nome_de, Nome_es, Unita, Tipo, Immagine, Gruppo_id)
	dim sql
	
	sql = " SELECT * FROM itb_anagrafiche_descrittori " + _
		  " WHERE (and_external_id LIKE '%(" & ExternalId & ")%' OR and_nome_it LIKE '" & ParseSQL(Nome_it, adChar) & "') " + _
		  " AND and_external_source LIKE '" & ExternalSource & "' "
	
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	if rs.eof then
		rs.AddNew
	end if
	rs("and_external_id") = IdList_ADD(rs("and_external_id"), ExternalId)
	rs("and_external_source") = ExternalSource
	rs("and_raggruppamento_id") = Gruppo_id
	rs("and_nome_it") = Nome_it
	rs("and_nome_en") = Nome_en
	rs("and_nome_fr") = Nome_fr
	rs("and_nome_de") = Nome_de
	rs("and_nome_es") = Nome_es
	rs("and_unita_it") = Unita
	rs("and_principale") = false
	rs("and_tipo") = Tipo
	rs("and_img") = Immagine
	rs.update
	rs.close
	
    rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    import_DescrittoreAnagrafiche = cInteger(rs("and_id"))
    rs.close
end function 


'.................................................................................................................
'procedura inserisce l'associazione tra le categorie selezionate nell'elenco ed i descrittori nella lista (Separati da ",").
'.................................................................................................................
sub import_Associa_DescrittoriCategorieAnagrafiche(conn, rs, SqlElencoCategorie, ListaDescrittori, locked)
    dim rsA, ArrayDescrittori, Descrittore
    
    set rsA = Server.CreateObject("ADODB.Recordset")
    ArrayDescrittori = split(replace(ListaDescrittori, " ", ""), ",")
    
    'apre elenco delle categorie da associare
    rs.open SqlElencoCategorie, conn, adOpenStatic, adLockOptimistic, adCmdtext
    while not rs.eof
        for each Descrittore in ArrayDescrittori
            CALL Syncro_AssociazioneDescrittoriCategorieAnagrafiche(conn, rsA, rs("ant_id"), cInteger(Descrittore), "", locked)
        next
        rs.movenext
    wend
    rs.close
    
    set rsA = nothing
end sub


'.................................................................................................................
'procedura che importa le foto relative al record della tabella specificata nelle immagini anagrafiche
'.................................................................................................................
sub import_ImmaginiAnagrafiche(Sconn, Dconn, Srs, Drs, AptCode, Table, KeyField, KeyValue, AnaId)
    dim rs
    set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM " & Table & " WHERE " & KeyField & " = " & KeyValue
	Srs.open sql, Sconn, adOpenStatic, adLockOptimistic, adCmdText
	
	CALL import_ClearImmaginiAnagrafica(conn, AnaId) 
		
	sql = "SELECT * FROM irel_anagrafiche_img WHERE ani_anagrafica_id=" & AnaId
	Drs.open sql, Dconn, adOpenStatic, adLockOptimistic, adCmdText
	
	while not Srs.eof
        CALL import_ImmagineAnagrafica(conn, rs, AptCode, AnaId, _
                                       "thumb/" & Srs("link"), _
                                       "prev/" & Srs("link"), _
                                       TextDecode(Srs("Descriz")), TextDecode(Srs("Desc_eng")), "", "", "")
        Srs.movenext
	wend
	
	Drs.close
	Srs.close
    
    set rs = nothing
end sub


'.................................................................................................................
'procedura ripulisce le immagini dell'anagrafica indicata
'.................................................................................................................
sub import_ClearImmaginiAnagrafica(conn, AnaId) 
    sql = "DELETE FROM irel_anagrafiche_img WHERE ani_anagrafica_id=" & AnaId
	CALL conn.execute(sql)
end sub


'.................................................................................................................
'procedura che inserisce una immagine
'.................................................................................................................
sub import_ImmagineAnagrafica(conn, rs, AptCode, AnaId, Thumb, Zoom, DidascaliaIt, DidascaliaEn, DidascaliaFr, DidascaliaDe, DidascaliaEs)
    dim sql, AutoOpened
    
    if cString(Thumb)<>"" OR cString(Zoom)<>"" then
        AutoOpened = rs.state = adStateClosed
        if AutoOpened then
            sql = "SELECT * FROM irel_anagrafiche_img WHERE ani_anagrafica_id=" & AnaId
    	    rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    	end if
        
    	rs.AddNew
    	rs("ani_anagrafica_id") = AnaId
    	rs("ani_visibile") = true
    	rs("ani_thumb") = AptCode & "/" & Thumb
        rs("ani_zoom") = AptCode & "/" & Zoom
        rs("ani_didascalia_it") = DidascaliaIt
    	rs("ani_didascalia_en") = DidascaliaEn
        rs("ani_didascalia_fr") = DidascaliaFr
    	rs("ani_didascalia_de") = DidascaliaDe
        rs("ani_didascalia_es") = DidascaliaEs
    	rs.Update
    	
        if AutoOpened then
            rs.close
        end if
    end if
    
end sub


'*****************************************************************************************************************
'FUNZIONI PER IMPORT DELLE CATEGORIE
'*****************************************************************************************************************


'.................................................................................................................
'funzione che importa le categorie dal portale apt indicato
'.................................................................................................................
sub import_CATEGORIE_APT(oCategorie, AptCode, CodiceRecord, CodicePadre, AptTabella, AptId, AptNomeIt, AptNomeEn, AptNomeFr, AptNomeDe, AptNomeEs)
    dim connApt, rs, Apt_rs
	dim sql, CategoriaBase
    dim Codice, DaImportare
	
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
	
	set rs = Server.CreateObject("ADODB.Recordset")
	set Apt_rs = Server.CreateObject("ADODB.Recordset")

    'recupera categoria base
    CategoriaBase = import_GetCategoriaByCode(oCategorie, GetCodice(AptCode, CodicePadre, ""))
    
    if CategoriaBase>0 then
        'recupera elenco categorie eventi dal database dell'apt
        sql = "SELECT * FROM " & AptTabella
        Apt_rs.open sql, connApt, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        while not Apt_rs.eof 
            Codice = GetCodice(AptCode, CodiceRecord, Apt_rs(AptId))
            
            if instr(1, oCategorie.tabella, "itb_aree", vbTextCompare)>0 then
                DaImportare = true
            else
                DaImportare = cInteger(import_GetCategoriaPrincipale(oCategorie, conn, rs, Codice))>0
            end if %>
            <% if DaImportare then
                CALL import_CATEGORIA(conn, rs, oCategorie, CategoriaBase, _
                                      Codice, _
                                      Apt_rs(AptNomeIt), _
                                      Apt_rs(AptNomeEn), _
                                      Apt_rs(AptNomeFr), _
                                      Apt_rs(AptNomeDe), _
                                      Apt_rs(AptNomeEs), _
                                      Apt_rs(AptId))
            else %>
                <tr>
                    <th class="warning"><%= Codice %></td>
                    <th class="warning" colspan="2"><%= Apt_rs(AptNomeIt) %></td>
                </tr>
            <%end if
		    Apt_rs.movenext
    	wend
        
        Apt_rs.close
	else %>
        <tr><th class="alert" colspan="3">CATEGORIA DI <%= CodicePadre %> PER <%= AptCode %> NON TROVATA!!!!</td></tr>
        <% response.end
    end if
	
	connApt.close
	
	set connApt = nothing
	set rs = nothing
	set Apt_rs = nothing
end sub


'.................................................................................................................
'funzione che importa le sottocategorie (Notizie utili e Locali e servizi) dal portale apt indicato
'.................................................................................................................
sub import_SOTTO_CATEGORIE_APT(oCategorie, AptCode, CodiceRecord, CodiceRecordPadre, AptTabella, AptId, AptPadreId, AptNomeIt, AptNomeEn, AptNomeFr, AptNomeDe, AptNomeEs)
     dim connApt, rs, Apt_rs
	dim sql, CategoriaPadre
    dim Codice, CodicePadre, DaImportare
	
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
	
	set rs = Server.CreateObject("ADODB.Recordset")
	set Apt_rs = Server.CreateObject("ADODB.Recordset")
    
    sql = "SELECT * FROM " & AptTabella
    Apt_rs.open sql, connApt, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    while not Apt_rs.eof 
        
        CodicePadre = GetCodice(AptCode, CodiceRecordPadre, Apt_rs(AptPadreId))
        CategoriaPadre = import_GetCategoriaByCode(oCategorie, CodicePadre)
        
        Codice = GetCodice(AptCode, CodiceRecord, Apt_rs(AptId))
        DaImportare = cInteger(import_GetCategoriaPrincipale(oCategorie, conn, rs, Codice))>0
        
        if DaImportare AND cInteger(CategoriaPadre)>0 then
            CALL import_CATEGORIA(conn, rs, oCategorie, CategoriaPadre, _
                                      Codice, _
                                      Apt_rs(AptNomeIt), _
                                      Apt_rs(AptNomeEn), _
                                      Apt_rs(AptNomeFr), _
                                      Apt_rs(AptNomeDe), _
                                      Apt_rs(AptNomeEs), _
                                      Apt_rs(AptId))
        elseif DaImportare then%>
            <tr><th class="alert" colspan="3">CATEGORIA DI <%= GetCodice(AptCode, CodiceRecordPadre, Apt_rs(AptPadreId)) %> PER <%= AptCode %> NON TROVATA!!!!</td></tr>
            <% response.end
        else %> 
             <tr>
                <th class="warning"><%= Codice %></td>
                <th class="warning" colspan="2"><%= Apt_rs(AptNomeIt) %></td>
            </tr>
        <% end if
        Apt_rs.movenext
    wend
        
    Apt_rs.close
	
    connApt.close
	
	set connApt = nothing
	set rs = nothing
	set Apt_rs = nothing
end sub


'.................................................................................................................
'funzione che importa le aree dalle zone e subzone dei portali APT
'.................................................................................................................
sub import_AREE(oCat, AptCode)
    dim connApt, rs, Apt_rsZ, Apt_rsS
	dim sql, AreaPadreID, AreaZonaId, AreaSubzonaId
	
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
	
	set rs = Server.CreateObject("ADODB.Recordset")
	set Apt_rsZ = Server.CreateObject("ADODB.Recordset")
	set Apt_rsS = Server.CreateObject("ADODB.Recordset")

	response.write "<!-- " & AptCode & " -->" & vbCrLf
    
    'recupera area riferita ad Ambito
    AreaPadreID = import_GetCategoriaByCode(oCat, GetCodice(AptCode, "", ""))
    
    if AreaPadreID>0 then
        'import delle zone
    	sql = "SELECT * FROM Zone_StruRic"
	    Apt_rsZ.open sql, connApt, adOpenForwardOnly, adLockReadOnly, adCmdText
	    
        while not Apt_rsZ.eof
            'importa zona
            AreaZonaId = import_CATEGORIA(conn, rs, oCat, AreaPadreID, _
                                          GetCodice(AptCode, CODE_ZONE, Apt_rsZ("id_zonaRic")), _
                                          Apt_rsZ("zona_nome_it"), _
                                          Apt_rsZ("zona_nome_eng"), _
                                          "", _
                                          "", _
                                          "", _
                                          Apt_rsZ("id_zonaRic"))
            
            'imposta relazioni tra zone e comuni-localita'
            CALL import_AREE_ImpostaRelazioni(conn, AreaZonaId, Apt_rsZ("Com_Provincia"), Apt_rsZ("Loc_Provincia"))
            
            'importa subzone
            response.write "<!-- SUBZONE -->" + vbCrLf
		    sql = "SELECT * FROM SubZone WHERE rif_zona=" & Apt_rsZ("id_zonaRic") 
            Apt_rsS.open sql, connApt, adOpenForwardOnly, adLockReadOnly, adCmdText
		    
            while not Apt_rsS.eof
                AreaSubzonaId = import_CATEGORIA(conn, rs, oCat, AreaZonaId, _
                                                 GetCodice(AptCode, CODE_SUBZONE, Apt_rsS("id_subzona")), _
                                                 Apt_rsS("subzona_nome"), _
                                                 "", _
                                                 "", _
                                                 "", _
                                                 "",_
                                                 Apt_rsS("id_subzona"))
                'imposta relazioni tra zone e comuni-localita'
                CALL import_AREE_ImpostaRelazioni(conn, AreaSubzonaId, Apt_rsS("Com_Provincia"), Apt_rsS("Loc_Provincia"))
			    
                Apt_rss.movenext
		    wend
		
		    Apt_rsS.close
		
            Apt_rsZ.movenext
	    wend
        
        Apt_rsZ.close
    else
        response.write "AREA " & AptCode & " NON TROVATA <br>"
        response.end
    end if
    
	connApt.close
	
	set connApt = nothing
	set rs = nothing
	set Apt_rsZ = nothing
	set Apt_rsS = nothing
end sub


'.................................................................................................................
'funzione che importa le relazioni tra le aree ed i comuni e le localita'
'.................................................................................................................
sub import_AREE_ImpostaRelazioni(conn, are_id, Comuni, Localita)
	dim sql, Xsql
	dim A_Comuni, Comune
	dim A_Localita, A2_Localita, Loc, LocId
	
	'ripulisce vecchi collegamenti con comuni e localita
	sql = "DELETE FROM irel_aree_comuni WHERE rac_are_id=" & are_id & vbCrLf + _
		  "DELETE FROM irel_aree_localita WHERE ral_are_id=" & are_id & vbCrLf
	CALL conn.execute(sql, , adExecuteNoRecords)
	
	if Comuni<>"" then
		response.write "<!-- Collegamento con COMUNI:" & comuni & "-->" + vbCrLf
		'collegamento con i comuni dell'area
		Xsql = ""
		A_Comuni = split(Comuni, ",")
		for each Comune in A_Comuni
			'verifica presenza comune
			sql = "SELECT COUNT(*) FROM tb_comuni WHERE codice_istat='" + Trim(comune) + "'"
			if cInteger(GetValueList(Conn, NULL, sql))>0 then
				Xsql = Xsql + " INSERT INTO irel_aree_comuni (rac_are_id, rac_comune_codice) " + _
							  " VALUES (" & are_id & ", '" + Trim(comune) + "' ) " + vbcRLF
			end if
		next
		if Xsql<>"" then
			CALL conn.execute(Xsql, , adExecuteNoRecords)
		end if
	end if
	
	if Localita<>"" then
		response.write "<!-- Collegamento con LOCALIT&Agrave;:" & Localita & " -->" + vbCRLf
		'collegamento con le localita dell'area
		Xsql = ""
		A_localita = split(Localita, ",")
		for each Loc IN A_Localita
			A2_Localita = split(Trim(Loc), "-")
			if uBound(A2_Localita)=1 then
				'verifica presenza localita
				sql = "SELECT loc_id FROM tb_localita WHERE loc_comune='" + Trim(A2_Localita(0)) + "' AND loc_cod='" + Trim(A2_Localita(1)) + "' "
				LocId = cInteger(GetValueList(Conn, NULL, sql))
				if LocId>0 then
					Xsql = Xsql + " INSERT INTO irel_aree_localita (ral_are_id, ral_loc_id) " + _
								  " VALUES (" & are_id & ", " & LocId & " ) " + vbcRLF
				end if
			end if
		next
		if Xsql<>"" then
			CALL conn.execute(Xsql, , adExecuteNoRecords)
		end if
	end if
end sub



'*****************************************************************************************************************
'FUNZIONI PER IMPOSTAZIONE DEI DESCRITTORI
'*****************************************************************************************************************


'.................................................................................................................
'funzione che importa i descrittori delle mappe e li associa con tutte le categorie
'.................................................................................................................
sub import_DESCRITTORI_ANAGRAFICHE__MAPPE() 
	dim GruppoId, DesId_list
    dim rs, sql
    set rs = Server.CreateObject("ADODB.Recordset")
    
	GruppoId = import_GruppoDescrittoreAnagrafiche(conn, rs, "general", "MappeApt", "Mappa da portali APT", "Map from the site of Tourist Board", "", "", "", 100)
    
    'gestione delle mappe
	DesId_list = import_DescrittoreAnagrafiche(conn, rs, "general", "collocazione", _
									           "Collocazione mappa", "", "", "", "", _
                                               "", adVarChar, "", GruppoId)
                                               
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "linkmappe", _
                                                                  "Url della mappa", "", "", "", "", _
                                                                  "", adUserDefined, "", GruppoId)
    
    'associa i descrittori delle mappe a tutte le categorie
    sql = "SELECT * FROM itb_anagrafiche_tipi WHERE ant_livello>0"
    CALL import_Associa_DescrittoriCategorieAnagrafiche(conn, rs, sql, DesId_list, false)
    
    set rs = nothing
end sub


'.................................................................................................................
'funzione che gestisce ed inserisce i valori dei descrittori delle mappe.
'.................................................................................................................
sub import_ValoriMappeAnagrafica(conn, rs, CntId, AptCode, Collocazione, PaginaMappe)
    dim URL
    
    'imposta la collocazione
    CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "collocazione", Collocazione, "", "", "", "")
    
    'url della mappa
    if cInteger(PaginaMappe)>0 then
        URL = "http://"
        Select case AptCode
            case CODE_APTBC
                URL = URL + "www.caorleapt.it"
            case CODE_APTJE
                URL = URL + "www.turismojesoloeraclea.it"
            case CODE_APTVE
                URL = URL + "www.turismovenezia.it"
            case CODE_APTCH
                URL = URL + "www.chioggiatourism.it"
        end select
        URL = URL + "/default.asp?PS=" & PaginaMappe & "&lingua="
        
        CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "linkmappe", URL & "IT", URL & "EN", URL & "FR", URL & "DE", URL & "ES")
        
    end if
    
end sub


'.................................................................................................................
'funzione che importa i descrittori delle spiagge
'.................................................................................................................
sub import_DESCRITTORI_ANAGRAFICHE__SPIAGGE() 
    dim GruppoId, DesId_list
    dim rs, sql
    
    set rs = Server.CreateObject("ADODB.Recordset")
    
    'prepara struttura gruppi e descrittori
	GruppoId = import_GruppoDescrittoreAnagrafiche(conn, rs, "general", "prezzi", "Prezzi", "Prices", "", "", "", 10)
	
	DesId_list = import_DescrittoreAnagrafiche(conn, rs, "general", "apertura", _
                                               "Periodo di apertura", "Opening period", "", "", "", _
                                               "", adVarChar, "", NULL)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "orario", _
                                                                  "Orari di apertura", "Opening times", "", "", "", _
                                                                  "", adVarChar, "", NULL)
	DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "ombr_giorn_", _
                                                                  "Ombrelloni - giornaliero", "Beach umbrellas - daily cost", "", "", "", _
                                                                  "&euro;", adDouble, "ombrll.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "ombr_mens_", _
										                          "Ombrelloni - mensile", "Beach umbrellas - monthly cost", "", "", "", _
                                                                  "&euro;", adDouble, "ombrll.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "ombr_stag_", _
                        										  "Ombrelloni - stagionale", "Beach umbrellas - seasonal cost", "", "", "", _
                                                                  "&euro;", adDouble, "ombrll.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "lett_giorn_", _
                                                                  "Lettini - giornaliero", "Deckchair - daily cost", "", "", "", _
                                                                  "&euro;", adDouble, "lettini.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "lett_mens_", _
                                                                  "Lettini - mensile", "Deckchair - monthly cost", "", "", "", _
                                                                  "&euro;", adDouble, "lettini.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "lett_stag_", _
                                                                  "Lettini - stagionale", "Deckchair - seasonal cost", "", "", "", _
                                                                  "&euro;", adDouble, "lettini.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "spo_giorn_", _
                                                                  "Spogliatoi - giornaliero", "Dressing rooms - daily cost", "", "", "", _
                                                                  "&euro;", adDouble, "spogl.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "spo_mens_", _
                                                                  "Spogliatoi - mensile", "Dressing rooms - monthly cost", "", "", "", _
                                                                  "&euro;", adDouble, "spogl.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "spo_stag_", _
                                                                  "Spogliatoi - stagionale", "Dressing rooms - seasonal cost", "", "", "", _
                                                                  "&euro;", adDouble, "spogl.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "camer_giorn_", _
                                                                  "Camerini - giornaliero", "Small rooms - daily cost", "", "", "", _
                                                                  "&euro;", adDouble, "camerin.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "camer_mens_", _
                                                                  "Camerini - mensile", "Small rooms - monthly cost", "", "", "", _
                                                                  "&euro;", adDouble, "camerin.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "camer_stag_", _
                                                                  "Camerini - stagionale", "Small rooms - seasonal cost", "", "", "", _
                                                                  "&euro;", adDouble, "camerin.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "minicap_giorn_", _
                                                                  "Minicapanne - giornaliero", "Minihuts - daily cost", "", "", "", _
                                                                  "&euro;", adDouble, "minicap.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "minicap_mens_", _
                                                                  "Minicapanne - mensile", "Minihuts - monthly cost", "", "", "", _
                                                                  "&euro;", adDouble, "minicap.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "minicap_stag_", _
                                                                  "Minicapanne - stagionale", "Minihuts - seasonal cost", "", "", "", _
                                                                  "&euro;", adDouble, "minicap.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "capf1_giorn_", _
                                                                  "Capanne 1^ fila - giornaliero", "Huts 1^ row - daily cost", "", "", "", _
                                                                  "&euro;", adDouble, "capann.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "capf1_mens_", _
                                                                  "Capanne 1^ fila - mensile", "Huts 1^ row - monthly cost", "", "", "", _
                                                                  "&euro;", adDouble, "capann.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "capf1_stag_", _
                                                                  "Capanne 1^ fila - stagionale", "Huts 1^ row - seasonal cost", "", "", "", _
                                                                  "&euro;", adDouble, "capann.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "capf2_giorn_", _
                                                                  "Capanne 2^ fila - giornaliero", "Huts 2^ row - daily cost", "", "", "", _
                                                                  "&euro;", adDouble, "capann.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "capf2_mens_", _
                                                                  "Capanne 2^ fila - mensile", "Huts 2^ row - monthly cost", "", "", "", _
                                                                  "&euro;", adDouble, "capann.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "capf2_stag_", _
                                                                  "Capanne 2^ fila - stagionale", "Huts 2^ row - seasonal cost", "", "", "", _
                                                                  "&euro;", adDouble, "capann.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "capf3_giorn_", _
                                                                  "Capanne 3^ fila - giornaliero", "Huts 3^ row - daily cost", "", "", "", _
                                                                  "&euro;", adDouble, "capann.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "capf3_mens_", _
                                                                  "Capanne 3^ fila - mensile", "Huts 3^ row - monthly cost", "", "", "", _
                                                                  "&euro;", adDouble, "capann.gif", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "spiagge", "capf3_stag_", _
                                                                  "Capanne 3^ fila - stagionale", "Huts 3^ row - seasonal cost", "", "", "", _
                                                                  "&euro;", adDouble, "capann.gif", GruppoId)
    
    'associa i descrittori alle categorie delle spiagge
    sql = " SELECT * FROM itb_anagrafiche_tipi " + _
          " WHERE " & import_QueryCategorieByCodice(iCatAnagrafiche, GetCodice(CODE_APTBC, CODE_SPIAGGE, ""), true, true) & " OR " & _
                  import_QueryCategorieByCodice(iCatAnagrafiche, GetCodice(CODE_APTJE, CODE_SPIAGGE, ""), true, true) & " OR " & _
                  import_QueryCategorieByCodice(iCatAnagrafiche, GetCodice(CODE_APTCH, CODE_SPIAGGE, ""), true, true) & " OR " & _
                  import_QueryCategorieByCodice(iCatAnagrafiche, GetCodice(CODE_APTVE, CODE_SPIAGGE, ""), true, true)
    CALL import_Associa_DescrittoriCategorieAnagrafiche(conn, rs, sql, DesId_list, false)
    
    set rs = nothing
end sub


'.................................................................................................................
'funzione che importa i descrittori dei luoghi
'.................................................................................................................
sub import_DESCRITTORI_ANAGRAFICHE__LUOGHI() 
    dim GruppoId, DesId_list
    dim rs, sql
    
    set rs = Server.CreateObject("ADODB.Recordset")
    
    'prepara struttura gruppi e descrittori
    GruppoId = import_GruppoDescrittoreAnagrafiche(conn, rs, "general", "prezzi", "Prezzi", "Prices", "", "", "", 10)
	
    DesId_list = import_DescrittoreAnagrafiche(conn, rs, "general", "apertoPubblico", _
                                               "Aperto al pubblico", "Open to the public", "", "", "", _
                                               "", adBoolean, "", NULL)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "accessibileDisabili", _
                                                                  "Accessibile ai disabili", "Suitable for disabled", "", "", "", _
                                                                  "", adBoolean, "", NULL)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "accessibileDisabiliInfo", _
                                                                  "Informazioni sull'accessibilit&agrave; ai disabili", "Information about suitable for disabled", "", "", "", _
                                                                  "", adLongVarChar, "", NULL)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "chiusoDal", _
                                                                  "Chiuso dal", "Closed from", "", "", "", _
                                                                  "", adVarChar, "", NULL)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "chiusoAl", _
                                                                  "Chiuso al", "Closed to", "", "", "", _
                                                                  "", adVarChar, "", NULL)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "mezziPubblici", _
                                                                  "Mezzi pubblici per raggiungerlo", "Public means of transport to arrive", "", "", "", _
                                                                  "", adVarChar, "", NULL)
                                          
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "prezzoIntero", _
                                                                  "Prezzo intero", "Price", "", "", "", _
                                                                  "", adVarChar, "", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "prezzoRidotto", _
                                                                  "Prezzo ridotto", "Price reduced", "", "", "", _
                                                                  "", adVarChar, "", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "ridottoPer", _
                                                                  "Prezzo ridotto per", "Price reduced to", "", "", "", _
                                                                  "", adVarChar, "", GruppoId)
    
    
    'associa i descrittori alle categorie delle spiagge
    sql = " SELECT * FROM itb_anagrafiche_tipi " + _
          " WHERE " & import_QueryCategorieByCodice(iCatAnagrafiche, CODE_APTBC & CODE_LUOGHI, true, true) & " OR " & _
                  import_QueryCategorieByCodice(iCatAnagrafiche, CODE_APTJE & CODE_LUOGHI, true, true) & " OR " & _
                  import_QueryCategorieByCodice(iCatAnagrafiche, CODE_APTCH & CODE_LUOGHI, true, true) & " OR " & _
                  import_QueryCategorieByCodice(iCatAnagrafiche, CODE_APTVE & CODE_LUOGHI, true, true)
    CALL import_Associa_DescrittoriCategorieAnagrafiche(conn, rs, sql, DesId_list, false)
    
    set rs = nothing
end sub


'.................................................................................................................
'funzione che importa i descrittori dei locali e servizi
'.................................................................................................................
sub import_DESCRITTORI_ANAGRAFICHE__LOCALI_E_SERVIZI() 
    dim GruppoId, DesId_list
    dim rs, sql
    
    set rs = Server.CreateObject("ADODB.Recordset")
    
    'prepara struttura gruppi e descrittori
    GruppoId = import_GruppoDescrittoreAnagrafiche(conn, rs, "general", "prezzi", "Prezzi", "Prices", "", "", "", 10)
	
    DesId_list = import_DescrittoreAnagrafiche(conn, rs, "general", "apertoPubblico", _
                                               "Aperto al pubblico", "Open to the public", "", "", "", _
                                               "", adBoolean, "", NULL)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "chiusoDal", _
                                                                  "Chiuso dal", "Closed from", "", "", "", _
                                                                  "", adVarChar, "", NULL)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "chiusoAl", _
                                                                  "Chiuso al", "Closed to", "", "", "", _
                                                                  "", adVarChar, "", NULL)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "mezziPubblici", _
                                                                  "Mezzi pubblici per raggiungerlo", "Public means of transport to arrive", "", "", "", _
                                                                  "", adVarChar, "", NULL)
                                          
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "prezzoIntero", _
                                                                  "Prezzo intero", "Price", "", "", "", _
                                                                  "", adVarChar, "", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "prezzoRidotto", _
                                                                  "Prezzo ridotto", "Price reduced", "", "", "", _
                                                                  "", adVarChar, "", GruppoId)
    DesId_list = DesId_list & "," & import_DescrittoreAnagrafiche(conn, rs, "general", "ridottoPer", _
                                                                  "Prezzo ridotto per", "Price reduced to", "", "", "", _
                                                                  "", adVarChar, "", GruppoId)
    
    'associa i descrittori alle categorie delle spiagge
    sql = " SELECT * FROM itb_anagrafiche_tipi " + _
          " WHERE " & import_QueryCategorieByCodice(iCatAnagrafiche, CODE_APTBC & CODE_LOCALI, true, true) & " OR " & _
                  import_QueryCategorieByCodice(iCatAnagrafiche, CODE_APTJE & CODE_LOCALI, true, true) & " OR " & _
                  import_QueryCategorieByCodice(iCatAnagrafiche, CODE_APTCH & CODE_LOCALI, true, true) & " OR " & _
                  import_QueryCategorieByCodice(iCatAnagrafiche, CODE_APTVE & CODE_LOCALI, true, true)
    CALL import_Associa_DescrittoriCategorieAnagrafiche(conn, rs, sql, DesId_list, false)
    
    set rs = nothing
end sub


'.................................................................................................................
'funzione che importa i descrittori delle strutture ricettive non sincronizzate
'.................................................................................................................
sub import_DESCRITTORI_ANAGRAFICHE__STRUTTURE_RICETTIVE_NON_SYNCRO() 
    dim GruppoId, DesId_list
    dim rs, sql
    
    set rs = Server.CreateObject("ADODB.Recordset")
    
    'Totale camere
	DesId_list = import_StruttureRicettive_NonSyncro__Descrittore(conn, rs, "tb_dotazioni", 173)
	'Totale posti letto
	DesId_list = DesId_list & ", " & import_StruttureRicettive_NonSyncro__Descrittore(conn, rs, "tb_dotazioni", 201)
	'collega Totale bagni
	DesId_list = DesId_list & ", " & import_StruttureRicettive_NonSyncro__Descrittore(conn, rs, "tb_dotazioni", 207)
	'collega Parco o giardino
	DesId_list = DesId_list & ", " & import_StruttureRicettive_NonSyncro__Descrittore(conn, rs, "tb_servizi", 121)
	'collega Ristorante
	DesId_list = DesId_list & ", " & import_StruttureRicettive_NonSyncro__Descrittore(conn, rs, "tb_servizi", 114)
    
    'associa i descrittori alle categorie delle spiagge
    sql = " SELECT * FROM itb_anagrafiche_tipi " + _
          " WHERE " & import_QueryCategorieByCodice(iCatAnagrafiche, CODE_APTBC & CODE_RICETTIVITA, true, true) & " OR " & _
                  import_QueryCategorieByCodice(iCatAnagrafiche, CODE_APTJE & CODE_RICETTIVITA, true, true) & " OR " & _
                  import_QueryCategorieByCodice(iCatAnagrafiche, CODE_APTCH & CODE_RICETTIVITA, true, true) & " OR " & _
                  import_QueryCategorieByCodice(iCatAnagrafiche, CODE_APTVE & CODE_RICETTIVITA, true, true)
    CALL import_Associa_DescrittoriCategorieAnagrafiche(conn, rs, sql, DesId_list, false)
    
end sub


'.................................................................................................................
'collega la categora al descrittore individuandolo tramite i parametri di sincronizzazione
'.................................................................................................................
function import_StruttureRicettive_NonSyncro__Descrittore(conn, rs, descr_external_source, descr_external_id)
	dim sql, value
	sql = " SELECT and_id FROM itb_anagrafiche_descrittori " + _
		  " WHERE and_external_id LIKE '%(" & descr_external_id & ")%' AND " + _
		  "		  and_external_source LIKE '" & descr_external_source & "' "
	import_StruttureRicettive_NonSyncro__Descrittore = cInteger(GetValueList(conn, rs, sql))
end function


'.................................................................................................................
'inserice valore al descrittore individuandolo tramite i parametri di sincronizzazione
'.................................................................................................................
sub import_StruttureRicettive_NonSyncro__ValoreDescrittore(conn, connApt, rs, ValueQuery, IsValueNumeric, Anagrafica, ExternalSource, ExternalId)
	dim value
	rs.open ValueQuery, connApt, adOpenStatic, adLockOptimistic, adCmdText
	if rs.eof then
		rs.close
	else
		value = rs(0)
		rs.close
		if IsValueNumeric then
			if cInteger(value)>0 then
				CALL Syncro_AnagraficheValoriDescrittori(conn, rs, Anagrafica, ExternalSource, ExternalId, value, "", "", "", "")
			end if
		else
			if value then
				CALL Syncro_AnagraficheValoriDescrittori(conn, rs, Anagrafica, ExternalSource, ExternalId, value, "", "", "", "")
			end if
		end if
	end if
end sub


'*****************************************************************************************************************
'FUNZIONI PER IMPORT DEGLI EVENTI
'*****************************************************************************************************************


'.................................................................................................................
'import degli eventi speciali
'.................................................................................................................
sub import_EVENTI_SPECIALI(AptCode)
    dim connApt, Apt_rs, rs, sql
    
    set rs = Server.CreateObject("ADODB.Recordset")
    set Apt_rs = Server.CreateObject("ADODB.Recordset")
    
    Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
    
    response.write "<!-- Import Eventi speciali " & AptCode & " -->" + vbCrLf
	sql = "SELECT * FROM EventiSpeciali"
	Apt_rs.open sql, connApt, adOpenStatic, adLockOptimistic
	while not Apt_rs.eof
		
		sql = "SELECT * FROM itb_eventi_tipologie WHERE evt_nome_it = '"& ParseSQL(Apt_rs("ev_speciale"), adChar) &"'"
		rs.open sql, conn, adOpenStatic, adLockOptimistic
		if rs.eof then
			rs.addNew
			rs("evt_nome_it") = Apt_rs("ev_speciale")
			rs("evt_nome_en") = Apt_rs("ev_spec_eng")
			rs("evt_visibile") = true
			rs.update
		end if
		rs.close
        
		Apt_rs.movenext
	wend
	Apt_rs.close
    
    connApt.close
    set ConnApt = nothing
    set Apt_rs = nothing
    set rs = nothing
end sub


'.................................................................................................................
'import degli eventi
'.................................................................................................................
sub import_EVENTI(AptCode)
    dim connApt, rsS, rsA, rsP, rsL, rsD, sql, value
    dim CategoriaCodice, CategoriaPrincipale, CategoriaAlternativa
    dim CntId, EveCodice, AnaCodice
    
    set rsS = Server.CreateObject("ADODB.Recordset")
	set rsA = Server.CreateObject("ADODB.Recordset")
	set rsP = Server.CreateObject("ADODB.Recordset")
	set rsL = Server.CreateObject("ADODB.Recordset")
    set rsD = Server.CreateObject("ADODB.Recordset")
    
    dim NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt
	CALL import_CaricaConfigurazione(conn, rsS, NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt)
	
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
    
    response.write "<!-- Import Eventi " & AptCode & " -->" + vbCrLf
    
    '.................................................................................................................
    'inserimento descrittori
	dim gestioneGruppi, prenotazioni
	gestioneGruppi = import_DescrittoreEventi(conn, rs, "Gestione gruppi", "For groups", "", "", "", "", adLongVarChar)
	prenotazioni = import_DescrittoreEventi(conn, rs, "Prenotazioni", "Bookings", "", "", "", "", adLongVarChar)
    
    'collegamento descrittori categorie
	sql = "SELECT * FROM itb_eventi_categorie"
	rsS.open sql, conn, adOpenStatic, adLockOptimistic
	while not rsS.eof
		
		CALL import_AssociazioneDescrittoriCategorieEventi(conn, rs, rsS("evc_id"), gestioneGruppi, 1)
		CALL import_AssociazioneDescrittoriCategorieEventi(conn, rs, rsS("evc_id"), prenotazioni, 2)
		
		rsS.movenext
	wend
	rsS.close
    '.................................................................................................................
    
    sql = "SELECT * FROM irel_luoghi"
	rsL.open sql, conn, adOpenKeyset, adLockOptimistic
    
    'recupera elenco eventi da APT
	sql = " SELECT *, e.descrizione AS eve_descrizione, c.descrizione AS evc_descrizione FROM Eventi e"& _
		  " INNER JOIN CategorieEventi c ON e.id_categoria = c.idc " & _
		  " LEFT JOIN EventiSpeciali s ON e.id_ev_spec = s.id_sp "
	rsS.open sql, connApt, adOpenStatic, adLockOptimistic
    
    while not rsS.eof 
        EveCodice = GetCodice(AptCode, CODE_EVENTI, rsS("id"))
        CategoriaCodice = GetCodice(AptCode, CODE_EVENTI, rsS("id_categoria"))
        CategoriaPrincipale = import_GetCategoriaPrincipale(iCatEventi, conn, rs, CategoriaCodice)
        CategoriaAlternativa = import_GetCategoriaAlternativa(iCatEventi, conn, rs, CategoriaCodice) 
        
        %>
        <!-- <%= EveCodice %> - <%= rsS.absoluteposition %> su <%= rsS.recordcount %>-->
        <%

        if CategoriaPrincipale>0 AND CategoriaAlternativa>0 then
        
    		sql = "SELECT * FROM itb_eventi WHERE eve_codice LIKE '" & EveCodice & "'"
    		rsA.open sql, conn, adOpenKeyset, adLockOptimistic
    		if rsA.eof then
    			rsA.Addnew
    			rsA("eve_codice") = EveCodice
    		end if
            rsA("eve_categoria_id") = CategoriaPrincipale
            rsA("eve_alt_categoria_id") = CategoriaAlternativa
            rsA("eve_titolo_it") = TextDecode(rsS("titolo"))
    		rsA("eve_titolo_en") = TextDecode(rsS("titol_eng"))
    		rsA("eve_descr_it") = TextDecode(rsS("eve_descrizione"))
    		rsA("eve_descr_en") = TextDecode(rsS("descr_eng"))
    		rsA("eve_telefono") = TextDecode(rsS("ev_tel_org"))
    		rsA("eve_ingresso_intero") = TextDecode(rsS("ingresso_intero"))
    		rsA("eve_ingresso_ridotto") = TextDecode(rsS("ingresso_ridotto"))
    		rsA("eve_ridotto_it") = TextDecode(rsS("ridotto_per"))
    		rsA("eve_ridotto_en") = TextDecode(rsS("rid_eng"))
    		rsA("eve_info_it") = TextDecode(rsS("info"))
    		rsA("eve_info_en") = TextDecode(rsS("info_eng"))
    		rsA("eve_bussola_it") = TextDecode(rsS("Descr_Bussola_Ita"))
    		rsA("eve_bussola_en") = TextDecode(rsS("Descr_Bussola_eng"))
    		rsA("eve_ranking") = 150
    		rsA("eve_visibile") = 1
    		rsA("eve_insData") = Now()
    		rsA("eve_insAdmin_id") = NEXTAIM_ADMIN_ID
    		rsA("eve_modData") = Now()
    		rsA("eve_modAdmin_id") = NEXTAIM_ADMIN_ID
    		value = GetValueList(connApt, NULL, "SELECT MIN(dal) FROM PeriodiEventi WHERE id_evento = "& rsS("id"))
    		if value <> "" then
    			rsA("eve_pubblData") = value
    		end if
    		'collegamento con eventi speciali
    		value = GetValueList(conn, rs, "SELECT evt_id FROM itb_eventi_tipologie WHERE evt_nome_it = '"& ParseSQL(rsS("ev_speciale"), adChar) &"'")
    		if value <> "" then
    			rsA("eve_tipologia_id") = value
    		end if
    		rsA.Update
    		
    		'gestione descrittori
    		CALL import_EventiValoriDescrittori(conn, rs, rsA("eve_id"), gestioneGruppi, rsS("gestione_gruppi"), rsS("gegr_eng"), "", "", "")
    		CALL import_EventiValoriDescrittori(conn, rs, rsA("eve_id"), prenotazioni, rsS("prenotazioni"), rsS("pren_eng"), "", "", "")
            
            'gestione periodi
    		sql = "SELECT * FROM doveAccade WHERE id_evento = "& rsS("id")
    		rs.open sql, connApt, adOpenStatic, adLockOptimistic
    		if not rs.eof then
    			while not rs.eof
    				'inserisco luogo
    				rsL.Addnew
    				rsL("rlu_evento_id") = rsA("eve_id")
                    
    				sql = " SELECT ana_id, ana_area_id FROM itb_anagrafiche WHERE ana_codice LIKE '" & GetCodice(AptCode, CODE_LUOGHI, rs("id_luogo")) & "' "
                    rsd.open sql, conn, adOpenStatic, adLockOptimistic
                    if rsd.eof then
                        'luogo non trovato in anagrafiche: recupera area da APT sorgente
    					rsL("rlu_area_id") = import_GetAptArea(AptCode, 0, 0)
    					rsL("rlu_descr_it") = TextDecode(rsS("descr_luogo"))
    					rsL("rlu_descr_en") = TextDecode(rsS("Descr_luogo_eng"))
                    else
                        'luogo collegato
                        rsL("rlu_anagrafica_id") = rsd("ana_id")
                        rsL("rlu_area_id") = rsd("ana_area_id")
                    end if
					rsd.close
    				rsL.Update
    				
    				'inserimento date
    				sql = "SELECT * FROM PeriodiEventi WHERE id_evento = "& rsS("id")
    				rsP.open sql, connApt, adOpenStatic, adLockOptimistic
    				while not rsP.eof
    					sql = " INSERT INTO irel_periodi (rpe_luogo_id, rpe_descr_it, rpe_dal, rpe_al)"& _
    						  " VALUES ("& rsL("rlu_id") &", '"& ParseSQL(IIF(CString(rsS("orario")) <> "", rsS("orario"), rsP("alleOre")), adChar) &"', "& _
    						  SQL_Date(conn, rsP("dal")) &", "& SQL_Date(conn, rsP("al")) &")"
    					CALL conn.Execute(sql)
    					
    					rsP.movenext
    				wend
    				rsP.close
    				
    				rs.movenext
    			wend
    		else
    			'inserisco luogo descrittivo
    			rsL.Addnew
    			rsL("rlu_evento_id") = rsA("eve_id")
    			rsL("rlu_descr_it") = TextDecode(rsS("descr_luogo"))
    			rsL("rlu_descr_en") = TextDecode(rsS("Descr_luogo_eng"))
    			rsL("rlu_area_id") = import_GetAptArea(AptCode, 0, 0)
    			rsL.Update
    			
    			'inserimento date
    			sql = "SELECT * FROM PeriodiEventi WHERE id_evento = "& rsS("id")
    			rsP.open sql, connApt, adOpenStatic, adLockOptimistic
    			while not rsP.eof
    				sql = " INSERT INTO irel_periodi (rpe_luogo_id, rpe_descr_it, rpe_dal, rpe_al)"& _
    					  " VALUES ("& rsL("rlu_id") &", '"& ParseSQL(IIF(CString(rsS("orario")) <> "", rsS("orario"), rsP("alleOre")), adChar) &"', "& _
    					  SQL_Date(conn, rsP("dal")) &", "& SQL_Date(conn, rsP("al")) &")"
    				conn.Execute(sql)
    				
    				rsP.movenext
    			wend
    			rsP.close
    		end if
    		rs.close
            
            'import delle immagini
    		sql = "SELECT * FROM imag_ev WHERE id_ev = " & rsS("id")
    		rsP.open sql, connApt, adOpenStatic, adLockOptimistic, adCmdText
    		
    		sql = "DELETE FROM irel_eventi_img WHERE evi_evento_id=" & rsA("eve_id")
    		CALL conn.execute(sql)
    		
    		sql = "SELECT * FROM irel_eventi_img WHERE evi_evento_id=" & rsA("eve_id")
    		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    	
    		while not rsP.eof
    			rs.AddNew
    			rs("evi_evento_id") = rsA("eve_id")
    			rs("evi_visibile") = true
    			rs("evi_thumb") = AptCode & "/thumb/" & rsP("link")
    			rs("evi_zoom") = AptCode & "/zoom/" & rsP("link")
    			rs("evi_didascalia_it") = TextDecode(rsP("Descriz"))
    			rs("evi_didascalia_en") = TextDecode(rsP("Desc_eng"))
    			rs.Update
    			rsP.movenext
    		wend
    	
    		rs.close
    		rsP.close
            
            
        else
            'categorie per l'import delle anagrafiche non trovate
            response.write "CATEGORIE NON TROVATE<br>" & _
                           "Categoria Principale:" & CategoriaPrincipale & "<br>" & _
                           "Categoria Alternativa:" & CategoriaAlternativa & "<br>"
            response.end
        end if
        
		rsS.movenext
		rsA.close
	wend
	rsS.close
	rsL.close
	
	
	connApt.close
	
	oCnt.conn = empty
	set oCnt = nothing
	
	set connApt = nothing
	set rsS = nothing
	set rsD = nothing
	set rsA = nothing
	set rsP = nothing
	set rsL = nothing
end sub



'.................................................................................................................
'funzione che ritorna i dati del descrittore indicato dal nome
'.................................................................................................................
function import_DescrittoreEventi(conn, rs, Nome_it, Nome_en, Nome_fr, Nome_de, Nome_es, Unita, Tipo)
	dim sql
	
	sql = " SELECT * FROM itb_eventi_descrittori WHERE evd_nome_it = '"& nome_it &"'"
	rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
	if rs.eof then
		rs.AddNew
		rs("evd_nome_it") = Nome_it
		rs("evd_nome_en") = Nome_en
		rs("evd_nome_fr") = Nome_fr
		rs("evd_nome_de") = Nome_de
		rs("evd_nome_es") = Nome_es
		rs("evd_unita_it") = Unita
		rs("evd_principale") = false
		rs("evd_tipo") = Tipo
		rs.update
	end if
	import_DescrittoreEventi = rs("evd_id")
	rs.close
	
end function 


'.................................................................................................................
'funzione che associa il descrittore alla categoria indicata
'.................................................................................................................
function import_AssociazioneDescrittoriCategorieEventi(conn, rs, categoria, descrittore, ordine)
	sql = " SELECT * FROM irel_evCategorie_descrittori WHERE " + _
		  " rcd_categoria_id=" & categoria & " AND rcd_descrittore_id=" & descrittore
	rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
	if rs.eof then
		rs.AddNew
		rs("rcd_categoria_id") = categoria
		rs("rcd_descrittore_id") = descrittore
	end if
	if cString(ordine)<>"" then
		rs("rcd_ordine") = ordine
	elseif IsNull(rs("rtd_ordine")) then
		rs("rcd_ordine") = descrittore
	end if
	
	rs.update
	import_AssociazioneDescrittoriCategorieEventi = rs("rcd_id")
	rs.close
	
end function


'.................................................................................................................
'funzione che salva valore del descrittore
'.................................................................................................................
sub import_EventiValoriDescrittori(conn, rs, evento, descrittore, valore_it, valore_en, valore_fr, valore_de, valore_es)
	dim sql
	dim DesId, DesTipo
	sql = " SELECT * FROM itb_eventi_descrittori " + _
		  " WHERE evd_id = "& descrittore
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	DesId = rs("evd_id")
	DesTipo = rs("evd_tipo")
	rs.close
	
	'inserisco solo un valore non vuoto
	if CString(valore_it) <> "" then
		'inserisce valore/i
		sql = " SELECT * FROM irel_eventi_DescrCat " + _
			  " WHERE red_evento_id=" & evento & " AND red_descrittore_id=" & DesId
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		if rs.eof then
			rs.AddNew
			rs("red_evento_id") = evento
			rs("red_descrittore_id") = DesId
		end if
		Select Case DesTipo
		case adBoolean
			if valore_it then
				rs("red_valore") = "1"
			end if
		Case adDouble
			rs("red_valore_it") = TextDecode(valore_it)
			rs("red_memo_it") = TextDecode(valore_en)
		case adLongVarChar
			rs("red_memo_it") = TextDecode(valore_it)
			rs("red_memo_en") = TextDecode(valore_en)
			rs("red_memo_fr") = TextDecode(valore_fr)
			rs("red_memo_de") = TextDecode(valore_de)
			rs("red_memo_es") = TextDecode(valore_es)
		case else
			rs("red_valore_it") = TextDecode(valore_it)
			rs("red_valore_en") = TextDecode(valore_en)
			rs("red_valore_fr") = TextDecode(valore_fr)
			rs("red_valore_de") = TextDecode(valore_de)
			rs("red_valore_es") = TextDecode(valore_es)
		end select
		rs.Update
		rs.close
	end if
end sub



'*****************************************************************************************************************
'FUNZIONI PER IMPORT DELLE ANAGRAFICHE
'*****************************************************************************************************************


'.................................................................................................................
'import delle spiagge
'.................................................................................................................
sub import_SPIAGGE(AptCode)
    dim connApt, Apt_rs, rs, sql
    dim CategoriaCodice, CategoriaPrincipale, CategoriaAlternativa
    dim CntId, AnaCodice
    
    set rs = Server.CreateObject("ADODB.Recordset")
    set Apt_rs = Server.CreateObject("ADODB.Recordset")
    
    dim NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt
	CALL import_CaricaConfigurazione(conn, rs, NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt)
	
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
    
    response.write "<!-- Import spiagge " & AptCode & " -->" + vbCrLf
	sql = "SELECT * FROM Spiagge "
	Apt_rs.open sql, connApt, adOpenStatic, adLockOptimistic
    
    'recupera dati delle categorie
    CategoriaCodice = GetCodice(AptCode, CODE_SPIAGGE, "")
    CategoriaPrincipale = import_GetCategoriaPrincipale(iCatAnagrafiche, conn, rs, CategoriaCodice)
    CategoriaAlternativa = import_GetCategoriaAlternativa(iCatAnagrafiche, conn, rs, CategoriaCodice)
    
    if CategoriaPrincipale>0 AND CategoriaAlternativa>0 then
    
        while not Apt_rs.eof
            
            oCnt.RemoveAll
		    
            AnaCodice = GetCodice(AptCode, CODE_SPIAGGE, Apt_rs("ID_spiagge"))
            
            %>
            <!-- <%= AnaCodice %> - <%= Apt_rs.absoluteposition %> su <%= Apt_rs.recordcount %>-->
            <%
            
		    'verifica se record collegato esiste gia'
		    sql = " SELECT ana_id FROM itb_anagrafiche WHERE ana_codice LIKE '" & AnaCodice & "' "
            CntId = cInteger(GetValueList(conn, rs, sql))
            
            if CntId > 0 then
                'carica record da database
			    oCnt.LoadFromDB(CntId)
			    oCnt("IDElencoIndirizzi") = CntId
		    end if
            
            'inserice dati contatto
    		oCnt("rubrica") = RUBRICA_ANAGRAFICHE
    		
    		oCnt("IsSocieta") = true
    		oCnt("NomeOrganizzazioneElencoIndirizzi") = Apt_rs("Denominazione")
    		oCnt("indirizzoElencoIndirizzi") = Apt_rs("Indir1_estate")
    		oCnt("CapElencoIndirizzi") = Apt_rs("cap")
    		oCnt("LocaltiaElencoIndirizzi") = Apt_rs("localita")
    		oCnt("CittaElencoIndirizzi") = Apt_rs("comune")
    		
    		oCnt("telefono") = Apt_rs("Tel1_estate")
    		oCnt("fax") = Apt_rs("Fax_estate")
    		oCnt("email") = Apt_rs("Email")
    		oCnt("web") = Apt_rs("web")
    				
    		if CntId > 0 then
    			'aggiorna contatto se esistente
    			oCnt.UpdateDB()
    		else
    			'inserisce nuovo contatto
    			CntId = oCnt.InsertIntoDB()
    		end if
            
            'imposta dati anagrafica
    		sql = "SELECT * FROM itb_anagrafiche WHERE ana_id=" & CntId
    		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    		if rs.eof then
    			rs.AddNew
    			rs("ana_id") = CntId
    			rs("ana_insData") = Now()
    			rs("ana_insAdmin_id") = NEXTAIM_ADMIN_ID
    		end if
            rs("ana_codice") = AnaCodice
    		rs("ana_modData") = Now()
    		rs("ana_modAdmin_id") = NEXTAIM_ADMIN_ID
    		rs("ana_tipo_id") = CategoriaPrincipale
            rs("ana_alt_tipo_id") = CategoriaAlternativa
    		rs("ana_area_id") = import_GetAptArea(AptCode, Apt_rs("zona"), Apt_rs("rif_subzona"))
    		rs("ana_descr_it") = TextDecode(Apt_rs("note_ita"))
    		rs("ana_descr_en") = TextDecode(Apt_rs("note_eng"))
    		rs("ana_visibile") = true
    		rs("ana_censurato") = false
    		rs("ana_ranking") = 150
    		rs.update
    		rs.close
            
            
    		'registra dati aggiuntivi nei descrittori
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "apertura", Apt_rs("apertura_stagionale"), "", "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "orario", Apt_rs("orari"), "", "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "ombr_giorn_", Apt_rs("ombr_giorn_min"), Apt_rs("ombr_giorn_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "ombr_mens_", Apt_rs("ombr_mens_min"), Apt_rs("ombr_mens_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "ombr_stag_", Apt_rs("ombr_stag_min"), Apt_rs("ombr_stag_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "lett_giorn_", Apt_rs("lett_giorn_min"), Apt_rs("lett_giorn_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "lett_mens_", Apt_rs("lett_mens_min"), Apt_rs("lett_mens_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "lett_stag_", Apt_rs("lett_stag_min"), Apt_rs("lett_stag_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "spo_giorn_", Apt_rs("spo_giorn_min"), Apt_rs("spo_giorn_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "spo_mens_", Apt_rs("spo_mens_min"), Apt_rs("spo_mens_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "spo_stag_", Apt_rs("spo_stag_min"), Apt_rs("spo_stag_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "camer_giorn_", Apt_rs("camer_giorn_min"), Apt_rs("camer_giorn_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "camer_mens_", Apt_rs("camer_mens_min"), Apt_rs("camer_mens_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "camer_stag_", Apt_rs("camer_stag_min"), Apt_rs("camer_stag_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "minicap_giorn_", Apt_rs("minicap_giorn_min"), Apt_rs("minicap_giorn_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "minicap_mens_", Apt_rs("minicap_mens_min"), Apt_rs("minicap_mens_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "minicap_stag_", Apt_rs("minicap_stag_min"), Apt_rs("minicap_stag_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "capf1_giorn_", Apt_rs("capf1_giorn_min"), Apt_rs("capf1_giorn_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "capf1_mens_", Apt_rs("capf1_mens_min"), Apt_rs("capf1_mens_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "capf1_stag_", Apt_rs("capf1_stag_min"), Apt_rs("capf1_stag_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "capf2_giorn_", Apt_rs("capf2_giorn_min"), Apt_rs("capf2_giorn_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "capf2_mens_", Apt_rs("capf2_mens_min"), Apt_rs("capf2_mens_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "capf2_stag_", Apt_rs("capf2_stag_min"), Apt_rs("capf2_stag_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "capf3_giorn_", Apt_rs("capf3_giorn_min"), Apt_rs("capf3_giorn_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "capf3_mens_", Apt_rs("capf3_mens_min"), Apt_rs("capf3_mens_max"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "spiagge", "capf3_stag_", Apt_rs("capf3_stag_min"), Apt_rs("capf3_stag_max"), "", "", "")
    		
            CALL import_ValoriMappeAnagrafica(conn, rs, CntId, AptCode, Apt_rs("collocazione"), Apt_rs("linkmappe"))
            Apt_rs.movenext
        wend
    else
        'categorie per l'import delle anagrafiche non trovate
        response.write "CATEGORIE NON TROVATE<br>" & _
                       "Categoria Principale:" & CategoriaPrincipale & "<br>" & _
                       "Categoria Alternativa:" & CategoriaAlternativa & "<br>"
        response.end
    end if
    
    Apt_rs.close
    
    connApt.close
	
	oCnt.conn = empty
	set oCnt = nothing
	
	set connApt = nothing
	set Apt_rs = nothing
	set rs = nothing
end sub


'.................................................................................................................
'import dei luoghi
'.................................................................................................................
sub import_LUOGHI(AptCode)
    dim connApt, Apt_rs, rs, rst, sql
    dim CategoriaCodice, CategoriaPrincipale, CategoriaAlternativa
    dim CntId, AnaCodice
    
    set rs = Server.CreateObject("ADODB.Recordset")
    set rst = Server.CreateObject("ADODB.Recordset")
    set Apt_rs = Server.CreateObject("ADODB.Recordset")
    
    dim NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt
	CALL import_CaricaConfigurazione(conn, rs, NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt)
	
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
    
    response.write "<!-- Import luoghi " & AptCode & " -->" + vbCrLf

    sql = "SELECT * FROM luoghi INNER JOIN TipoLuoghi ON Luoghi.id_tipo = TipoLuoghi.idl "
	Apt_rs.open sql, connApt, adOpenStatic, adLockOptimistic
	
	while not Apt_rs.eof

        'recupera dati delle categorie
        CategoriaCodice = GetCodice(AptCode, CODE_LUOGHI, Apt_rs("id_tipo"))
        CategoriaPrincipale = import_GetCategoriaPrincipale(iCatAnagrafiche, conn, rs, CategoriaCodice)
        CategoriaAlternativa = import_GetCategoriaAlternativa(iCatAnagrafiche, conn, rs, CategoriaCodice)
        
        if CategoriaPrincipale>0 AND CategoriaAlternativa>0 then
            oCnt.RemoveAll
		    
            AnaCodice = GetCodice(AptCode, CODE_LUOGHI, Apt_rs("id"))

            %>
            <!-- <%= AnaCodice %> - <%= Apt_rs.absoluteposition %> su <%= Apt_rs.recordcount %>-->
            <%
            
		    'verifica se record collegato esiste gia'
		    sql = " SELECT ana_id FROM itb_anagrafiche WHERE ana_codice LIKE '" & AnaCodice & "' "
            CntId = cInteger(GetValueList(conn, rs, sql))
            
            if CntId > 0 then
                'carica record da database
			    oCnt.LoadFromDB(CntId)
			    oCnt("IDElencoIndirizzi") = CntId
		    end if
            
            'inserice dati contatto
    		oCnt("rubrica") = RUBRICA_ANAGRAFICHE
    		
    		oCnt("IsSocieta") = true
            oCnt("NomeOrganizzazioneElencoIndirizzi") = Apt_rs("nominativo")
    		oCnt("indirizzoElencoIndirizzi") = Apt_rs("indirizzo")
    		oCnt("CapElencoIndirizzi") = Apt_rs("cap")
    		oCnt("LocaltiaElencoIndirizzi") = Apt_rs("frazione")
    		oCnt("CittaElencoIndirizzi") = Apt_rs("citta")
    		
    		oCnt("telefono") = Apt_rs("telefono")
    		oCnt("fax") = Apt_rs("fax")
    		oCnt("email") = Apt_rs("indirizzo_Email")
    		oCnt("web") = Apt_rs("indirizzo_http")
    		
    		if CntId > 0 then
    			'aggiorna contatto se esistente
    			oCnt.UpdateDB()
    		else
    			'inserisce nuovo contatto
    			CntId = oCnt.InsertIntoDB()
    		end if
            
		
    		'imposta dati anagrafica
    		sql = "SELECT * FROM itb_anagrafiche WHERE ana_id=" & CntId
    		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    		if rs.eof then
    			rs.AddNew
    			rs("ana_id") = CntId
    			rs("ana_insData") = Now()
    			rs("ana_insAdmin_id") = NEXTAIM_ADMIN_ID
    		end if
            rs("ana_codice") = AnaCodice
    		rs("ana_modData") = Now()
    		rs("ana_modAdmin_id") = NEXTAIM_ADMIN_ID
    		rs("ana_tipo_id") = CategoriaPrincipale
            rs("ana_alt_tipo_id") = CategoriaAlternativa
    		rs("ana_area_id") = import_GetAptArea(AptCode, Apt_rs("zona"), Apt_rs("subzona"))
            rs("ana_descr_it") = TextDecode(Apt_rs("note"))
    		rs("ana_descr_en") = TextDecode(Apt_rs("note_eng"))
    		rs("ana_visibile") = true
    		rs("ana_censurato") = false
    		rs("ana_ranking") = 150
    		rs.update
    		rs.close
            
            'registra dati aggiuntivi nei descrittori
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "apertura", Apt_rs("orario"), Apt_rs("orar_eng"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "apertoPubblico", Apt_rs("apertura_pubblico"), "", "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "accessibileDisabili", Apt_rs("accessibile_disabili"), "", "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "accessibileDisabiliInfo", Apt_rs("info_disabili_it"), Apt_rs("info_disabili_en"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "chiusoDal", Apt_rs("chiuso_dal"), "", "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "chiusoAl", Apt_rs("chiuso_al"), "", "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "mezziPubblici", Apt_rs("linea_actv"), "", "", "", "")
    		
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "prezzoIntero", Apt_rs("costo_intero"), "", "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "prezzoRidotto", Apt_rs("costo_ridotto"), "", "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "ridottoPer", Apt_rs("ridotto_per"), Apt_rs("rid_eng"), "", "", "")
    		
            CALL import_ValoriMappeAnagrafica(conn, rs, CntId, AptCode, Apt_rs("collocazione"), Apt_rs("linkmappe"))
            
            'gestione delle foto dei luoghi
        	CALL Import_ImmaginiAnagrafiche(connApt, conn, rs, rst, AptCode, "imag_lu", "id_luogo", Apt_rs("id"), CntId)
        else
            'categorie per l'import delle anagrafiche non trovate
            response.write "CATEGORIE NON TROVATE<br>" & _
                           "Categoria Principale:" & CategoriaPrincipale & "<br>" & _
                           "Categoria Alternativa:" & CategoriaAlternativa & "<br>"
            response.end
        end if
        
        Apt_rs.movenext
    wend
    
    Apt_rs.close
    
    connApt.close
	
	oCnt.conn = empty
	set oCnt = nothing
	
	set connApt = nothing
	set Apt_rs = nothing
	set rs = nothing
    set rst = nothing
end sub


'.................................................................................................................
'import delle notizie utili
'................................................................................................................
sub import_NOTIZIE_UTILI(AptCode)
    dim connApt, rsS, rsA, rst, sql
    dim CategoriaCodice, CategoriaPrincipale, CategoriaAlternativa
    dim CntId, AnaCodice
    
    set rsS = Server.CreateObject("ADODB.Recordset")
	set rsA = Server.CreateObject("ADODB.Recordset")
	set rsT = Server.CreateObject("ADODB.Recordset")
    
    dim NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt
	CALL import_CaricaConfigurazione(conn, rs, NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt)
	
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
    
    response.write "<!-- Import Notizie utili " & AptCode & " -->" + vbCrLf
	sql = "SELECT * FROM not_util INNER JOIN Tipi_notutil ON Not_util.tipo = tipi_notutil.id_tipoutil "
	rsS.open sql, connApt, adOpenStatic, adLockOptimistic
	
	while not rsS.eof
        CategoriaPrincipale = 0
        CategoriaAlternativa = 0
		
        'recupera categoria principale e categoria alternativa
        'verifica per primi i sottotipi
        sql = "SELECT rel_sTipoNot_sTipo FROM rel_sottoTipi_NotUtil WHERE rel_sTipoNot_Not=" & rsS("id_UTIL")
        rsT.open sql, connApt, adOpenStatic, adLockOptimistic
        while not rsT.eof AND CategoriaPrincipale = 0
        
            CategoriaCodice = GetCodice(AptCode, CODE_NOTIZIE, rsT("rel_sTipoNot_sTipo"))
            
            CategoriaPrincipale = cIntero(import_GetCategoriaPrincipale(iCatAnagrafiche, conn, rs, CategoriaCodice))
            CategoriaAlternativa = cIntero(import_GetCategoriaAlternativa(iCatAnagrafiche, conn, rs, CategoriaCodice))
            rsT.movenext
        wend
        rsT.close
        'categoria non trovata tra i sottotipi, la cerca nel tipo
        if CategoriaPrincipale = 0 OR CategoriaAlternativa = 0 then
            
            CategoriaCodice = GetCodice(AptCode, CODE_NOTIZIE_T, rsS("Tipo"))
            
            CategoriaPrincipale = cIntero(import_GetCategoriaPrincipale(iCatAnagrafiche, conn, rs, CategoriaCodice))
            CategoriaAlternativa = cIntero(import_GetCategoriaAlternativa(iCatAnagrafiche, conn, rs, CategoriaCodice))
        end if
        
         if CategoriaPrincipale>0 AND CategoriaAlternativa>0 then
    		oCnt.RemoveAll
            
            AnaCodice = GetCodice(AptCode, CODE_LUOGHI, rsS("id_UTIL"))

            %>
            <!-- <%= AnaCodice %> - <%= rsS.absoluteposition %> su <%= rsS.recordcount %>-->
            <%
            
		    'verifica se record collegato esiste gia'
    		sql = " SELECT ana_id FROM itb_anagrafiche WHERE ana_codice LIKE '" & AnaCodice & "' "
            CntId = cInteger(GetValueList(conn, rs, sql))
                
    		if CntId > 0 then
    			'carica record da database
    			oCnt.LoadFromDB(CntId)
    			oCnt("IDElencoIndirizzi") = CntId
    		end if
    		
    		'inserice dati contatto
    		oCnt("rubrica") = RUBRICA_ANAGRAFICHE
            
            oCnt("IsSocieta") = true
    		oCnt("NomeOrganizzazioneElencoIndirizzi") = rsS("denom_util")
    		oCnt("indirizzoElencoIndirizzi") = rsS("indir1")
    		oCnt("CapElencoIndirizzi") = rsS("cap")
    		oCnt("CittaElencoIndirizzi") = rsS("Local")
    		
    		oCnt("telefono") = rsS("telef1")
    		oCnt("fax") = rsS("fax")
    		oCnt("email") = rsS("e_mail")
    		oCnt("web") = rsS("web")
            
            
    		if CntId > 0 then
    			'aggiorna contatto se esistente
    			oCnt.UpdateDB()
    		else
    			'inserisce nuovo contatto
    			CntId = oCnt.InsertIntoDB()
    		end if
    		
    		'imposta dati anagrafica
    		sql = "SELECT * FROM itb_anagrafiche WHERE ana_id=" & CntId
    		rsA.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    		if rsA.eof then
    			rsA.AddNew
    			rsA("ana_id") = CntId
    			rsA("ana_insData") = Now()
    			rsA("ana_insAdmin_id") = NEXTAIM_ADMIN_ID
    		end if
            rsA("ana_codice") = AnaCodice
    		rsA("ana_modData") = Now()
    		rsA("ana_modAdmin_id") = NEXTAIM_ADMIN_ID
            rsA("ana_tipo_id") = CategoriaPrincipale
            rsA("ana_alt_tipo_id") = CategoriaAlternativa
    		rsA("ana_area_id") = import_GetAptArea(AptCode, rsS("zona"), rsS("subzona"))
            rsA("ana_descr_it") = TextDecode(rsS("descrizione"))
    		rsA("ana_descr_en") = TextDecode(rsS("descr_eng"))
    		rsA("ana_visibile") = true
    		rsA("ana_censurato") = false
    		rsA("ana_ranking") = 150
    		rsA.update
    		rsA.close
            
            CALL import_ClearImmaginiAnagrafica(conn, CntId)
            CALL import_ImmagineAnagrafica(conn, rs, AptCode, CntId, rsS("image"), "", "", "", "", "", "")
            
            CALL import_ValoriMappeAnagrafica(conn, rs, CntId, AptCode, rsS("rif_mappa"), rsS("linkmappe"))
            
        else
            'categorie per l'import delle anagrafiche non trovate
            response.write "CATEGORIE NON TROVATE<br>" & _
                           "Categoria Principale:" & CategoriaPrincipale & "<br>" & _
                           "Categoria Alternativa:" & CategoriaAlternativa & "<br>"
            response.end
        end if
        
        rsS.movenext
    wend
    
    rsS.close
    
    connApt.close
	
	oCnt.conn = empty
	set oCnt = nothing
	
	set connApt = nothing
	set rsS = nothing
	set rsA = nothing
    set rst = nothing
end sub


'.................................................................................................................
'import dei locali e servizi
'................................................................................................................
sub import_LOCALI_E_SERVIZI(AptCode)
    dim connApt, rsS, rsA, rst, sql
    dim CategoriaCodice, CategoriaPrincipale, CategoriaAlternativa
    dim CntId, AnaCodice
    
    set rsS = Server.CreateObject("ADODB.Recordset")
	set rsA = Server.CreateObject("ADODB.Recordset")
	set rsT = Server.CreateObject("ADODB.Recordset")
    
    dim NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt
	CALL import_CaricaConfigurazione(conn, rs, NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt)
	
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
    
    response.write "<!-- Import Locali e servizi " & AptCode & " -->" + vbCrLf
	sql = "SELECT * FROM LocaliEServizi INNER JOIN Tipi_LS ON LocaliEServizi.tipo = Tipi_LS.id_tipoutil "
	rsS.open sql, connApt, adOpenStatic, adLockOptimistic
	
	while not rsS.eof
        CategoriaPrincipale = 0
        CategoriaAlternativa = 0
		
        'recupera categoria principale e categoria alternativa
        'verifica per primi i sottotipi
        sql = "SELECT rel_sLS_sTipo FROM rel_sottoTipi_LS WHERE rel_sLS_LocaleServizio=" & rsS("id_Ls")
        rsT.open sql, connApt, adOpenStatic, adLockOptimistic
        while not rsT.eof AND CategoriaPrincipale = 0
        
            CategoriaCodice = GetCodice(AptCode, CODE_LOCALI, rsT("rel_sLS_sTipo"))
            
            CategoriaPrincipale = cIntero(import_GetCategoriaPrincipale(iCatAnagrafiche, conn, rs, CategoriaCodice))
            CategoriaAlternativa = cIntero(import_GetCategoriaAlternativa(iCatAnagrafiche, conn, rs, CategoriaCodice))
            rsT.movenext
        wend
        rsT.close
        'categoria non trovata tra i sottotipi, la cerca nel tipo
        if CategoriaPrincipale = 0 OR CategoriaAlternativa = 0 then
            
            CategoriaCodice = GetCodice(AptCode, CODE_LOCALI_T, rsS("Tipo"))
            
            CategoriaPrincipale = cIntero(import_GetCategoriaPrincipale(iCatAnagrafiche, conn, rs, CategoriaCodice))
            CategoriaAlternativa = cIntero(import_GetCategoriaAlternativa(iCatAnagrafiche, conn, rs, CategoriaCodice))
        end if
        
         if CategoriaPrincipale>0 AND CategoriaAlternativa>0 then
    		oCnt.RemoveAll
            
            AnaCodice = GetCodice(AptCode, CODE_LUOGHI, rsS("id_Ls"))

            %>
            <!-- <%= AnaCodice %> - <%= rsS.absoluteposition %> su <%= rsS.recordcount %>-->
            <%
            
    		'verifica se record collegato esiste gia'
    		sql = " SELECT ana_id FROM itb_anagrafiche WHERE ana_codice LIKE '" & AnaCodice & "' "
            CntId = cInteger(GetValueList(conn, rs, sql))
                
    		if CntId > 0 then
    			'carica record da database
    			oCnt.LoadFromDB(CntId)
    			oCnt("IDElencoIndirizzi") = CntId
    		end if
    		
    		'inserice dati contatto
    		oCnt("rubrica") = RUBRICA_ANAGRAFICHE
            
            oCnt("IsSocieta") = true
    		oCnt("NomeOrganizzazioneElencoIndirizzi") = rsS("Denominazione_LS")
    		oCnt("indirizzoElencoIndirizzi") = rsS("indir1")
    		oCnt("CapElencoIndirizzi") = rsS("cap")
    		oCnt("CittaElencoIndirizzi") = rsS("Local")
    		
    		oCnt("telefono") = rsS("telef1")
    		oCnt("fax") = rsS("fax")
    		oCnt("email") = rsS("e_mail")
    		oCnt("web") = rsS("web")
            
            
    		if CntId > 0 then
    			'aggiorna contatto se esistente
    			oCnt.UpdateDB()
    		else
    			'inserisce nuovo contatto
    			CntId = oCnt.InsertIntoDB()
    		end if
    		
    		'imposta dati anagrafica
    		sql = "SELECT * FROM itb_anagrafiche WHERE ana_id=" & CntId
    		rsA.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    		if rsA.eof then
    			rsA.AddNew
    			rsA("ana_id") = CntId
    			rsA("ana_insData") = Now()
    			rsA("ana_insAdmin_id") = NEXTAIM_ADMIN_ID
    		end if
            rsA("ana_codice") = AnaCodice
    		rsA("ana_modData") = Now()
    		rsA("ana_modAdmin_id") = NEXTAIM_ADMIN_ID
            rsA("ana_tipo_id") = CategoriaPrincipale
            rsA("ana_alt_tipo_id") = CategoriaAlternativa
    		rsA("ana_area_id") = import_GetAptArea(AptCode, rsS("zona"), rsS("subzona"))
            rsA("ana_descr_it") = TextDecode(rsS("note_ita"))
    		rsA("ana_descr_en") = TextDecode(rsS("note_eng"))
    		rsA("ana_visibile") = true
    		rsA("ana_censurato") = false
    		rsA("ana_ranking") = 150
    		rsA.update
    		rsA.close
            CALL import_ValoriMappeAnagrafica(conn, rs, CntId, AptCode, rsS("rif_mappa"), rsS("linkmappe"))
            
            'registra dati aggiuntivi nei descrittori
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "apertura", rsS("orario"), rsS("orar_eng"), "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "apertoPubblico", rsS("apertura_pubblico"), "", "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "chiusoDal", rsS("chiuso_dal"), "", "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "chiusoAl", rsS("chiuso_al"), "", "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "mezziPubblici", rsS("linea_actv"), "", "", "", "")
    		
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "prezzoIntero", rsS("costo_intero"), "", "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "prezzoRidotto", rsS("costo_ridotto"), "", "", "", "")
    		CALL Syncro_AnagraficheValoriDescrittori(conn, rs, CntId, "general", "ridottoPer", rsS("ridotto_per"), rsS("rid_eng"), "", "", "")
    		
		    'gestione immagini
		    CALL Import_ImmaginiAnagrafiche(connApt, conn, rs, rsA, AptCode, "imag_ls", "id_ls", rsS("id_ls"), CntId)
            CALL import_ImmagineAnagrafica(conn, rs, AptCode, CntId, rsS("image"), "", "", "", "", "", "")

        else
            'categorie per l'import delle anagrafiche non trovate
            response.write "CATEGORIE NON TROVATE<br>" & _
                           "Categoria Principale:" & CategoriaPrincipale & "<br>" & _
                           "Categoria Alternativa:" & CategoriaAlternativa & "<br>"
            response.end
        end if
        
        rsS.movenext
    wend
    
    rsS.close
    
    connApt.close
	
	oCnt.conn = empty
	set oCnt = nothing
	
	set connApt = nothing
	set rsS = nothing
	set rsA = nothing
    set rst = nothing
end sub


'.................................................................................................................
'import delle strutture ricettive non sincronizzate
'.................................................................................................................
sub import_STRUTTURE_Non_Sincronizzate(AptCode)
    dim connApt, rsS, rsA, rst, sql
    dim CategoriaCodice, CategoriaPrincipale, CategoriaAlternativa
    dim CntId, AnaCodice
    
    set rsS = Server.CreateObject("ADODB.Recordset")
	set rsA = Server.CreateObject("ADODB.Recordset")
	set rsT = Server.CreateObject("ADODB.Recordset")
    
    dim NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt
	CALL import_CaricaConfigurazione(conn, rs, NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt)
	
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
    
    response.write "<!-- Import Strutture ricettive non sincronizzate " & AptCode & " -->" + vbCrLf
	sql = " SELECT * FROM stru_Ric INNER JOIN tipi_ric ON stru_ric.tipo = tipi_ric.IdTipo " + _
	      " WHERE IsNull(ID_tipi_provincia,'')='' "
	rsS.open sql, connApt, adOpenStatic, adLockOptimistic
	
	while not rsS.eof

        'recupera dati delle categorie
        CategoriaCodice = GetCodice(AptCode, CODE_RICETTIVITA, rsS("tipo"))
        CategoriaPrincipale = import_GetCategoriaPrincipale(iCatAnagrafiche, conn, rs, CategoriaCodice)
        CategoriaAlternativa = CategoriaPrincipale
        
        if CategoriaPrincipale>0 AND CategoriaAlternativa>0 then
            oCnt.RemoveAll
		    
            AnaCodice = GetCodice(AptCode, CODE_RICETTIVITA, rsS("id_albergo"))

            %>
            <!-- <%= AnaCodice %> - <%= rsS.absoluteposition %> su <%= rsS.recordcount %>-->
            <%
            
		    'verifica se record collegato esiste gia'
		    sql = " SELECT ana_id FROM itb_anagrafiche WHERE ana_codice LIKE '" & AnaCodice & "' "
            CntId = cInteger(GetValueList(conn, rs, sql))
            
            if CntId > 0 then
                'carica record da database
			    oCnt.LoadFromDB(CntId)
			    oCnt("IDElencoIndirizzi") = CntId
		    end if
            
            'inserice dati contatto
    		oCnt("rubrica") = RUBRICA_ANAGRAFICHE
    		
    		oCnt("IsSocieta") = true
            oCnt("NomeOrganizzazioneElencoIndirizzi") = rsS("denominazione")
    		oCnt("indirizzoElencoIndirizzi") = rsS("indir_1")
    		oCnt("CapElencoIndirizzi") = rsS("cap")
    		oCnt("CittaElencoIndirizzi") = rsS("Localita")
    		
    		oCnt("telefono") = rsS("telefono")
    		oCnt("fax") = rsS("fax")
    		oCnt("email") = rsS("email")
    		oCnt("web") = rsS("web")
    		
    		if CntId > 0 then
    			'aggiorna contatto se esistente
    			oCnt.UpdateDB()
    		else
    			'inserisce nuovo contatto
    			CntId = oCnt.InsertIntoDB()
    		end if
        
            'imposta dati anagrafica
    		sql = "SELECT * FROM itb_anagrafiche WHERE ana_id=" & CntId
    		rsA.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    		if rsA.eof then
    			rsA.AddNew
    			rsA("ana_id") = CntId
    			rsA("ana_insData") = Now()
    			rsA("ana_insAdmin_id") = NEXTAIM_ADMIN_ID
    		end if
            rsA("ana_codice") = AnaCodice
    		rsA("ana_modData") = Now()
    		rsA("ana_modAdmin_id") = NEXTAIM_ADMIN_ID
            rsA("ana_tipo_id") = CategoriaPrincipale
            rsA("ana_alt_tipo_id") = CategoriaAlternativa
    		rsA("ana_area_id") = import_GetAptArea(AptCode, rsS("zona"), rsS("rif_subzona"))
            
            rsA("ana_descr_it") = TextDecode(rsS("descr_it"))
            if cString(rsS("le_note"))<>"" then
                if cString(rsA("ana_descr_it"))<>"" then
                    rsA("ana_descr_it") = rsA("ana_descr_it") & vbCrLf & vbCrLf
                end if
                rsA("ana_descr_it") = rsA("ana_descr_it") & TextDecode(rsS("le_note"))
            end if
            
            rsA("ana_descr_en") = TextDecode(rsS("descr_en"))
            if cString(rsS("le_note_eng"))<>"" then
                if cString(rsA("ana_descr_en"))<>"" then
                    rsA("ana_descr_en") = rsA("ana_descr_en") & vbCrLf & vbCrLf
                end if
                rsA("ana_descr_en") = rsA("ana_descr_en") & TextDecode(rsS("le_note_eng"))
            end if
            
    		rsA("ana_visibile") = true
    		rsA("ana_censurato") = false
    		rsA("ana_ranking") = 150
    		rsA.update
    		rsA.close
            
            
            'registra dati aggiuntivi dei descrittori (dotazioni)
    		'collega Totale camere
    		sql = " SELECT rel_valore FROM rel_ric_caratt INNER JOIN Caratt_tipiRic ON rel_ric_caratt.rel_id_caratt = Caratt_tipiRic.id_caratt " + _
    			  " WHERE rel_id_struric=" & rsS("ID_Albergo") & " AND Caratt_TipiRic.caratt_sigla LIKE 'num_camere'"
    		CALL import_StruttureRicettive_NonSyncro__ValoreDescrittore(conn, connApt, rs, sql, true, CntId, "tb_dotazioni", 173)
    		
    		'collega Totale posti letto
    		sql = " SELECT rel_valore FROM rel_ric_caratt INNER JOIN Caratt_tipiRic ON rel_ric_caratt.rel_id_caratt = Caratt_tipiRic.id_caratt " + _
    			  " WHERE rel_id_struric=" & rsS("ID_Albergo") & " AND Caratt_TipiRic.caratt_sigla LIKE 'num_letti'"
    		CALL import_StruttureRicettive_NonSyncro__ValoreDescrittore(conn, connApt, rs, sql, true, CntId, "tb_dotazioni", 201)
    		
    		'collega Totale bagni
    		sql = " SELECT rel_valore FROM rel_ric_caratt INNER JOIN Caratt_tipiRic ON rel_ric_caratt.rel_id_caratt = Caratt_tipiRic.id_caratt " + _
    			  " WHERE rel_id_struric=" & rsS("ID_Albergo") & " AND Caratt_TipiRic.caratt_sigla LIKE 'num_bagni'"
    		CALL import_StruttureRicettive_NonSyncro__ValoreDescrittore(conn, connApt, rs, sql, true, CntId, "tb_dotazioni", 207)
    
    		'collega Parco o giardino
    		sql = " SELECT COUNT(*) FROM rel_Ric_Servizi INNER JOIN rel_Serv_TipiRic ON rel_Ric_Servizi.rel_relServ_Ric = rel_Serv_TipiRic.rel_tipiRicServ_id " + _
    			  "	INNER JOIN Servizi_TipiRic ON rel_Serv_TipiRic.rel_tRicServ_idServ = Servizi_TipiRic.id_serv_tipiric " + _
    			  " WHERE Servizi_TipiRic.serv_nome_it LIKE 'giardino%'"
    		CALL import_StruttureRicettive_NonSyncro__ValoreDescrittore(conn, connApt, rs, sql, false, CntId, "tb_servizi", 121)
    		
    		'collega Ristorante
    		sql = " SELECT COUNT(*) FROM rel_Ric_Servizi INNER JOIN rel_Serv_TipiRic ON rel_Ric_Servizi.rel_relServ_Ric = rel_Serv_TipiRic.rel_tipiRicServ_id " + _
    			  "	INNER JOIN Servizi_TipiRic ON rel_Serv_TipiRic.rel_tRicServ_idServ = Servizi_TipiRic.id_serv_tipiric " + _
    			  " WHERE Servizi_TipiRic.serv_nome_it LIKE '%ristorante%'"
    		CALL import_StruttureRicettive_NonSyncro__ValoreDescrittore(conn, connApt, rs, sql, false, CntId, "tb_servizi", 114)
    		
            CALL import_ValoriMappeAnagrafica(conn, rs, CntId, AptCode, rsS("collocazione"), rsS("linkmappe"))
            
            'collegamento immagini
            CALL import_ClearImmaginiAnagrafica(conn, CntId)
            CALL import_ImmagineAnagrafica(conn, rs, AptCode, CntId, rsS("foto_int"), rsS("foto_int_big"), "", "", "", "", "")
            CALL import_ImmagineAnagrafica(conn, rs, AptCode, CntId, rsS("foto_est"), rsS("foto_est_big"), "", "", "", "", "")
            
         else
            'categorie per l'import delle anagrafiche non trovate
            response.write "CATEGORIE NON TROVATE<br>" & _
                           "Categoria Principale:" & CategoriaPrincipale & "<br>" & _
                           "Categoria Alternativa:" & CategoriaAlternativa & "<br>"
            response.end
        end if
        
        rsS.movenext
    wend
    
    rsS.close
    
    connApt.close
	
	oCnt.conn = empty
	set oCnt = nothing
	
	set connApt = nothing
	set rsS = nothing
	set rsA = nothing
    set rst = nothing
end sub


'.................................................................................................................
'import delle strutture ricettive sincronizzate con applicativi assessorato
'.................................................................................................................
sub import_STRUTTURE_SINCRONIZZATE(AptCode)
    dim connApt, rsS, rsA, rst, sql
    dim CategoriaCodice, CategoriaPrincipale, CategoriaAlternativa
    dim CntId, AnaCodice
    
    set rsS = Server.CreateObject("ADODB.Recordset")
	set rsA = Server.CreateObject("ADODB.Recordset")
	set rsT = Server.CreateObject("ADODB.Recordset")
    
    dim NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt
	CALL import_CaricaConfigurazione(conn, rs, NEXTAIM_ADMIN_ID, RUBRICA_ANAGRAFICHE, oCnt)
	
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
    
    response.write "<!-- Import Strutture ricettive sincronizzate " & AptCode & " -->" + vbCrLf
	sql = " SELECT * FROM stru_Ric INNER JOIN tipi_ric ON stru_ric.tipo = tipi_ric.IdTipo " + _
	      " WHERE RTrim(IsNull(ID_tipi_provincia,''))<>'' AND IsNull(RegCode,'')<>'' "
	rsS.open sql, connApt, adOpenStatic, adLockOptimistic
    
    while not rsS.eof
        %>
        <!-- <%= rsS("RegCode") %> - <%= rsS.absoluteposition %> su <%= rsS.recordcount %>-->
        <%
        'recupera id struttura nel next-com
		sql = " SELECT IDElencoIndirizzi FROM tb_Indirizzario " + _
			  " WHERE SyncroTable LIKE 'VIEW_valid_strutture' AND SyncroKey='" & rsS("RegCode") & "'"
		CntId = cInteger(GetValueList(conn, rsA, sql))
        
        if CntId>0 then
            sql = "SELECT * FROM itb_anagrafiche WHERE ana_id=" & CntId
			rsA.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
            if not rsA.eof then
                rsA("ana_modData") = Now()
    			rsA("ana_modAdmin_id") = NEXTAIM_ADMIN_ID
                rsA("ana_area_id") = import_GetAptArea(AptCode, rsS("zona"), rsS("rif_subzona"))
                
    			rsA("ana_descr_it") = TextDecode(rsS("descr_it"))
                if cString(rsS("le_note"))<>"" then
                    if cString(rsA("ana_descr_it"))<>"" then
                        rsA("ana_descr_it") = rsA("ana_descr_it") & vbCrLf & vbCrLf
                    end if
                    rsA("ana_descr_it") = rsA("ana_descr_it") & TextDecode(rsS("le_note"))
                end if
                
                rsA("ana_descr_en") = TextDecode(rsS("descr_en"))
                if cString(rsS("le_note_eng"))<>"" then
                    if cString(rsA("ana_descr_en"))<>"" then
                        rsA("ana_descr_en") = rsA("ana_descr_en") & vbCrLf & vbCrLf
                    end if
                    rsA("ana_descr_en") = rsA("ana_descr_en") & TextDecode(rsS("le_note_eng"))
                end if
                rsA.update
                
                'collegamento immagini
                CALL import_ClearImmaginiAnagrafica(conn, CntId)
                CALL import_ImmagineAnagrafica(conn, rs, AptCode, CntId, rsS("foto_int"), rsS("foto_int_big"), "", "", "", "", "")
                CALL import_ImmagineAnagrafica(conn, rs, AptCode, CntId, rsS("foto_est"), rsS("foto_est_big"), "", "", "", "", "")
                
                CALL import_ValoriMappeAnagrafica(conn, rs, CntId, AptCode, rsS("collocazione"), rsS("linkmappe"))
            else %>
                <tr>
                    <th><%= rsS("RegCode") %></th>
                    <th colspan="2">STRUTTURA NON TROVATA (<%= rsS("Denominazione") %>, <%= AptCode %>)</th>
                </tr>
            <%end if
            rsA.close
        end if
        
        rsS.movenext
    wend
    
    rsS.close
    
    connApt.close
	
	oCnt.conn = empty
	set oCnt = nothing
	
	set connApt = nothing
	set rsS = nothing
	set rsA = nothing
    set rst = nothing
end sub

%>