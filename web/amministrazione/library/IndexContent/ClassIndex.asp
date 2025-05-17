<%
'PERMESSI
const prm_pagine_altera 			= 1				'DB: permesso di creazione / modifica / cancellazione delle pagine
const prm_indice_accesso			= 30			'permesso di accesso all'area indice generale
const prm_indice_permessi			= 31			'permesso di modifica dei permessi dell'indice
const prm_indice_trasparente		= 32			'indica se l'utente non ha limitazioni rispetto l'indice
const prm_template_accesso			= 40			'permesso di accesso all'area template
const prm_stili_accesso				= 50			'permesso di accesso all'area stili
const prm_plugin_accesso			= 60			'permesso di accesso all'area plugin
const prm_strumenti_accesso			= 70			'permesso di accesso all'area strumenti
const prm_siti_gestione				= 81			'permesso di creazione / modifica / cancellazione dei siti
const prm_menu_accesso				= 90			'permesso di accesso all'area menu
const prm_pubblicazioni_accesso 	= 100			'permesso di accesso all'area di gestione della pubblicazione automatica dei dati strutturati all'interno dell'albero dei contenuti.
const prm_immaginiFormati_accesso	= 110			'permesso di accesso all'area di gestione dei formati delle immagini


'TIPI LINK
const lnk_interno				= 1
const lnk_esterno				= 2


'POSIZIONE del pulsante di INDICIZZAZIONE
const POS_TESTATA 				= "testata"
const POS_ELENCO				= "elenco"
const POS_INDICE				= "indice"


'TITOLO DELLA TABELLA CHE IDENTIFICA I RAGGRUPPAMENTI
const tabRaggruppamento = "Raggruppamento"
const tabRaggruppamentoTable = "tb_contents_index"
'ALTER TABELLE DI "SISTEMA"
const tabSitoTable = "tb_webs"
const tabPagineTable = "tb_paginesito"


'inizializzazione istanza dell'indice
dim index
set index = New ObjIndex

'*******************************************************************************************************************
'CLASSE DELL'INDICE
Class ObjIndex

    'connessione DB
    Public conn
    
    'dictionary contenente i dati da salvare
    Public dizionario
    
    'oggetto per la gestione dei contenuti
    Public content
    
    'altre prorieta e variabili interne
    Public OrdineLenght						'# max cifre che puo avere il campo ordine
	
	'se a true disattiva la ricorsione per permettere gli aggiornamenti bulk dell'indice
	Public DisableRicorsione
    
	Private UrlRewritingAttivo
	Private UrlRewritingWebId
    
    '******************************************************************************************************************************************
    '******************************************************************************************************************************************
    'COSTRUTTORI CLASSE
    '******************************************************************************************************************************************
    
    Private Sub Class_Initialize()
    	OrdineLenght = 3
		DisableRicorsione = false
    	
    	Set conn = Server.CreateObject("ADODB.connection")
    	conn.open Application("DATA_ConnectionString")
    	
    	set content = new ObjContent
    	set content.conn = conn
    	set dizionario = request.form
		
		UrlRewritingAttivo = false
		UrlRewritingWebId = 0
    End Sub
    
    Private Sub Class_Terminate()
    	set content = nothing
    End Sub
    
    
    '******************************************************************************************************************************************
    '******************************************************************************************************************************************
    'FUNZIONI GENERICHE DELLA CLASSE
    '******************************************************************************************************************************************
    
	
	'funzinoe che restituisce lo stato di attivazione dell'url rewriting
	function IsUrlRewritingAttivo(conn, rs, WebId)
		
		if UrlRewritingWebId <> WebId then
			dim sql
			sql = "SELECT id_webs FROM tb_webs WHERE id_webs=" & cIntero(WebId) & " AND " & Sql_IsTrue(conn, "URL_rewriting_attivo")
			if cIntero(GetValueList(conn, rs, sql)) = cIntero(webId) then
				UrlRewritingAttivo = true
			else
				UrlRewritingAttivo = false
			end if
			UrlRewritingWebId = WebId
		end if
		
		IsUrlRewritingAttivo = UrlRewritingAttivo
	end function
	   
	   
        
    '.................................................................................................
    '..			restituisce la lista di ID dei discendenti data la lista delle tipologie
    '..			IdList				Lista dei nodi di partenza separati da ","
    '.................................................................................................
    Public Function DiscendentiID(IdList)
        DiscendentiID = ValueList(Discendenti(IdList, ""), "idx_id")
    end function
    
    
    '.................................................................................................
    '..			restituisce il recordset contenente la lista di voci discendenti della\e tipologie indicate
    '..			IdList				Lista dei nodi di partenza separati da ","
    '..			order				campi per l'ordinamento
    '.................................................................................................
    Public Function DiscendentiOrder(IdList, order)
        dim sql, Id, rs
        set rs = Server.CreateObject("ADODB.recordset")
        
        sql = "SELECT * FROM v_indice WHERE "
        if IdList<>"" then
            sql = sql + "("
            for each Id in split(IdList, ",")
                sql = sql + SQL_IdListSearch(conn, "idx_tipologie_padre_lista", Trim(Id)) + " OR "
            next
            sql = left(sql, len(sql) - 4) + ")"
        else
            sql = sql & "(1=0)"
        end if
		
		if order <> "" then
			sql = sql &" ORDER BY "& order
		end if
        rs.open sql, conn, adOpenStatic, adLockOptimistic
        
        set DiscendentiOrder = rs
    end function
	
	
	Public Function Discendenti(idList)
		set Discendenti = DiscendentiOrder(idList, "")
	End Function
	
    
    '.................................................................................................
    '..			restituisce gli ID delle foglie data la tipologia separati da ","
    '..			conn			aperta sul database
    '..			ID				ID della foglia padre
    '.................................................................................................
    Public Function FoglieID(ID)
    	dim sql, rs
    	'recupera informazioni su categoria attuale
    	sql = "SELECT idx_foglia FROM tb_contents_index WHERE idx_id="& cIntero(ID)
    	set rs = conn.execute(sql)
    	
    	'aggiunge le categorie a cui possono essere collegati i record: categorie foglie o 
    	'tutte le categorie nella modalita' mista
    	FoglieID = FoglieID & ID
    	if not rs("idx_foglia") then
    		'se non &egrave; foglia va a verificare i figli
    		sql = "SELECT idx_id FROM tb_contents_index WHERE idx_padre_id="& cIntero(ID)
    		set rs = conn.execute(sql, ,adCmdText)
    		if not rs.eof then
    			FoglieID = FoglieID & ", "			'aggiunge separatore per categoria padre solo in gestione mista
    		end if
    		while not rs.eof
    			FoglieID = FoglieID & FoglieID(rs("idx_id"))
    			rs.moveNext
    			if not rs.eof then
    				FoglieID = FoglieID & ", "
    			end if
    		wend
    	end if
    	Set rs = nothing
    End Function
    
    
    '.................................................................................................
    '..			ricalcola il flag foglia per tutti i rami dell'albero
    '.................................................................................................
    Public Sub RicalcolaFoglie()
    	dim sql
        
        sql = " UPDATE tb_contents_index SET idx_foglia = 1 " + _
              " WHERE idx_id NOT IN (SELECT idx_padre_id FROM tb_contents_index i WHERE idx_padre_id = tb_contents_index.idx_id)"
        CALL conn.Execute(sql, ,adExecuteNoRecords)
        
        sql = " UPDATE tb_contents_index SET idx_foglia = 0 " + _
              " WHERE idx_id IN (SELECT idx_padre_id FROM tb_contents_index i WHERE idx_padre_id = tb_contents_index.idx_id)"
        CALL conn.Execute(sql, ,adExecuteNoRecords)
    End Sub
    
    
    '.................................................................................................
    '..			restituisce la query sql che genera l'elenco delle categorie con relative sottocategorie
    '..			conn			aperta sul database
    '..			rs				oggetto recordset chiuso e creato
    '..			SoloFoglie		se true visualizza solo le categorie "foglie"
    '.................................................................................................
    Public Function QueryElenco(SoloFoglie, condition)
    	dim sql, level, WHERE_sql, rs
    	sql = "SELECT idx_livello FROM tb_contents_index GROUP BY idx_livello"
    	Set rs = server.CreateObject("ADODB.recordset")
    	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
    	
    	if SoloFoglie then
    		WHERE_sql = WHERE_sql &" AND "& SQL_IsTrue(conn, "TIP_L0.idx_foglia")
    	end if
    	sql = ""
    	while not rs.eof
    		sql = sql & "SELECT TIP_L0.idx_id, TIP_L0.idx_livello, TIP_C0.co_id, TIP_C0.co_titolo_it, TIP_C0.co_visibile, TIP_C0.co_F_key_id, TIP_C0.co_F_table_id, (" 
    		for level = rs("idx_livello") to 1 step -1 
    			sql = sql & "TIP_C" & level & ".co_titolo_it " & SQL_concat(conn) & " ' - ' " & SQL_concat(conn)
    		next
    		sql = sql & " TIP_C0.co_titolo_it) AS NAME"& _
    				    " FROM " & String(CIntero(rs("idx_livello")) * 2, "(") & " (tb_contents_index TIP_L0 "& _
    					" INNER JOIN tb_contents TIP_C0 ON TIP_L0.idx_content_id = TIP_C0.co_id)"
    		for level = 1 to rs("idx_livello")
    			sql = sql & " INNER JOIN tb_contents_index TIP_L" & level & " ON TIP_L" & (level-1) & ".idx_padre_id = TIP_L" & level & ".idx_id )"& _
    					    " INNER JOIN tb_contents TIP_C" & level & " ON TIP_L" & level & ".idx_content_id = TIP_C" & level & ".co_id )"
    		next
    		
    		'se sono in inserimento e ho limitazioni rispetto alle sezioni allora filtro
    		if NOT ChkPrm(prm_indice_trasparente, 0) then
    			sql = sql &" INNER JOIN rel_index_admin ON ((1=0"
    			for level = 0 to rs("idx_livello")
    				sql = sql &" OR TIP_L"& level &".idx_id = rel_index_admin.ria_index_id "
    			next
    			sql = sql &") AND ria_admin_id = "& session("ID_ADMIN") &")"
    		end if
    		
    		sql = sql & " WHERE TIP_L0.idx_livello=" & rs("idx_livello") & WHERE_sql
    		if condition <> "" then
    			sql = sql & " AND ( "
    			for level = 0 to rs("idx_livello")
    				if level > 0 then sql = sql & " OR "
    				sql = sql & " ( " & replace(condition, "tb_contents_index", "TIP_L" & level) & " ) "
    			next
    			sql = sql & " ) "
    		end if
    		rs.movenext
    		if not rs.eof then
    			sql = sql & " UNION "
    		end if
    	wend
    	if sql <> "" then
    		sql = sql & " ORDER BY " & IIF(rs.recordcount > 1, "NAME", "co_titolo_it")
    	else		'nessun record presente
    		sql = " SELECT *, (co_titolo_it) AS NAME FROM tb_contents_index TIP_L0"& _
    			  " INNER JOIN tb_contents TIP_C0 ON TIP_L0.idx_content_id = TIP_C0.co_id"& _
    			  " WHERE (1=1) " & WHERE_sql & " ORDER BY NAME"
    	end if
    	rs.close
    	QueryElenco = sql
    	Set rs = nothing
    End Function
    
    
    '.................................................................................................
    '..			restituisce il percorso ed il nome completo del nodo
    '..			idx_id 			id del nodo di cui recuperare il nome
    '.................................................................................................
    Public Function NomeCompleto(idx_id)
    	NomeCompleto = NomeCompletoByLanguage(idx_id, LINGUA_ITALIANO)
    End Function
	
	
	'.................................................................................................
    '..			restituisce il percorso ed il nome completo del nodo nella lingua richiesta
    '..			idx_id 			id del nodo di cui recuperare il nome
    '..			lingua			lingua in cui recuperare il nome
    '.................................................................................................
	Public Function NomeCompletoByLanguage(idx_id, Lingua)
		dim sql, rs, level
    	Set rs = server.CreateObject("ADODB.recordset")
    	
    	sql = "SELECT idx_tipologie_padre_lista FROM tb_contents_index WHERE idx_id = "& cIntero(idx_id)
    	sql = GetValueList(conn, rs, sql)
    	if sql = "" then
    		sql = 0
    	end if
    	sql = " SELECT " & SQL_MultiLanguage( "co_titolo_<LINGUA>", ", ") + _
			  " FROM v_indice "& _
    		  " WHERE idx_id IN ( " & sql & " )"& _
    		  " ORDER BY idx_livello"
    	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    	
		NomeCompletoByLanguage = ""
		
		while not rs.eof
			NomeCompletoByLanguage = NomeCompletoByLanguage & CBLL(rs, "co_titolo", Lingua)
			rs.movenext
			if not rs.eof then
				NomeCompletoByLanguage = NomeCompletoByLanguage + " - "
			end if
		wend
	
    	rs.close
    	Set rs = nothing
	end function
    
    
    '.................................................................................................
    '..			setta il dizionario dell'index con i parametri del content
    '.................................................................................................
    Public Sub SetIndexFromContent()
    	dim lingua
        if UCase(TypeName(dizionario)) = "RECORDSET" then
            for each lingua in Application("LINGUE")
    	    	dizionario("idx_link_url_" + lingua) = content.dizionario("co_link_url_" + lingua).value
        	next
        	dizionario("idx_link_pagina_id") = content.dizionario("co_link_pagina_id")
        	dizionario("idx_link_tipo") = content.dizionario("co_link_tipo")
        else
        	for each lingua in Application("LINGUE")
    	    	if content.dizionario.Exists("co_link_url_" + lingua) then
        	    	dizionario("idx_link_url_" + lingua) = content.dizionario("co_link_url_" + lingua)
        		end if
        	next
        	if content.dizionario.Exists("co_link_pagina_id") then
        		dizionario("idx_link_pagina_id") = content.dizionario("co_link_pagina_id")
        	end if
        	if content.dizionario.Exists("co_link_tipo") then
        		dizionario("idx_link_tipo") = content.dizionario("co_link_tipo")
        	end if
        end if
    End Sub
    
    
    '..................................................................................................
    '..		PER LA PARTE VISIBILE
    '..		ritorna il nome completo di percorso della tipologia corrente
    '..		conn			connessione aperta a database
    '..		rs				recordset creato e chiuso
    '..		tip_id			id della tipologia di cui reperire il percorso ed il nome
    '..................................................................................................
    Public Function NomeCompletoVisibile(tip_id)
    	dim sql, rs, level
    	Set rs = server.CreateObject("ADODB.recordset")
    	
    	sql = "SELECT idx_tipologie_padre_lista FROM tb_contents_index WHERE idx_id = "& cIntero(tip_id)
    	sql = " SELECT c.* FROM tb_contents c"& _
    		  " INNER JOIN tb_contents_index i ON c.co_id = i.idx_content_id"& _
    		  " WHERE idx_id IN ("& GetValueList(conn, rs, sql) &")"& _
    		  " ORDER BY idx_livello"
    	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    	
    	if not rs.eof then
    		NomeCompleto = CBL(rs, "co_titolo")
    		while not rs.eof
    			NomeCompleto = NomeCompleto &" - "& CBL(rs, "co_titolo")
    			rs.movenext
    		wend
    	end if
    	
    	rs.close
    	Set rs = nothing
    End Function
	
	
	'..................................................................................................
	'funzione che restituisce il link del nodo contenuto nel recordset
	'..................................................................................................
	Public Function GetNodeUrl(rsi, lingua)
		if CIntero(rsi("idx_link_tipo")) <> lnk_interno then
			GetNodeUrl = CBLL(rsi, "idx_link_url", lingua)
		elseif CBLL(rsi, "idx_link_url", lingua)<>"" then
			'recupera parte itnerna dell'url
			if IsUrlRewritingAttivo(rsi.ActiveConnection, NULL, rsi("idx_webs_id")) AND _
			   cString(rsi("idx_link_url_rw_" & lingua))<>"" then
			    'url con url rewriting
				GetNodeUrl = "/" & rsi("idx_link_url_rw_" & lingua)
			elseif cString(rsi("idx_link_url_" & lingua))<>"" then
				'url interno semplice
				GetNodeUrl = "/" & rsi("idx_link_url_" & lingua)
			else
				GetNodeUrl = "/" & CBLL(rsi, "idx_link_url", lingua)
			end if
		else
			GetNodeUrl = ""
		end if
	end function
	
	
	'..................................................................................................
	'funzione che restituisce il link completo anche con dominio del nodo contenuto nel recordset
	'..................................................................................................
	Public Function GetNodeCompleteUrl(rsi, lingua)
		GetNodeCompleteUrl = GetNodeUrl(rsi, lingua)
		
		if CIntero(rsi("idx_link_tipo")) = lnk_interno then
			GetNodeCompleteUrl = GetSiteUrl(rsi.ActiveConnection, rsi("idx_webs_id"), NULL) & GetNodeCompleteUrl
		end if
		
    end function
	

	'..................................................................................................
	'funzione che scrive la parte di html per linkare il nodo presente nel recordset
	'..................................................................................................
	Public Sub WriteNodeLink(rsi, htmlAttributes, lingua)
		CALL WriteNodeLabelLink(rsi, NomeCompleto(rsi("idx_id")), "", htmlAttributes, lingua)
	end sub
	
	
	'..................................................................................................
	'funzione che scrive la parte di html per linkare il nodo presente nel recordset 
	'con la personalizzazione del testo del link
	'..................................................................................................
	Public Sub WriteNodeLabelLink(rsi, linkLabel, linkParameters, htmlAttributes, lingua)
		dim url
		if not rsi.eof then 
			if cString(CBLL(rsi, "idx_link_url", lingua))<>"" then 		
				url = GetNodeCompleteUrl(rsi, lingua) %>
				<a target="_blank" href="<%= url + linkParameters %>" title="apri il link &ldquo;<%= url %>&rdquo; in una nuova finestra." <%= htmlAttributes %>>
					<%= linkLabel %> 
				</a>
			<% else %>
				<%= linkLabel %>
			<% end if
		end if
	end sub
	
    
    '.................................................................................................
    '..		scrive il sistema di input per la selezione di una categoria
    '..		conn			    aperta sul database
    '..		rs				    oggetto recordset chiuso e creato
    '..     VociSqlCondition    eventuale filtro sql per selezione delle voci dell'indice
    '..     TipiSqlCondition    eventuale filtro sql per selezioone dei tipi di voci
    '..		FormName		    Nome del form in cui viene generato l'input
    '..		InputName 		    Nome dell'input generato
    '..		InputValue		    Valore/categoria selezionata
	'..		NodeWebId			Eventuale id del sito a cui devono appartenere i nodi selezionabili.
    '..		OnlyLeaf		    Indica se vengono visualizzate solo le foglie (TRUE) o tutte le categorie
    '..		DisplayReduced	    Indica se viene visualizzato un input 
    '.						    ridotto (TRUE, per selezione nei motori di ricerca) o esteso (con link testuali)
    '..		disabled		    disabilita tutto
    '..     obbligatorio        indica se viene visualizzato il pulsante reset o il simbolo di "(*) obbligatorio"
    '.................................................................................................
    Public Sub WritePicker(VociSqlCondition, TipiSqlCondition, FormName, InputName, InputValue, NodeWebId, OnlyLeaf, DisplayReduced, InputSize, disabled, obbligatorio)
    	dim ViewName, ViewValue, rs
    	Set rs = server.CreateObject("ADODB.recordset")
    	if cInteger(InputValue)>0 then
    		'recupera valore dell'input
    		ViewValue = NomeCompleto(InputValue)
    	else
    		ViewValue = ""
    	end if
    	ViewName = "view_" & InputName
        
        'imposta variabile per passaggio query per filto degli elementi 
        session("CONDIZIONE_SELEZIONE_VOCI_" & FormName & "_" & InputName) = cString(VociSqlCondition)
        session("CONDIZIONE_SELEZIONE_TIPI_" & FormName & "_" & InputName) = cString(TipiSqlCondition)
        %>
		<input type="hidden" name="<%= InputName %>" id="<%= InputName %>" value="<%= InputValue %>" <%= Disable(disabled) %>>
    	<table cellpadding="0" cellspacing="0">
    		<tr>
    			<td <%= IIF(DisplayReduced, " colspan=""2"" ", " style=""padding-top:2px;"" ") %>>
    				<input READONLY type="text" <%= DisableClass(disabled, "") %> name="<%= ViewName %>" id="<%= ViewName %>" value="<%= ViewValue %>" style="padding-left:3px;" size="<%= InputSize %>" onmouseover="<%= FormName %>_<%= InputName %>_UpdateTitle(this)" onclick="if (!document.getElementById('<%= InputName %>').disabled) <%= FormName %>_<%= InputName %>_ApriFinestra()">
    			</td>
    		<% if DisplayReduced then %>
    			</tr>
    			<tr>
    		<% end if %>
    			<td style="<%= IIF(DisplayReduced, IIF(obbligatorio, "width:85%; ", "width:68%; ") + "padding-bottom:2px;", "padding-top:1px;") %>" nowrap><a <%= DisableClass(disabled, IIF(DisplayReduced, "button_input_bottom", "button_input")) %> id="link_scegli_<%= InputName %>" href="javascript:void(0)" onclick="<%= FormName %>.<%= ViewName %>.onclick();" title="Apre l'elenco delle voci dell'indice per selezionarne una." <%= ACTIVE_STATUS %>>SCEGLI</a></td>
                <% if not obbligatorio then %>
        			<td style="<%= IIF(DisplayReduced, "width:32%; padding-bottom:3px;", "padding-top:1px;") %>"><a <%= DisableClass(disabled, IIF(DisplayReduced, "button_input_bottom", "button_input")) %> id="link_reset_<%= InputName %>" style="border-left:0px;" href="javascript:void(0)" onclick="if (!document.getElementById('<%= InputName %>').disabled) {<%= FormName %>.<%= InputName %>.value='';<%= FormName %>.<%= viewName %>.value=''}" title="cancella la selezione eseguita" <%= ACTIVE_STATUS %>>RESET</a></td>
                <% else %>
                    <td style="<%= IIF(DisplayReduced, "width:15%;", "") %>">&nbsp;(*)</td>
                <% end if %>
    		</tr>
    	</table>
       
       <script language="JavaScript" type="text/javascript">
            function <%= FormName %>_<%= InputName %>_ApriFinestra(){
                OpenAutoPositionedScrollWindow('<%= GetAmministrazionePath() %>library/IndexContent/IndexSeleziona.asp?formname=<%= FormName %>&inputname=<%= InputName %>&<%= IIF(OnlyLeaf, "SoloFoglie=1&", "") %><%= IIF(cIntero(NodeWebId)>0, "WebIdFilter=" & NodeWebId & "&", "") %>selected=' + <%= FormName %>.<%= InputName %>.value, 'selezione_voce', 760, 450, true)
                <%= FormName %>_<%= InputName %>_UpdateTitle(<%= FormName %>.<%= InputName %>)
            }
            
            function <%= FormName %>_<%= InputName %>_UpdateTitle(viewInput){
                viewInput.title = viewInput.value;
            }
    
            <%= FormName %>_<%= InputName %>_UpdateTitle(<%= FormName %>.<%= InputName %>)
        </script>
    <%
    End Sub
    
    
    '.................................................................................................
    '..		scrive la riga per inserire il pulsante "indicizza"
    '..		N.B.: funzione collegata in ClassContent.Associazioni() che tramite JS aggiorna lo stato del bottone (indicizza / indicizzato)
    '.................................................................................................
    Public Sub WriteButton(tableName, co_F_key_id, posizione) 
	
		dim sql, rsTab, co_id, val, newsletter
		
		sql = " SELECT * FROM tb_siti_tabelle " + _
			  " WHERE tab_name LIKE '" & ParseSql(tableName, adChar) & "' " + _
			  " ORDER BY tab_priorita_base DESC "
		set rsTab = conn.execute(sql)
	
		'Giacomo - 24/10/2012 -------------------------
		'modifica per correggere bug: ora controllo tutte le tabelle riguardanti il contenuto, prima di decidere come "disegnare" il pulsante
		dim sqlVer, filtroTabId
		filtroTabId = 0
		do while (not rsTab.EOF)
			sqlVer = "SELECT "&rsTab("tab_id")&" FROM "&rsTab("tab_from_sql")
			if inStr(sqlVer, "WHERE") > 0 then
				sqlVer = sqlVer & " AND "&rsTab("tab_field_chiave")&"="&co_F_key_id
			else
				sqlVer = sqlVer & " WHERE "&rsTab("tab_field_chiave")&"="&co_F_key_id
			end if	
			filtroTabId = cIntero(GetValueList(conn, NULL, sqlVer))
			if filtroTabId > 0 then
				exit do
			end if
			rsTab.moveNext
		loop
		rsTab.close
		sql = " SELECT * FROM tb_siti_tabelle " & _
			  " WHERE tab_id = " & filtroTabId & _
			  " ORDER BY tab_priorita_base DESC "
		set rsTab = conn.execute(sql)
		'----------------------------------------------

		if not rsTab.eof then
		
			sql = " SELECT co_id FROM tb_contents WHERE co_F_key_id = "& cIntero(co_F_key_id) & _
													  " AND co_F_table_id = "& CIntero(rsTab("tab_id"))
			co_id = CIntero(GetValueList(conn, NULL, sql))		

			if index.content.ChkPrm(co_id) then
				sql = " SELECT COUNT(*) FROM tb_contents_index WHERE idx_content_id = "& co_id
					  
					  'modificato il 06/09/2011 per ottimizzare
					  'IN (SELECT tab_id FROM tb_siti_tabelle WHERE tab_name LIKE '" & ParseSql(tableName, adChar) & "')"
    			val = CIntero(GetValueList(conn, NULL, sql))
				
                Select case posizione 
       				CASE POS_TESTATA %>
       					<!-- </td> Giacomo, commentato il 04/04/2013 -->
       				<% CASE POS_ELENCO %>
       					&nbsp;
       				<% CASE else
       			end select
				
				if TagAbilitati(conn) AND rsTab("tab_tags_abilitati") then
					Select case posizione 
	    				CASE POS_TESTATA %>
	    					<td style="padding-left:3px;" nowrap>
	    				<% CASE else
	    			end select
					%>
					<a class="button<%= IIF(posizione<>POS_INDICE, "", "_L2") %>" 
					   href="javascript:void(0)"
					   onclick="OpenAutoPositionedScrollWindow('<%= GetLibraryPath() %>IndexContent/Tagga.asp?co_F_table_id=<%= rsTab("tab_id") %>&co_F_key_id=<%= co_F_key_id %>', '_blank', 800, 450, true)"
       			   	   title="Tagga il record" <%= ACTIVE_STATUS %>
       			   	   name="tagga_<%= co_F_key_id %>"
       			   	   id="tagga_<%= co_F_key_id %>">TAGS</a>
					<%
	                Select case posizione 
	       				CASE POS_TESTATA %>
	       					</td>
	       				<% CASE POS_ELENCO %>
	       					&nbsp;
	       				<% CASE else
	       			end select
				end if
  			  
    			Select case posizione 
    				CASE POS_TESTATA %>
    					<td style="padding-left:3px;" nowrap>
    				<% CASE else
    			end select 
				
				sql = "SELECT COUNT(*) FROM tb_newsletters WHERE " & SQL_IsTrue(conn, "nl_gestione_dinamica_contenuti")
				newsletter = cBoolean(cIntero(GetValueList(conn, NULL, sql))>0, false)
				%>
                <% if content.IsRaggruppamento(tableName) AND posizione<>POS_INDICE then %>
                    <a class="button_disabled Indicizzato"
        			   href="javascript:void(0)" style="white-space:nowrap;"
        			   title="Raggruppamento" <%= ACTIVE_STATUS %>
        			   name="indicizza_<%= co_F_key_id %>"
        			   id="indicizza_<%= co_F_key_id %>"><%=IIF(newsletter, "INDICE", "COLLEGATO ALL'INDICE")%></a>
                <% elseif not content.IsRaggruppamento(tableName) OR posizione<>POS_INDICE then %>
        			<a class="button<%= IIF(posizione<>POS_INDICE, IIF(val = 0, " DaIndicizzare", " Indicizzato"), "_L2") %>"
        			   href="javascript:void(0)"
        			   onclick="OpenAutoPositionedScrollWindow('<%= GetLibraryPath() %>IndexContent/Indicizza.asp?co_F_table_id=<%= rsTab("tab_id") %>&co_F_key_id=<%= co_F_key_id %>', '_blank', 800, 450, true)"
        			   title="<%= IIF(val = 0, "Indicizza il record nel sito", "Record gi&agrave; indicizzato") %>" <%= ACTIVE_STATUS %>
        			   name="indicizza_<%= co_F_key_id %>"
        			   id="indicizza_<%= co_F_key_id %>"
					   ><%= IIF(posizione<>POS_INDICE, IIF(val = 0, IIF(newsletter, "INDICE", "COLLEGA ALL'INDICE"), IIF(newsletter, "INDICE", "COLLEGATO ALL'INDICE")), "CONTENUTO") %></a>
       			<%end if
				
				
				'SEZIONE PULSANTE NEWSLETTER
				if newsletter then
					Select case posizione 
						CASE POS_TESTATA %>
							</td><td style="padding-left:3px;" nowrap>
						<% CASE POS_ELENCO %>
							&nbsp;
						<% CASE else
					end select
					
					dim pulsante_attivo
					sql = " SELECT nlc_id FROM tb_newsletters_contents INNER JOIN tb_contents ON tb_newsletters_contents.nlc_co_id = tb_contents.co_id " & _
						  " WHERE co_f_table_id=" & rsTab("tab_id") & " AND co_f_key_id=" & co_F_key_id & " AND ISNULL(nlc_data_invio, 0)=0 "
					pulsante_attivo = cString(GetValueList(conn, NULL, sql))
					%>
					<a class="button<%=IIF(pulsante_attivo<>"", " newsletter", "")%>"
					   href="javascript:void(0)" style="white-space:nowrap;"
					   title="Gestione newsletter" <%= ACTIVE_STATUS %>
					   name="newsletter_<%= co_F_key_id %>"
					   id="newsletter_<%= co_F_key_id %>"
					   onclick="OpenAutoPositionedScrollWindow('<%= GetLibraryPath() %>IndexContent/Newsletter.asp?co_F_table_id=<%= rsTab("tab_id") %>&co_F_key_id=<%= co_F_key_id %>', '_blank', 650, 250, true)"
					   >NEWSLETTER</a> 
					<%
				end if
			end if
			
    	end if
		
		rsTab.close
		set rsTab = nothing
    End Sub
    
    	
    '.................................................................................................
    '..		scrive il pulsante per l'apertura della finestra di cancellazione della voce dell'indice
    '.................................................................................................
    Public Sub WriteDeleteButton(CssClass, idx_id) %>
        <a class="button<%= CssClass %>" href="javascript:void(0);" 
           onclick="OpenAutoPositionedScrollWindow('<%= GetLibraryPath() %>IndexContent/DeleteIndexVoce.asp?ID=<%= idx_id %>', 'delete', 500, 300, false);">
            CANCELLA
        </a>
    <% end sub
    
	'.................................................................................................
    '..		scrive il pulsante per l'apertura della finestra di modifica del collegamento
    '.................................................................................................
	Public Sub WriteCollegamentoButton(CssClass, co_F_table_id, co_F_key_id, idx_id) %>
		<a class="button<%= CssClass %>" href="javascript:void(0);" 
           onclick="OpenAutoPositionedScrollWindow('<%= GetLibraryPath() %>IndexContent/IndicizzaAssocia.asp?co_F_table_id=<%= co_F_table_id %>&co_F_key_id=<%= co_F_key_id %>&ID=<%= idx_id %>', '_blank', 800, 450, true)">
			MODIFICA VOCE
		</a>
		<%
	end sub
    
    '.................................................................................................
    '..		restituisce l'ID della tabella (tb_siti_tabelle) dato il nome
    '.................................................................................................
    Public Function GetTable(nome)
    	GetTable = CIntero(GetValueList(conn, NULL, "SELECT tab_id FROM tb_siti_tabelle WHERE tab_name LIKE '"& ParseSql(nome, adChar) &"'"))
    End Function
    
	'.................................................................................................
    '..		restituisce il nome della tabella (tb_siti_tabelle) dato l'id
    '.................................................................................................
    Public Function GetTableName(id)
		GetTableName = GetValueList(conn, NULL, "SELECT tab_name FROM tb_siti_tabelle WHERE tab_id=" & id)
	end function
    
    
    '.................................................................................................
    '..		restituisce l'ID della tabella (tb_siti_tabelle) dati il nome e il titolo
    '.................................................................................................
    Public Function GetTableNT(nome, titolo)
    	GetTableNT = CIntero(GetValueList(conn, NULL, " SELECT tab_id FROM tb_siti_tabelle"& _
    												  " WHERE tab_titolo LIKE '"& ParseSql(titolo, adChar) &"' AND tab_name LIKE '"& ParseSql(nome, adChar) &"'"))
    End Function
    
    
    '.................................................................................................
    '..		restituisce la lista delle pubblicazioni automatiche che bloccano il record
    '.................................................................................................
    Public Function GetPubblicazioniLockers(idx_id)
        dim sql
        sql = " SELECT pub_titolo FROM tb_siti_tabelle_pubblicazioni " + _
              " WHERE pub_id IN (SELECT rip_pub_id FROM rel_index_pubblicazioni WHERE rip_idx_id=" & cIntero(idx_id) & ") "
        GetPubblicazioniLockers = GetValueList(conn, NULL, sql)
    end function
	
	
	'.................................................................................................
    '..		restituisce la lista delle pubblicazioni automatiche che bloccano il record
    '.................................................................................................
    Public Function IsPubblicatoPrincipale(idx_id)
		IsPubblicatoPrincipale = false
		
		if CIntero(idx_id) > 0 then
	        dim rs, sql
			set rs = server.createobject("adodb.recordset")
	        sql = " SELECT * FROM (((tb_siti_tabelle t" + _
				  " INNER JOIN tb_siti_tabelle_pubblicazioni p ON t.tab_id = p.pub_tabella_id)" + _
				  " INNER JOIN rel_index_pubblicazioni r ON p.pub_id = r.rip_pub_id)" + _
				  " INNER JOIN tb_contents_index i ON r.rip_idx_id = i.idx_id)" + _
				  " INNER JOIN tb_contents c ON i.idx_content_id = c.co_id" + _
				  " WHERE rip_idx_id = "& cIntero(idx_id)
			rs.open sql, conn, adOpenStatic, adLockOptimistic
			
			do while not rs.eof
				if CString(rs("pub_field_principale")) <> "" then
					if CBoolean(rs("pub_field_principale"), false) then
						IsPubblicatoPrincipale = true
						exit do
					else
						sql = " SELECT ("& rs("pub_field_principale") &") FROM " & rs("tab_from_sql") & SQL_AddOperator(rs("tab_from_sql"), "AND") & rs("tab_field_chiave") &" = "& cIntero(rs("co_F_key_id"))
						if CBoolean(GetValueList(conn, NULL, sql), false) then
							IsPubblicatoPrincipale = true
							exit do
						end if
					end if
				end if
				
				rs.movenext
			loop
			
			rs.close
	        set rs = nothing
		end if
    end function
    
    
	
    
    '******************************************************************************************************************************************
    '******************************************************************************************************************************************
    'GESTIONE INSERIMENTO NUOVA CATEGORIA
    '******************************************************************************************************************************************
    
    'scrive il form per l'inserimento dati (inserimento/modifica).
    Public Sub Modifica(ID)
    	dim label, labelArticolo, sql, rs, rsi,rsv,sqlv
		dim i, LockIcon, value
		
		set rs = server.CreateObject("ADODB.Recordset")
    	
    	if request("co_F_table_id") <> "" then		'vengo dagli applicativi
    		label = "collegamento"
    		labelArticolo = " collegamento"
    	else										'vengo dall'indice
    		label = "voce"
    		labelArticolo = "la voce"
    	end if
    	
    	ID = CIntero(ID)
		
    	if ID > 0 AND request.servervariables("REQUEST_METHOD") <> "POST" then
    		set dizionario = conn.execute("SELECT * FROM tb_contents_index WHERE idx_id = "& cIntero(ID))
			
			if not dizionario.eof then
				sql = "SELECT * FROM v_indice INNER JOIN tb_siti ON v_indice.tab_sito_id = tb_siti.id_sito WHERE idx_id=" & cIntero(ID)
				set rsi = conn.execute(sql)
				
				if dizionario("idx_autopubblicato") then
					LockIcon = "<img src=""" & GetAmministrazionePath() & "grafica/padlock.gif"" style=""border:2px solid #f4F4F4; margin-left:2px; margin-right:2px;""" + _
							   "alt=""Valore modificabile esclusivamente dall'applicativo &ldquo;" + GetApplicationShortName(rsi("sito_nome")) + _
							   "&rdquo; dalla sezione che gestisce i contenuti di tipo &ldquo;" & rsi("tab_titolo") & "&rdquo;."">"
				end if
			end if
    	end if 
		
		
		'calcolo il content ID
		if CIntero(dizionario("idx_content_id")) = 0 AND _
		   CIntero(request("co_F_table_id")) > 0 AND CIntero(request("co_F_key_id")) > 0 then
			contentID = content.GetId(request("co_F_table_id"), request("co_F_key_id"))
		else
   			contentID = dizionario("idx_content_id")
   		end if
		
		set rsv = server.CreateObject("ADODB.Recordset")
		sqlv= " SELECT * FROM (tb_siti_tabelle_pubblicazioni INNER JOIN rel_index_pubblicazioni " & _
			  " ON tb_siti_tabelle_pubblicazioni.pub_id=rel_index_pubblicazioni.rip_pub_id) " & _
			  " INNER JOIN tb_contents_index ON tb_contents_index.idx_id=rel_index_pubblicazioni.rip_idx_id " & _
			  " WHERE idx_autopubblicato=1 AND idx_id=" &cIntero(ID)
		'response.write sqlv
		rsv.open sqlv,conn

		%>
   		<script type="text/javascript">
   			// abilita i campi disabilitati in caso di contenuto di tipo pagina
   			function Submit() {
               <% 	if not dizionario("idx_autopubblicato") then %>
       				var o
       				o = document.getElementById("idx_link_tipo_<%= lnk_interno %>")
   					if (o)
       					o.disabled = !o.checked
       				document.getElementById("idx_link_pagina_id").disabled = false	
       				
       				o = document.getElementById("idx_link_tipo_<%= lnk_esterno %>")
   					if (o)
       					o.disabled = !o.checked
       				<% 	for each i in Application("LINGUE") %>
       				o = document.getElementById("idx_link_url_<%= i %>")
       				if (o)
       					o.disabled = false
       				<% 	next
               	end if %>
   			}
   		</script>

   		<form action="" method="post" id="form1" name="form1" onsubmit="Submit()">
		<% CALL WriteAdminIndexDataViewMode(false) %>
   		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
   			<caption>
   				<% 	if ID = 0 then %>
   				    Inserimento nuov<%= IIF(label = "collegamento", "o", "a") %>&nbsp;<%= label %>
   				<%	else %>
       				Modifica&nbsp;<%= label %>
   				<%	end if%>
   			</caption>
   			<tr>
                   <th colspan="4">DATI DEL<%= UCase(labelArticolo) %></th>
            </tr>
   			<tr>
   				<td class="label_no_width" nowrap><%= IIF(label = "collegamento", "collegato a", "voce collegata a") %>:</td>
   				<td class="content" colspan="3">
					<% if dizionario("idx_autopubblicato") then %>	
						<input type="hidden" name="idx_padre_id" value="<%= dizionario("idx_padre_id") %>">
						<%= LockIcon %>
						<%'recupera dati del nodo padre
						sql = "SELECT * FROM v_indice WHERE idx_id=" & cIntero(dizionario("idx_padre_id"))
						rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText 
						if not rs.eof then %>
							<% CALL WriteNodeLink(rs, "", LINGUA_ITALIANO) %>
						<% else %>
							Radice di base.
						<% end if
						rs.close
					else
					 	'sottocategoria creabile in tutte le categorie, anche con record associati
   						CALL index.WritePicker("", "", "form1", "idx_padre_id", _
   							IIF(CIntero(dizionario("idx_padre_id")) > 0, dizionario("idx_padre_id"), request.querystring("idx_padre_id")), _
   							0, false, false, "87", false, true) 
					end if%>
   				</td>
   			</tr>
   			<tr>
   				<td class="label">contenuto:</td>
   				<td colspan="3">
	   				<table cellpadding="0" cellspacing="0" width="100%">
	   					<tr>
							<% if dizionario("idx_autopubblicato") OR request("co_F_table_id") <> "" then
								if NOT IsObject(rsi) then
									sql = " SELECT * FROM tb_contents INNER JOIN tb_siti_tabelle " + _
										  " ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id WHERE co_id= " & content.GetID(cIntero(request("co_F_table_id")), cIntero(request("co_F_key_id"))) 
									set rsi = conn.execute(sql)
								end if %>
								<input type="hidden" name="idx_content_id" value="<%= rsi("co_id") %>">
	   							<td class="content">
									<%= LockIcon %>
									<%= rsi("co_titolo_it") %>&nbsp;<% content.WriteTipoRS(rsi) %>
								</td>
								<%if request("co_F_table_id") = "" then %>
									<td class="content_right" style="width=150px;">
										<a class="button_block" style="padding-top:9px; padding-bottom:9px; width=140px;" 
							   			   href="javascript:void(0);"
										   onclick="OpenAutoPositionedScrollWindow('<%= GetAmministrazionePath() %>library/IndexContent/ContentGestione.asp?FROM=indice&co_F_key_id=<%= rsi("co_F_key_id") %>&co_F_table_id=<%= rsi("co_F_table_id") %>&ID=<%= rsi("co_id") %>', 'contenuto', 760, 450, true);" title="Modifica e completa i dati del contenuto." <%= ACTIVE_STATUS %>>
											COMPLETA I DATI
										</a>
									</td>
								<% end if
							else %>
								<td class="content">
	   								<% 	CALL content.WritePicker("form1", _
	                                                                IIF(request("co_F_table_id") <> "", request("co_F_table_id"), GetTable(tabRaggruppamentoTable)), _
	                                                                IIF(request("co_F_key_id"), request("co_F_key_id"), 0), _
	                                                                "idx_content_id", _
	                                                                dizionario("idx_content_id"), _
	                                                                false, "76", false) %>
	   							</td>
	   							<td class="content">&nbsp;(*)</td>
	   						<% end if %>
						</tr>
	   				</table>
   				</td>
   			</tr>
			<tr>
				<td class="label">id:</td>
				<td class="content"><%= dizionario("idx_id") %></td>
				<td class="label">livello:</td>
				<td class="content">
					<% if dizionario("idx_livello")=0 then %>
						voce base
					<% else %>
						voce livello <%= dizionario("idx_livello") %>
					<% end if %>
				</td>
			</tr>
		</table>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<tr><th colspan="4" title="<%= title %>">LINK DEL<%= UCase(labelArticolo) %></th></tr>
			<% dim IsPrincipale
			IsPrincipale = IsPubblicatoPrincipale(ID) 
			if IsPrincipale then %>
				<input type="hidden" name="idx_principale" value="1">
			<% else
				if cIntero(dizionario("idx_id"))>0 then
					'in modifica: recupera valore dal record
					value = CBoolean(dizionario("idx_principale"), false)
				else
					'in inserimento: imposta il collegamento come "principale"
					if dizionario("idx_principale") = "" then
						
						'se è il primo collegamento lo mette principale, altrimenti lo mette alternativo
						sql = "SELECT COUNT(*) FROM v_indice WHERE co_id =" & cIntero(contentID) & " AND " & SQL_IsTrue(conn, "idx_principale")
						if cIntero(GetValueList(Conn, NULL, sql))>0 then
							'esiste già un collegamento principale
							value = false
						else
							'non esiste alcun collegamento principale
							value = true
						end if
					else
						value = CBoolean(dizionario("idx_principale"), false)
					end if
				end if
			end if %>
			<tr>
   				<td class="label" rowspan="2">tipo di link:</td>
   				<td class="content" style="width:5%;">
					<input type="radio" class="checkbox" <% if IsPrincipale then %> name="principale" disabled checked <% else %> value="1" name="idx_principale" <%= Chk(value) %><% end if %>>
				</td>
				<td class="content" style="width:8%;">
					principale
   				</td>
				<td class="content notes" rowspan="2">
					Impostando il link come &ldquo;principale&rdquo; questo diverr&agrave; l'url effettivo al quale la navigazione di tutti i collegamenti alternativi faranno riferimento.
				</td>
   			</tr>
			<tr>
				<td class="content">
					<input type="radio" class="checkbox" <% if IsPrincipale then %> name="nonprincipale" disabled <% else %> value="0" name="idx_principale" <%= Chk(not value) %><% end if %>>
				</td>
				<td class="content" style="width:8%;">
					alternativo
				</td>
			</tr>
   			<% if CIntero(request("co_F_table_id")) = GetTable(tabPagineTable) AND CIntero(request("co_F_key_id")) > 0 then 
				'visualizzazione delle pagine 
				%>
	   			<input type="hidden" name="idx_link_tipo" value="<%= lnk_interno %>">
   				<input type="hidden" name="idx_link_pagina_id" value="<%= request("co_F_key_id") %>">
   			<% elseif dizionario("idx_autopubblicato") then %>
				<input type="hidden" name="idx_link_tipo" value="<%= dizionario("idx_link_tipo") %>">
   				<input type="hidden" name="idx_link_pagina_id" value="<%= dizionario("idx_link_pagina_id") %>">
				<%	for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE")) %>
					<tr>
						<% if i = 0 then %>
					         <td class="label_no_width" rowspan="<%= ubound(Application("LINGUE"))+1 %>">link:</td>
			            <% end if %>
						<td class="content" colspan="3" title="<%=rsi("idx_link_url_" + Application("LINGUE")(i))%>">
							<table cellpadding="0" cellspacing="0">
								<tr>
									<td><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
									<td><%= LockIcon %></td>
									<td><% CALL WriteNodeLink(rsi, "", Application("LINGUE")(i)) %></td>
								</tr>
							</table>
						</td>
					</tr>
				<% next
			else
   				dim disabled, contentID, linkEsterno
   				
				'precalcolo il link
   				if CIntero(contentID) > 0 then
                	sql = " SELECT * FROM tb_contents c"& _
   						  " INNER JOIN tb_siti_tabelle t ON c.co_F_table_id = t.tab_id"& _
						  " WHERE co_id = "& cIntero(contentID)
					rs.open sql, conn, adOpenStatic, adLockOptimistic
   					disabled = content.LinkPrecalcola(rs, CString(dizionario("idx_link_tipo")) = "")
   					if disabled then
   						disabled = " disabled "
   					else
   						disabled = ""
   					end if
   				else
    				set rs = Server.CreateObject("Scripting.Dictionary")
   					rs("co_link_tipo") = ""
   					rs("co_link_pagina_id") = ""
   					for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
   						rs("co_link_url_"& Application("LINGUE")(i)) = ""
   					next
   					disabled = ""
   				end if
   				
   				'calcolo il tipo
   				linkEsterno = CIntero(rs("co_link_tipo")) = lnk_esterno OR CIntero(dizionario("idx_link_tipo")) = lnk_esterno
					
				dim title
				title = "url della voce:" + vbCrLF
				for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
					title = title + " url in lingua " + LINGUE_NOMI(i) & ":" & vbCrLf & _
									"interno: " & dizionario("idx_link_url_"& Application("LINGUE")(i)) & vbCrLF & _
									"di navigazione:" & dizionario("idx_link_url_rw_"& Application("LINGUE")(i)) & vbCrLF
  				next %>
   				<tr>
   					<td class="label" rowspan="4" style="width: 17%;">link a:</td>
   					<td class="content" rowspan="2" style="width:5%;">
	   					<input type="radio" class="noBorder" name="idx_link_tipo" id="idx_link_tipo_<%= lnk_interno %>"<%= IIF(dizionario("idx_autopubblicato"), " disabled ", " onclick=""Abilita(this.value)""") %>"
	   						   value="<%= lnk_interno %>" <%= disabled %> <%= Chk(NOT linkEsterno) %>>
	   				</td>
	   				<td class="content" colspan="2">pagina interna</td>
	   			</tr>
	   			<tr>
	   				<td class="content" colspan="2">
						<script type="text/javascript">
							function Abilita(v) {
								var o
								o = document.getElementById("idx_link_pagina_id");
								DisableControl(o, (v == "<%= lnk_esterno %>"))
								
								<%	for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
								//o = document.getElementById("idx_link_url_<%= Application("LINGUE")(i) %>");
								o = form1.idx_link_url_<%= Application("LINGUE")(i) %>;
								DisableControl(o, (v == "<%= lnk_interno %>"));
								<%  next %>
							}
						</script>
		   				<% dim pagina_id
						pagina_id = IIF(CIntero(dizionario("idx_link_pagina_id")) > 0, dizionario("idx_link_pagina_id"), rs("co_link_pagina_id"))
						if cIntero(pagina_id) = 0 and cIntero(request("tab_pagina_default_id")) > 0 then
							pagina_id = cIntero(request("tab_pagina_default_id"))
						end if
						CALL DropDownPages(conn, "form1", "430", 0, "idx_link_pagina_id", pagina_id, false, false)
		   					if disabled <> "" then %>
		   						<script language="javascript" type="text/javascript">
		   							DisableControl(document.getElementById("idx_link_pagina_id"), true)
		   						</script>
		   				<% 	end if %>
	   				</td>
	   			</tr>
	   			<tr>
	   				<td class="content_center" rowspan="2">
	   					<input type="radio" class="noBorder" name="idx_link_tipo" id="idx_link_tipo_<%= lnk_esterno %>" onclick="Abilita(this.value)" value="<%= lnk_esterno %>" <%= disabled %> <%= Chk(linkEsterno) %>>
	   				</td>
	   				<td class="content" colspan="2">pagina esterna</td>
	   			</tr>
	   			<tr>
	   				<td class="content" colspan="2">
	   					<table cellpadding="0" cellspacing="0">
	   						<%	for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE")) %>
		   						<tr>
		   							<td class="content" colspan="3">
		   								<img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
		   								<input type="text" <%= DisableClass(disabled <> "", "text") %> name="idx_link_url_<%= Application("LINGUE")(i) %>" maxlength="255" size="75"
		   									   value="<%= IIF(linkEsterno, IIF(cString(dizionario("idx_link_url_"& Application("LINGUE")(i)))<>"", dizionario("idx_link_url_"& Application("LINGUE")(i)), rs("co_link_url_"& Application("LINGUE")(i))), "") %>">
		   							</td>
		   						</tr>
	   						<%	next%>
	   						<%	if disabled = "" then %>
	   						<script type="text/javascript">
	   							if (document.getElementById("idx_link_tipo_<%= lnk_interno %>").checked)
	   								Abilita(<%= lnk_interno %>)
	   							else
	   								Abilita(<%= lnk_esterno %>)
	   						</script>
	   						<%	end if %>
	   					</table>
	   				</td>
	   			</tr>
       			<% if UCase(TypeName(rs)) = "RECORDSET" then
       			    rs.cancelUpdate
       				rs.close
       			end if
       		end if 
			
			if cIntero(dizionario("idx_id"))>0 then %>
				<tr <%= AdminIndexDataViewMode_CssStyle(true) %>>
					<td class="content">&nbsp;</td>
					<td class="content_right notes" colspan="3">
						<a class="button_L2_block" style="float:right; width:120px;" href="javascript:void(0);" 
						   onclick="OpenAutoPositionedScrollWindow('<%= GetAmministrazionePath() %>library/IndexContent/IndexRedirect.asp?FROM=indice&co_F_key_id=<%= rsi("co_F_key_id") %>&co_F_table_id=<%= rsi("co_F_table_id") %>&IDX=<%= ID %>', 'linkalternativi', 760, 450, true);" 
						   title="Modifica e visualizza gli indirizzi alternativi del nodo." <%= ACTIVE_STATUS %>>
							INDIRIZZI ALTERNATIVI
						</a>
						&Egrave; possibile impostare altri url alternativi "esterni" ai link generati automaticamente dal sistema.
						<% sql = "SELECT COUNT(*) FROM rel_index_url_redirect WHERE riu_idx_id=" & cIntero(dizionario("idx_id"))
						value = cIntero(GetValueList(Conn, rs, sql))
						if value > 0 then%>
							<br>
							Presenti n&ordm; <%= value %> url alternativi per la voce.
						<% end if %>
					</td>
				</tr>
				
				<tr <%= AdminIndexDataViewMode_CssStyle(true) %>>
					<td class="content">&nbsp;</td>
					<td class="content_right notes" colspan="3">
						<a class="button_L2_block" style="float:right; width:160px;" href="javascript:void(0);" 
						   onclick="OpenAutoPositionedScrollWindow('<%= GetAmministrazionePath() %>library/IndexContent/ImportaIndirizziAlternativi.asp?FROM=indice&co_F_key_id=<%= rsi("co_F_key_id") %>&co_F_table_id=<%= rsi("co_F_table_id") %>&IDX=<%= ID %>', 'importlinkalternativi', 760, 450, true);" 
						   title="Modifica e visualizza gli indirizzi alternativi del nodo." <%= ACTIVE_STATUS %>>
							IMPORT INDIRIZZI ALTERNATIVI
						</a>
						Import url alternativi (partendo da file excel esportato da google webmaster tools).
					</td>
				</tr>				
			<% end if %>
			<tr <%= AdminIndexDataViewMode_CssStyle(true) %>>
				<td class="content">&nbsp;</td>
				<td class="content_right notes" colspan="3">
					<% CALL WriteCopiaIndirizziAlternativi(request("ID"),"Se vuoi spostare gli URL di questa voce dell'indice su un'altra, clicca sul pulsante a destra.","GESTIONE URL") %>
				</td>
			</tr>
			</table>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
   			<tr><th colspan="4">INFORMAZIONI ALTERNATIVE DI PUBBLICAZIONE</th></tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr <%= AdminIndexDataViewMode_CssStyle(true) %> >
				<% 	if i = 0 then %>
					<td class="label_no_width" colspan="2" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo voce:</td>
				<% 	end if %>
					<td class="content" colspan="3">
						<table width="100%" cellpadding="0" cellspacing="0">
							<tr>
								<td width="4%" valign="top"><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<% 
								if not rsv.eof then
									if CString(rsv("pub_field_titolo_" + Application("LINGUE")(i))) <> "" then
										'campo sincronizzato: non modificabile 
										%>
										<td width="1%" valign="top"><%= LockIcon %></td>
										<td class="content_disabled">
											<%= TextHtmlEncode(dizionario("idx_titolo_"& Application("LINGUE")(i))) %>
										</td>									
									<% 
									else %>
										<td class="content">
											<input type="text" class="text" name="idx_titolo_<%= Application("LINGUE")(i) %>" value="<%= dizionario("idx_titolo_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:95%;">
										</td>
									<%
									end if
								else %>
								<td class="content">
									<input type="text" class="text" name="idx_titolo_<%= Application("LINGUE")(i) %>" value="<%= dizionario("idx_titolo_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:95%;">
								</td>
								<% end if %>
							</tr>
						</table>
					</td>
					
				</tr>
			<% next %>
			<tr>
   				<td class="label_no_width" colspan="2" style="width:17%;">ordine:</td>
   				<td class="content" colspan="2">
   					<input type="text" name="idx_ordine" value="<%= dizionario("idx_ordine") %>" maxlength="<%= index.OrdineLenght %>" size="3" class="text">
   				</td>
   			</tr>
   			<tr <%= AdminIndexDataViewMode_CssStyle(true) %>>
   				<td class="label_no_width" rowspan="2" style="width:6%;">immagini</td>
   				<td class="label_no_width" style="width:7%;">thumbnail:</td>
				<% if not rsv.eof then
						if CString(rsv("pub_field_foto_thumb")) <> "" then
							'campo sincronizzato: non modificabile 
							%>
							<td width="1%" valign="top"><%= LockIcon %></td>
							<td class="content_disabled">
								<%= TextHtmlEncode(dizionario("idx_foto_thumb")) %>
							</td>
						<%else %>
						<td class="content" colspan="2">
							<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "idx_foto_thumb", dizionario("idx_foto_thumb"), "", false) %>
						</td>
				<% 		end if
				else %>
					<td class="content" colspan="2">
						<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "idx_foto_thumb", dizionario("idx_foto_thumb"), "", false) %>
					</td>
				<% end if %>				   				
   			</tr>
   			<tr <%= AdminIndexDataViewMode_CssStyle(true) %>>
   				<td class="label_no_width">zoom:</td>
   				<% if not rsv.eof then
						if CString(rsv("pub_field_foto_zoom")) <> "" then
							'campo sincronizzato: non modificabile 
							%>
							<td width="1%" valign="top"><%= LockIcon %></td>
							<td class="content_disabled">
								<%= TextHtmlEncode(dizionario("idx_foto_zoom")) %>
							</td>
						<%else %>
							<td class="content" colspan="2">
								<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "idx_foto_zoom", dizionario("idx_foto_zoom"), "", false) %>
							</td>
				<%  	end if
				else %>
					<td class="content" colspan="2">
						<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "idx_foto_zoom", dizionario("idx_foto_zoom"), "", false) %>
					</td>
				<% end if %>	
   			</tr>
   			
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
				<% 	if i = 0 then %>
					<td class="label_no_width" colspan="2" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo alternativo:</td>
				<% 	end if %>
					<td class="content" colspan="3">
						<table width="100%" cellpadding="0" cellspacing="0">
							<tr>
								<td width="4%" valign="top"><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<% if not rsv.eof then
										if CString(rsv("pub_field_titolo_alt_" + Application("LINGUE")(i))) <> "" then
											'campo sincronizzato: non modificabile 
											%>
											<td width="1%" valign="top"><%= LockIcon %></td>
											<td class="content_disabled">
												<%= TextHtmlEncode(dizionario("idx_alt_"& Application("LINGUE")(i))) %>
											</td>
										<%else %>
											<td class="content">
												<input type="text" class="text" name="idx_alt_<%= Application("LINGUE")(i) %>" value="<%= dizionario("idx_alt_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:95%;">
											</td>
								<% 		end if
								else %>
								<td class="content">
									<input type="text" class="text" name="idx_alt_<%= Application("LINGUE")(i) %>" value="<%= dizionario("idx_alt_"& Application("LINGUE")(i)) %>" maxlength="255" style="width:95%;">
								</td>
								<% end if %>
							</tr>
						</table>
					</td>
					
				</tr>
			<% next %>						
			
			<tr <%= AdminIndexDataViewMode_CssStyle(true) %>><th class="L2" colspan="4">DESCRIZIONE</th></tr>
				<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr <%= AdminIndexDataViewMode_CssStyle(true) %>>
					<td class="content" colspan="4">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>							
								<td width="4%" valign="top"><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
									<% if not rsv.eof then
											if CString(rsv("pub_field_descrizione_" + Application("LINGUE")(i))) <> "" then
												'campo sincronizzato: non modificabile 
												%>
												<td width="1%" valign="top"><%= LockIcon %></td>
												<td class="content_disabled">
													<%= TextHtmlEncode(dizionario("idx_descrizione_"& Application("LINGUE")(i))) %>
												</td>
											<%else %>
												<td class="content">
													<textarea style="width:100%;" rows="3" name="idx_descrizione_<%= Application("LINGUE")(i) %>"><%= dizionario("idx_descrizione_" & Application("LINGUE")(i)) %></textarea>
												</td>
									<%  	end if
									else %>
										<td class="content">
											<textarea style="width:100%;" rows="3" name="idx_descrizione_<%= Application("LINGUE")(i) %>"><%= dizionario("idx_descrizione_" & Application("LINGUE")(i)) %></textarea>
										</td>
									<% end if %>																
							</tr>
						</table>
					</td>
				</tr>
			<%next %>
			
			
			<tr <%= AdminIndexDataViewMode_CssStyle(true) %>><th colspan="4">META TAG</th></tr>
			<tr <%= AdminIndexDataViewMode_CssStyle(true) %>><th colspan="4" class="L2">KEYWORDS</th></tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr <%= AdminIndexDataViewMode_CssStyle(true) %>>
					<td class="content" colspan="4">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>
								<td width="4%" valign="top"><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<% if not rsv.eof then
										if CString(rsv("pub_field_meta_keywords_" + Application("LINGUE")(i))) <> "" then
											'campo sincronizzato: non modificabile 
											%>
											<td width="1%" valign="top"><%= LockIcon %></td>
											<td class="content_disabled">
												<%= TextHtmlEncode(dizionario("idx_meta_keywords_"& Application("LINGUE")(i))) %>
											</td>
										<%
										else %>
											<td><textarea style="width:100%;" rows="2" name="idx_meta_keywords_<%= Application("LINGUE")(i) %>"><%= dizionario("idx_meta_keywords_" & Application("LINGUE")(i)) %></textarea></td>
								<% 		end if
								else %>
									<td><textarea style="width:100%;" rows="2" name="idx_meta_keywords_<%= Application("LINGUE")(i) %>"><%= dizionario("idx_meta_keywords_" & Application("LINGUE")(i)) %></textarea></td>
								<% end if %>																
							</tr>
						</table>
					</td>
				</tr>
			<%next %>
			<tr <%= AdminIndexDataViewMode_CssStyle(true) %>><th colspan="4" class="L2">DESCRIPTION</th></tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr <%= AdminIndexDataViewMode_CssStyle(true) %>>
					<td class="content" colspan="4">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>
								<td width="4%" valign="top"><img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<% if not rsv.eof then
										if CString(rsv("pub_field_meta_description_" + Application("LINGUE")(i))) <> "" then
											'campo sincronizzato: non modificabile 
											%>
											<td width="1%" valign="top"><%= LockIcon %></td>
											<td class="content_disabled">
												<%= TextHtmlEncode(dizionario("idx_meta_description_"& Application("LINGUE")(i))) %>
											</td>
										<%
										else %>
											<td><textarea style="width:100%;" rows="2" name="idx_meta_description_<%= Application("LINGUE")(i) %>"><%= dizionario("idx_meta_description_" & Application("LINGUE")(i)) %></textarea></td>
								<%  	end if
								else %>
									<td><textarea style="width:100%;" rows="2" name="idx_meta_description_<%= Application("LINGUE")(i) %>"><%= dizionario("idx_meta_description_" & Application("LINGUE")(i)) %></textarea></td>
								<% end if %>	
							</tr>
						</table>
					</td>
				</tr>
			<%next %>
   		</table>
   		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;<%= AdminIndexDataViewMode_CssStyle(false) %>">
   			<% 	'check dei permessi dell'utente
   				if ChkPrm(prm_indice_permessi, 0) then
   					'form di gestione dei permessi
   					CALL prmForm(ID, IIF(request("co_F_table_id") <> "", "AL COLLEGAMENTO", "ALLA VOCE"))
   				end if %>
   		</table>
   		<table cellspacing="1" cellpadding="0" class="tabella_madre">
   			<tr>
   				<td class="footer" colspan="4">
   					(*) Campi obbligatori.
   					<input type="submit" class="button" name="salva" value="SALVA">
   				</td>
   			</tr>
   		</table>
   		&nbsp;
   		</form>
    	<%
		set rs = nothing
    End Sub
    
    
    '******************************************************************************************************************************************
    '******************************************************************************************************************************************
    'GESTIONE SALVATAGGIO CATEGORIE
    '******************************************************************************************************************************************
    
    'funzione che formatta l'ordine della tipologia (non valida per "ordine assoluto")
    Private Function FormatOrdine(ordine)
    dim i
    	if CInteger(ordine) = 0 then
    		FormatOrdine = ""
    	else
    		FormatOrdine = Replace(ordine, " ", "")
			if OrdineLenght >= Len(FormatOrdine) then
				FormatOrdine = String(OrdineLenght - Len(FormatOrdine), "0") & FormatOrdine
			end if
    	end if
    End Function
    
    'controlla la validata dei dati dell'indice.
    'ritorna false se errore e imposta session("ERRORE")
    Public Function ChkIndex(byref ID)
        dim value
    	if session("ERRORE") = "" then
    		if CIntero(dizionario("idx_content_id")) = 0 then
    			session("ERRORE") = "Scegliere un contenuto da indicizzare."
    		else
    			'controllo che non ci sia gia una voce con gli stessi dati
    			dim sql, lingua
    			sql = " SELECT idx_id FROM tb_contents_index"& _
    				  " WHERE idx_id <> "& cIntero(ID) & _
    				  " AND idx_content_id = "& cIntero(dizionario("idx_content_id")) & _
    				  " AND "& SQL_IfIsNull(conn, "idx_padre_id", 0) &" = "& CIntero(dizionario("idx_padre_id")) & _
    				  " AND (idx_link_tipo = '"& ParseSql(lnk_interno, adChar) &"' AND idx_link_pagina_id = "& CIntero(dizionario("idx_link_pagina_id")) & _
    				  " 	 OR idx_link_tipo = '"& ParseSql(lnk_esterno, adChar) &"'"
    			for each lingua in application("LINGUE")
    				sql = sql &" AND idx_link_url_"& lingua &" = '"& ParseSQL(CString(dizionario("idx_link_url_"& lingua)), adChar) &"'"
    			next
    			sql = sql &")"
    			
    			value = cString(GetValueList(conn, NULL, sql))
    			if value<>"" then
                    if cIntero(ID)=0 then
                        ID = cIntero(value)
                    end if
                    session("ERRORE") = "Voce gi&agrave; presente nell'indice."
    			end if
    		end if
    	end if
    	
    	ChkIndex = (session("ERRORE") = "")
    End Function
    
    'funzione che salva i  dati della voce dell'indice dal form
    Public function Salva(ID)
        Salva = SalvaPubblicazione(ID, 0)
    end function
    
    'funzione che salva i dati dell'indice.
    public function SalvaPubblicazione(ID, PubblicazioneId)
        dim rs, sql, ContenutoIsSito
    	set rs = server.createobject("adodb.recordset")
        
    	ID = CIntero(ID)
		
'		response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><hr>"
		
		
    	if ChkIndex(ID) then
    		dim padre, rsP, lingua
    		set rsP = server.createobject("adodb.recordset")
    		padre = CIntero(dizionario("idx_padre_id"))
			
			if cIntero(padre) = 0 then
				'padre impostabile a zero solo per i siti. (Solo i siti possono essere le root dell'indice)
				sql = " SELECT COUNT(*) FROM tb_siti_tabelle INNER JOIN tb_contents ON tb_siti_tabelle.tab_id = tb_contents.co_F_table_id " + _
					  " WHERE co_id = " & dizionario("idx_content_id") & " AND tab_name LIKE '"& ParseSql(tabSitoTable, adChar) &"' "
				if cIntero(GetValueList(conn, rs, sql))=0 then
					'non &egrave; un sito: non puo' avere il padre nullo.
					Session("ERRORE") = "Impostare il nodo a cui collegare la voce"
				else
					ContenutoIsSito = true
				end if
			else
				ContenutoIsSito = false
			end if
			
			if Session("ERRORE") = "" then
	
	    		'INSERIMENTO NELL'INDICE
	    		sql = "SELECT * FROM tb_contents_index WHERE idx_id="& cIntero(ID)
	    		rs.open sql, conn, adOpenKeySet, adLockOptimistic
	    		if rs.eof then
	    			rs.addnew
	                rs("idx_autopubblicato") = false
					rs("idx_contUtenti") = 0
					rs("idx_contCrawler") = 0
					rs("idx_contAltro") = 0
					rs("idx_contatore") = 0
					rs("idx_contRes") = Now
	    			CALL SetUpdateParamsRS(rs, "idx_", true)
					
					'li inizializzo a caso perche' obbligatori, verranno calcolati in operazioni_ricorsive...
		    		rs("idx_livello") = 0
	    			rs("idx_foglia") = 1
	    		else
	    			CALL SetUpdateParamsRS(rs, "idx_", false)
	    		end if
	    		rs("idx_content_id") = dizionario("idx_content_id")
				rs("idx_principale") = CBoolean(dizionario("idx_principale"), false)
	    		
	    		'se cambio padre guarda se quello vecchio e' diventato foglia
	    		if CIntero(rs("idx_padre_id")) > 0 AND rs("idx_padre_id") <> padre then
	    			sql = "SELECT COUNT(*) FROM tb_contents_index WHERE idx_padre_id = "& cIntero(rs("idx_padre_id"))
	    			if CIntero(GetValueList(conn, rsP, sql)) = 1 then
	    				conn.Execute("UPDATE tb_contents_index SET idx_foglia = 1 WHERE idx_id = "& cIntero(rs("idx_padre_id")))
	    			end if
	    		end if
	            
	    		if cIntero(padre) = 0 then
	    			rs("idx_padre_id") = null
	    		else
	    			rs("idx_padre_id") = padre
	    		end if
	    		
	    		'gestione link
	    		sql = " SELECT tab_parametro FROM tb_siti_tabelle t"& _
	    			  " INNER JOIN tb_contents c ON t.tab_id = c.co_F_table_id"& _
	    			  " WHERE co_id = "& cIntero(dizionario("idx_content_id"))
'response.write sql & "<br>"
'response.end

	    		CALL LinkCalculate(conn, "idx", rs, dizionario, "idx_link_pagina_id", "idx_link_url_", GetValueList(conn, rsP, sql))
	    		
				'informazioni di pubblicazione personalizzati
				if Left(UCase(TypeName(dizionario)), Len("IREQUEST")) = "IREQUEST" OR _
					content.IsRaggruppamentoById(dizionario("idx_content_id")) then		'se ho i dati dal form li salvo o se è un raggruppamento.
		    		if dizionario("idx_ordine")<>"" then
		    			rs("idx_ordine") = cInteger(dizionario("idx_ordine"))
		    		else
		    			rs("idx_ordine") = NULL
		    		end if
					
		    		rs("idx_foto_thumb") = dizionario("idx_foto_thumb")
		    		rs("idx_foto_zoom") = dizionario("idx_foto_zoom")
					
					for each lingua in Application("LINGUE")
						rs("idx_titolo_" + lingua) = dizionario("idx_titolo_" + lingua)
						rs("idx_descrizione_" + lingua) = dizionario("idx_descrizione_" + lingua)
						rs("idx_alt_" + lingua) = dizionario("idx_alt_" + lingua)
						rs("idx_meta_keywords_" + lingua) = dizionario("idx_meta_keywords_" + lingua)
						rs("idx_meta_description_" + lingua) = dizionario("idx_meta_description_" + lingua)
					next
				end if
				
				'imposta id del sito di appartenenza
				if cIntero(rs("idx_link_pagina_id"))>0 then
					'lo determina direttamente dalla pagina associata.
					sql = " SELECT id_web FROM tb_pagineSito WHERE id_pagineSito=" & cIntero(rs("idx_link_pagina_id"))
					rs("idx_webs_id") = cIntero(GetValueList(conn, NULL, sql))
				else
					'voce senza pagina associata: calcola id_Web
					if ContenutoIsSito then
						'e' un sito lo imposta direttamente all'id del contenuto stesso.
						sql = "SELECT co_F_key_id FROM tb_contents WHERE co_id=" & dizionario("idx_content_id")
					else
						'e' un raggruppamento o una voce senza link
						sql = "SELECT idx_webs_id FROM tb_contents_index WHERE idx_id = " & rs("idx_padre_id")
					end if
					rs("idx_webs_id") = cIntero(GetValueList(Conn, rsp, sql))
				end if
	    		rs.update
	    		ID = rs("idx_id")
				
				if cIntero(rs("idx_link_pagina_id"))>0 AND cIntero(rs("idx_padre_id"))>0 then
					sql = " SELECT COUNT(*) FROM tb_webs " + _
						  " WHERE id_webs IN (SELECT id_web FROM tb_paginesito WHERE id_paginesito=" & rs("idx_link_pagina_id") & ") " + _
						  "    OR id_webs IN (SELECT idx_webs_id FROM tb_contents_index WHERE idx_id = " & rs("idx_padre_id") & ")"
					if cIntero(GetValueList(conn, rsp, sql))>1 then
						rs.close
						Session("ERRORE") = "Il nodo non pu&ograve; essere pubblicato in un ramo appartenente ad un sito diverso da quello associato alla pagina."
						exit function
					end if
				end if
				rs.close
	    		
	    		'se co_F_key_id = 0 significa che sto inserendo una voce dell'indice come contenuto
	    		sql = " SELECT co_F_key_id FROM tb_contents WHERE co_id = "& cIntero(dizionario("idx_content_id"))
	    		if CIntero(GetValueList(conn, rs, sql)) = 0 then
	    			sql = " UPDATE tb_contents SET co_F_key_id = co_id WHERE co_id = "& cIntero(dizionario("idx_content_id"))
	    			conn.Execute(sql)
	    		end if
	    		
	    		'OPERAZIONI RICORSIVE
				CALL operazioni_ricorsive_tipologia(ID)
	    		'imposta operazioni tipologia e tipologie figlie
		    	if DB_Type(conn) = DB_Access then
		    		conn.committrans
		    		conn.begintrans
		    	end if
	    		if ChkPrm(prm_indice_permessi, 0) then
	    			'GESTIONE PERMESSI SEZIONE
	    			CALL SalvaPermessi(ID, padre)
	    		end if
            end if
			
    		set rsP = nothing
            
            SalvaPubblicazione = ID
        else
            SalvaPubblicazione = 0
    	end if
        
        if cIntero(ID)>0 AND cIntero(PubblicazioneId)>0 then
            'blocca la voce perche' pubblicata automaticamente
            
            'imposta relazione con pubblicazione automatica se non gia' presente.
            sql = " SELECT COUNT(*) FROM rel_index_pubblicazioni " + _
                  " WHERE rip_idx_id = " & cIntero(ID) & " AND rip_pub_id = " & cIntero(PubblicazioneId)
            if CIntero(GetValueList(conn, rs, sql)) = 0 then
                sql = "INSERT INTO rel_index_pubblicazioni (rip_idx_id, rip_pub_id) " + _
                      " VALUES (" & cIntero(ID) & ", " & cIntero(PubblicazioneId) & ")"
                CALL conn.execute(sql, ,adExecuteNoRecords)
            end if
            
            'imposta flag di blocco
            sql = "UPDATE tb_contents_index SET idx_autopubblicato=1 WHERE idx_id=" & cIntero(ID)
            CALL conn.execute(sql,,adexecuteNoRecords)
            
        end if
		
		UpdateSentinelTable("tb_contents_index_sentinel")
		
'		sql = "SELECT * FROM tb_contents_index WHERE idx_id=" & ID
'		rs.open sql, conn, adOpenstatic, adlockoptimistic, adcmdtext
'		CALL listrecordset(rs, true)
'		rs.close
		
    	set rs = nothing
      
	  
    End function
	
	
	'-------------------------------------------------------------------------------------------- GENERAZIONE URL RW
	'restituisce una chiave dal record
	Function GetChiave(rs, lingua)
		GetChiave = CBLL(rs, "co_chiave", lingua)
		if CString(GetChiave) = "" then
			GetChiave = content.Codifica(CBLL(rs, "co_titolo", lingua))
		end if
		if CString(GetChiave) = "" then
			GetChiave = rs("idx_id")
		end if
	End Function
	
	
	'calcola l'URL RW finale dal precedente piu il recordset aperto sulla JOIN indice - contenuto
	Function ConcatenaUrl(urlPrecedente, rs, chiave, lingua)
		if content.IsSito(rs("tab_name")) OR content.IsHomePage(rs) then										'sito o home page
			ConcatenaUrl = lingua &"/"
		elseif CIntero(rs("idx_link_pagina_id")) = 0 AND NOT content.IsRaggruppamento(rs("tab_name")) then		'link esterno
			ConcatenaUrl = CBLL(rs, "idx_link_url", lingua)
		elseif CString(chiave) <> "" then																		'chiave sovrascritta
			ConcatenaUrl = urlPrecedente & chiave & IIF(Right(chiave, 1) = "/", "", "/")
		else																									'link interno
			ConcatenaUrl = urlPrecedente & GetChiave(rs, lingua) &"/"
		end if
		
		if Len(ConcatenaUrl) > rs("idx_link_url_rw_it").DefinedSize then
			session("ERRORE") = "URL del record troppo lungo."
			SendEmailSupport("URL del record troppo lungo." & vbCrLf & "ID: " & rs("idx_id") &";"& vbCrLf & "URL:" & ConcatenaUrl & ";")
			ConcatenaUrl = Left(ConcatenaUrl, rs("idx_link_url_rw_it").DefinedSize - 1 - Len(rs("idx_id"))) &"_"& rs("idx_id")
		end if
		
	End Function
	
	
	'restituisce l'ultima chiave dell'url RW
	Function GetUltimaChiave(urlRW)
		if CString(urlRW) <> "" then
			GetUltimaChiave = Right(urlRW, Len(urlRW) - InStrRev(urlRW, "/", Len(urlRW) - 1, vbTextCompare))
		end if
	End Function
	
	
	'-------------------------------------------------------------------------------------------- UNIVOCITA URL RW
	'restituisce true se l'indice in input ha un URL RW univoco
	Function IsUnivoco(conn, rs, url, lingua)
		if CIntero(rs("idx_link_pagina_id")) = 0 AND NOT content.IsRaggruppamento(rs("tab_name")) then			'link esterno
			IsUnivoco = true
		elseif content.IsSito(rs("tab_name")) then																'sito
			IsUnivoco = true
		elseif content.IsHomePage(rs) then																		'home page
			IsUnivoco = true
		else																									'link interno
			dim l, sql
			
			'controlla in tutte le lingue
			'sql = " SELECT COUNT(*) FROM tb_contents_index"& _
			'	  " WHERE idx_webs_id = "& cIntero(rs("idx_webs_id")) &" AND (1=0"
			'for each l in Application("LINGUE")
			'	'controllo univocità fra lingue all'interno della stessa voce
			'	if lingua <> l AND CString(rs("idx_link_url_rw_" + l)) = url then
			'		IsUnivoco = false
			'		exit function
			'	end if
				
			'	sql = sql &" OR idx_link_url_rw_"& l &" LIKE '"& ParseSql(url, adChar) &"'"
			'	if l = lingua then
			'		sql = sql &" AND NOT idx_id = "& cIntero(rs("idx_id"))
			'	end if
			'next
			
			'NICOLA: 24/11/2011
			'modificato, in quanto gli url registrati sono per forza diversi lingua per lingua, quindi basta cercare all'itnerno della lingua corrente
			sql = " SELECT COUNT(*) FROM tb_contents_index"& _
				  " WHERE idx_webs_id = "& cIntero(rs("idx_webs_id")) & _
						" AND idx_link_url_rw_"& lingua &" LIKE '"& ParseSql(url, adChar) &"' " & _
						" AND idx_id <> " & cIntero(rs("idx_id"))
			IsUnivoco = CIntero(GetValueList(conn, NULL, sql)) = 0
		end if
	End Function
	
	
	'rende l'url RW univoco
	Sub RendiUnivoco(conn, urlPrecedente, rs, lingua)
		dim url
		url = ConcatenaUrl(urlPrecedente, rs, "", lingua)
		if NOT IsUnivoco(conn, rs, url, lingua) then
			'concateno ID
			url = ConcatenaUrl(urlPrecedente, rs, GetChiave(rs, lingua) &"_"& lingua &"_"& rs("idx_id"), lingua)
		end if

		rs("idx_link_url_rw_" + lingua) = url
	End Sub
	
	
	'-------------------------------------------------------------------------------------------- SALVATAGGIO VECCHIO URL RW
	'salva i vecchi url RW nel dizionario
	Sub SalvaUrlRW(rs, dizionario)
		dim lingua
		for each lingua in Application("LINGUE")
			dizionario(lingua) = CString(rs("idx_link_url_rw_" + lingua))
		next
		dizionario("paginaId") = CIntero(rs("idx_link_pagina_id"))
	End Sub
	
	
	'salva il vecchio url RW del record nella relazione per il redirect
	Sub SalvaRedirect(rs, oldURLs)
		dim sql, lingua
		for each lingua in Application("LINGUE")
			if CString(oldURLs(lingua)) <> "" AND CString(oldURLs(lingua)) <> CString(rs("idx_link_url_rw_" + lingua)) AND _
			   oldURLs("paginaId") > 0 then
				sql = " INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id) " & _
					  " SELECT TOP 1 " & cIntero(rs("idx_id")) &", '"& ParseSql(oldURLs(lingua), adChar) &"', '"& lingua &"', "& _
					  					 	 SQL_Now(conn) &", "& cIntero(Session("ID_ADMIN")) &", "& SQL_Now(conn) &", "& cIntero(Session("ID_ADMIN")) & _
					  " FROM aa_versione" + _
					  " WHERE ( SELECT COUNT(*) FROM rel_index_url_redirect " + _
					  " 		WHERE riu_idx_id=" & cIntero(rs("idx_id")) &" AND riu_url LIKE '"& ParseSql(oldURLs(lingua), adChar) &"' AND riu_lingua='"& lingua &"')=0 "
				conn.Execute(sql)
			end if
		next
	End Sub
	
	Function CalcolaChecksum(rs)
	
		dim campi,i,lingua,campi_ar
		
		campi="idx_id,idx_livello,idx_tipologia_padre_base,idx_visibile_assoluto," &_
			  "idx_ordine_assoluto,idx_ordine,co_ordine,idx_foglia"
		for each lingua in Application("LINGUE")
			campi=campi + ",idx_link_url_rw_" + lingua
			campi=campi + ",idx_link_url_" + lingua
		next
		
		campi_ar=split(campi,",")
		
		for i=0 to Ubound(campi_ar)
			CalcolaChecksum=CalcolaChecksum & "|" & rs(cString(campi_ar(i)))
		next
		
		CalcolaChecksum=Right(CalcolaChecksum,Len(CalcolaChecksum)-1)
		
	End Function
	
	
	
	'ID: id del nodo da cui far partire la ricorsione
	'padreId: id del padre del nodo ID
	Sub operazioni_ricorsive_tipologia(ID)
		dim rs, sql, lingua, d
		set rs = server.CreateObject("ADODB.recordset")
		
		dim indicePrecedente, ordineS, listaPadriUpdate, i
		'dati ricorsivi
		const idx = 0
		const livello = 1
		const tipologiaBaseId = 2
		const tipologieListaIds = 3
		const visibile = 4
		const ordine = 5
		const ordineStop = 6
		const foglia = 7
		const urlRW = 8
		redim preserve stack(8, 0)			'il valore 8 va modificato anche sul redim successivo
		
		dim oldUrlRW		'dizionario contenente gli url RW prima della modifica
		set oldUrlRW = CreateObject("Scripting.Dictionary")
		
		'memorizzo i dati del primo padre
		sql = " SELECT " + _
			  " idx_id, idx_livello, idx_tipologia_padre_base, idx_tipologie_padre_lista, idx_visibile_assoluto, idx_ordine_assoluto, idx_ordine, " + _
			  " co_ordine, idx_foglia, " & SQL_MultiLanguage("idx_link_url_rw_<LINGUA>", ", ") & _ 
			  " FROM v_indice " & _
			  " WHERE idx_id = (SELECT idx_padre_id FROM tb_contents_index WHERE idx_id = " & cIntero(id) &")"
		rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdtext
		set stack(urlRW, 0) = CreateObject("Scripting.Dictionary")
		if rs.eof then			'il nodo corrente e' di livello 0, nessun padre
			stack(livello, 0) = -1
			stack(tipologiaBaseId, 0) = ID
			stack(tipologieListaIds, 0) = ""
			stack(visibile, 0) = true
			stack(ordine, 0) = ""
			stack(ordineStop, 0) = false
			stack(foglia, 0) = false
		else
			stack(idx, 0) = rs("idx_id").value
			stack(livello, 0) = rs("idx_livello").value
			stack(tipologiaBaseId, 0) = rs("idx_tipologia_padre_base").value
			stack(tipologieListaIds, 0) = rs("idx_tipologie_padre_lista").value
			stack(visibile, 0) = rs("idx_visibile_assoluto").value
			stack(ordine, 0) = rs("idx_ordine_assoluto").value
			'se l'ordine vuoto lo annullo per tutto il ramo seguente
			stack(ordineStop, 0) = CString(rs("idx_ordine_assoluto")) = "" OR (CString(rs("idx_ordine")) = "" AND CString(rs("co_ordine")) = "")
			stack(foglia, 0) = rs("idx_foglia").value
			
			for each lingua in Application("LINGUE")
				stack(urlRW, 0)(lingua) = rs("idx_link_url_rw_" + lingua).value
			next
		end if
		indicePrecedente = 0
		rs.close
		
		
		sql = " SELECT * FROM (tb_contents_index i " + _
		      " INNER JOIN tb_contents c ON i.idx_content_id = c.co_id) " + _
			  " INNER JOIN tb_siti_tabelle t ON c.co_F_table_id = t.tab_id "
		if DisableRicorsione then
			'ricorsione disattivata per aggiornamento bulk. Esegue solo ricalcolo del record modificato e non dei figli.
			sql = sql + " WHERE idx_id = " & cIntero(ID)
		else
			sql = sql + " WHERE "& SQL_IdListSearch(conn, "i.idx_tipologie_padre_lista", cIntero(ID)) & " OR idx_id = "& cIntero(ID) & _
				  " ORDER BY idx_tipologie_padre_lista"
		end if
		'response.write sql
		rs.open sql, conn, 3, adLockOptimistic, adCmdtext
		
'response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>" & sql & "<br>" + vbcrlf
'response.end
'CALL ListRecordset(rs, true)
		'controllo univocita e calcolo url del nodo in modifica: il primo nodo è sempre quello in modifica corrente.
		
		dim check_before,check_after
		check_before=CalcolaChecksum(rs)
		
		dim url
		if not rs.eof then
			CALL SalvaUrlRW(rs, oldUrlRW)
			for each lingua in Application("LINGUE")
				CALL RendiUnivoco(conn, stack(urlRW, 0)(lingua), rs, lingua)
			next
			CALL SalvaRedirect(rs, oldUrlRW)
		end if
		
		do while not rs.eof			
			
			'cerco il nodo padre
			while indicePrecedente > 0 AND stack(idx, indicePrecedente) <> rs("idx_padre_id")
				indicePrecedente = indicePrecedente - 1
			wend
			
			'lista tipologie id
			if stack(tipologieListaIds, indicePrecedente) = "" then
				rs("idx_tipologie_padre_lista") = rs("idx_id")
			else
				rs("idx_tipologie_padre_lista") = stack(tipologieListaIds, indicePrecedente) &","& rs("idx_id")
			end if
			
			'ordine
			ordineS = false
			if NOT stack(ordineStop, indicePrecedente) then
				if CString(rs("idx_ordine")) <> "" then
		    		rs("idx_ordine_assoluto") = stack(ordine, indicePrecedente) & FormatOrdine(rs("idx_ordine")) &"-"
		    	elseif CString(rs("co_ordine")) <> "" then
		    		rs("idx_ordine_assoluto") = stack(ordine, indicePrecedente) & FormatOrdine(rs("co_ordine")) &"-"
				else
					ordineS = true
					rs("idx_ordine_assoluto") = stack(ordine, indicePrecedente)
		    	end if
			else
				rs("idx_ordine_assoluto") = stack(ordine, indicePrecedente)
				ordineS = true
			end if
			
			'foglia
			if stack(foglia, indicePrecedente) then
				listaPadriUpdate = listaPadriUpdate & stack(idx, indicePrecedente) &","
			end if
			
			'url RW
			if indicePrecedente > 0 then		'se 0: nodo principale con URL gia modificati
				CALL SalvaUrlRW(rs, oldUrlRW)
				for each lingua in Application("LINGUE")
					if CString(rs("idx_link_url_rw_" + lingua)) <> "" then		'ho gia verificato l'univocita dell'ultima chiave
						rs("idx_link_url_rw_" + lingua) = ConcatenaUrl(stack(urlRW, indicePrecedente)(lingua), rs, _
																	   GetUltimaChiave(rs("idx_link_url_rw_" + lingua)), lingua)
					else
						CALL RendiUnivoco(conn, stack(urlRW, indicePrecedente)(lingua), rs, lingua)
					end if
				next
				CALL SalvaRedirect(rs, oldUrlRW)
			end if
			
			rs("idx_livello") = stack(livello, indicePrecedente) + 1
			rs("idx_tipologia_padre_base") = stack(tipologiaBaseId, indicePrecedente)
			rs("idx_visibile_assoluto") = stack(visibile, indicePrecedente) AND rs("co_visibile")
			
			
			
			if cString(rs("idx_tipologie_padre_lista")) <>"" then
				'calcolo della priorità
				dim rs_count, priorita
				set rs_count = Server.CreateObject("ADODB.recordset")
				sql = "SELECT COUNT(*) AS CONTATI FROM v_indice WHERE idx_id IN (" & rs("idx_tipologie_padre_lista") & ")" + _
					  " AND tab_name LIKE '" & tabRaggruppamentoTable & "'"
					  
				rs_count.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

				priorita = rs("tab_priorita_base") + (10000 - (rs_count("CONTATI")*1000))
				
				rs_count.close
				sql = "SELECT COUNT(*) AS CONTATI FROM v_indice WHERE idx_id IN (" & rs("idx_tipologie_padre_lista") & ")" + _
					  " AND tab_sito_id = " & rs("tab_sito_id")
				rs_count.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				priorita = priorita + (rs_count("CONTATI")*100)

				rs("idx_priorita") = priorita
			end if

			
				','" + Connection.NextSql.Concat() + "v.idx_tipologie_padre_lista" + 
'Connection.NextSql.Concat() + "',' LIKE '%,'" + Connection.NextSql.Concat() + Connection.NextSql.CastAsString(" v_indice.idx_id ") + Connection.NextSql.Concat() + "',%' " +
'"  AND tab_name LIKE '" + BLLIndiceTabella.TabellaNome.Raggruppamenti + "' ) AS rag, " +
'" (SELECT COUNT(*) FROM v_indice " +
'"  WHERE ','" + Connection.NextSql.Concat() + "v.idx_tipologie_padre_lista" + 
'Connection.NextSql.Concat() + "',' LIKE '%,'" + Connection.NextSql.Concat() + 
'Connection.NextSql.CastAsString(" v_indice.idx_id ") + Connection.NextSql.Concat() + "',%' " +
'"  AND tab_sito_id = v.tab_sito_id ) AS tipo, " +			
			
			'aggiunta stack
			if NOT rs("idx_foglia") then
				'aggiungo nodo
				indicePrecedente = indicePrecedente + 1
				redim preserve stack(8, indicePrecedente)
				stack(idx, indicePrecedente) = rs("idx_id").value
				stack(livello, indicePrecedente) = rs("idx_livello").value
				stack(tipologiaBaseId, indicePrecedente) = rs("idx_tipologia_padre_base").value
				stack(tipologieListaIds, indicePrecedente) = rs("idx_tipologie_padre_lista").value
				stack(visibile, indicePrecedente) = rs("idx_visibile_assoluto").value
				stack(ordine, indicePrecedente) = rs("idx_ordine_assoluto").value
				stack(ordineStop, indicePrecedente) = ordineS
				
				set stack(urlRW, indicePrecedente) = CreateObject("Scripting.Dictionary")
				for each lingua in Application("LINGUE")
					stack(urlRW, indicePrecedente)(lingua) = rs("idx_link_url_rw_" + lingua).value
				next
			end if
			
			'update viene chiamato automaticamente dal movenext
			if check_before = "" then
				rs.movenext
			else			
				check_after=CalcolaChecksum(rs)								
				if check_after = check_before then
					rs.update
					Exit Do
				else
					rs.movenext
				end if				
				check_before=""
			end if				
			
		loop 
		
		rs.close
		set rs = nothing
		set oldUrlRW = nothing
		
		'tolgo il flag foglia dai padri (il reinserimento del flag avviene in un'altra procedura)
		if listaPadriUpdate <> "" then
			listaPadriUpdate = Left(listaPadriUpdate, Len(listaPadriUpdate)-1)
			sql = "UPDATE tb_contents_index SET idx_foglia = 0 WHERE idx_id IN ("& listaPadriUpdate &")"
			conn.Execute(sql)
		end if
		
	End Sub
    
    
    '******************************************************************************************************************************************
    '******************************************************************************************************************************************
    'GESTIONE CANCELLAZIONE VOCE
    '******************************************************************************************************************************************
    
    'cancellazione della voce, di tutti i figli ed eventualmente anche del contenuto se la voce indicata e' un raggruppamento
    Public Sub Delete(ID)
        dim ContenutoDaCancellare, PadreId
        dim rs, sql
        set rs = server.CreateObject("ADODB.Recordset")
        
        ID = cIntero(ID)
        
        'apre elemento per determinarne padre e tipo
        sql = "SELECT idx_padre_id, tab_name, idx_content_id FROM v_indice WHERE idx_id=" & cIntero(ID)
        rs.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
        if not rs.eof then
        if Content.IsRaggruppamento(rs("tab_name")) then
            'e' un raggruppamento: cancella anche il contenuto
            ContenutoDaCancellare = cIntero(rs("idx_content_id"))
        else
            'non e' un raggruppamento: il contenuto non va considerato.
            ContenutoDaCancellare = 0
        end if
        PadreId = cIntero(rs("idx_padre_id"))
            rs.close
            
            
            'cancellazione delle voci indicata ed eventuali sottovoci dell'indice
            sql = " SELECT idx_id FROM tb_contents_index WHERE " & SQL_IdListSearch(conn, "idx_tipologie_padre_lista", cIntero(ID)) & _
                  " ORDER BY idx_livello DESC, idx_foglia DESC"
            rs.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
            'compone query di cancellazione
            sql = ""
            while not rs.eof 
                sql = sql + "DELETE FROM tb_contents_index WHERE idx_id=" & cIntero(rs("idx_id")) & ";"
                rs.movenext
            wend
            rs.close
            'esegue cancellazione delle voci
            CALL ExecuteMultipleSql(conn, sql, false)
            
            'cancellazione del contenuto (solo raggruppamenti)
            if ContenutoDaCancellare>0 then
                CALL content.Delete(ContenutoDaCancellare)
            end if
            
            'aggiornamento dello stato del padre (se diventa foglia o ha altri figli)
            if PadreId > 0 then
                sql = "SELECT COUNT(*) FROM tb_contents_index WHERE idx_padre_id=" & cIntero(PadreId) & " AND idx_id<>" & cIntero(ID)
                if cIntero(GetValueList(conn, rs, sql))=0 then
                    'nessun altro figlio per il padre: diventa foglia
                    sql = "UPDATE tb_contents_index SET idx_foglia=1 WHERE idx_id=" & cIntero(PadreId)
                    CALL conn.execute(sql, , adExecuteNoRecords)
                end if
            end if
        else
            rs.close
        end if
        set rs = nothing
    End Sub
    
    
    '.................................................................................................
    'GESTIONE DEI PERMESSI
    '.................................................................................................
    
    'CREAZIONE PERMESSI
    
    'visualizza il form per l'inserimento dei permessi sia in inserimento sia in modifica
    Sub prmForm(ID, label)
    	if session("WEB_ADMIN") <> "" then
    		dim rs, sql
    		set rs = server.createobject("adodb.recordset") %>
    		<tr><th colspan="4">PERMESSI DI ACCESSO <%= label %></th></tr>
    		<tr>
    			<td colspan="4">
    				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
    					<tr>
    						<td colspan="4">
    							<table cellpadding="0" cellspacing="0" width="100%">
    								<tr>
    									<td class="label" nowrap>Elenco degli utenti</td>
    									<td class="content_right" style="font-size: 1px; padding-right:1px;">
    										<a id="tutti" class="button_L2" href="javascript:void(0);" onclick="tutti()" title="seleziona tutti gli utenti elencati" <%= ACTIVE_STATUS %>>
    											ABILITA TUTTI
    										</a>
    										&nbsp;
    										<a id="nessuno" class="button_L2" href="javascript:void(0);" onclick="nessuno()" title="toglie la selezione a tutti gli utenti elencati" <%= ACTIVE_STATUS %>>
    											DISABILITA TUTTI
    										</a>
    									</td>
    								</tr>
    							</table>
    						</td>
    					</tr>
                        <tr>
                            <td valign="top">
    					<%dim antenati, nUtenti, nRighe, nRiga
                		sql = "SELECT idx_tipologie_padre_lista FROM tb_contents_index WHERE idx_id = "& CIntero(dizionario("idx_padre_id"))
                		antenati = GetValueList(conn, rs, sql)
                		if antenati = "" then
                			antenati = "0"
                		end if
                        
                        sql = " FROM tb_admin a LEFT JOIN rel_index_admin r"& _
                			  " ON (a.id_admin = r.ria_admin_id"& _
                			  "	    AND (ria_index_id IN ("& antenati &") OR ria_index_id = "& CIntero(ID) &"))"& _
                			  " WHERE EXISTS (SELECT 1 FROM rel_admin_sito WHERE admin_id = a.id_admin) "
    
                        'recupera numero di utenti
                        nUtenti = cIntero(GetValueList(conn, rs, "SELECT COUNT(id_admin) " + sql))
                        nRighe = (nUtenti \ 2) + (nUtenti mod 2)
                        nRiga = 0 
                        
                		sql = " SELECT id_admin, admin_cognome, admin_nome, MIN(ria_index_id) AS idx " + _
                              sql + " GROUP BY id_admin, admin_cognome, admin_nome" + _
                              " ORDER BY admin_cognome, admin_nome"
                		rs.open sql, conn, adOpenStatic, adLockOptimistic
                        
                		while not rs.eof 
                            if nRiga mod nRighe = 0 then
                                if nRiga > 0 then%>
                                        </table>
                                    </td>
                                    <td>
                                <% end if %>
                                    <table cellpadding="0" cellspacing="1" align="left">
                                        <tr>
                                            <th class="l2_center" style="width:1%;">abilita</th>
    						                <th class="l2">cognome e nome</th>
                                        </tr>
                            <% end if %>
        					<tr>
        						<td class="content_center">
        							<input type="checkbox" class="checkbox" id="prm_<%= rs("id_admin") %>" name="prm_<%= rs("id_admin") %>" value="1"
        							 <%= Chk(CIntero(rs("idx")) > 0 OR request("prm_"& rs("id_admin")) <> "") %>
        							 <%= Disable(CIntero(rs("idx")) > 0 AND rs("idx") <> CIntero(ID)) %>>
        						</td>
        						<td class="label"><%= rs("admin_cognome") %>&nbsp;<%= rs("admin_nome") %></td>
        					</tr>
                            <%rs.movenext
                            nRiga = nRiga + 1
            		    wend 
                        if nRiga mod nRighe > 0 then%>
                           <tr><td class="content_center">&nbsp;</td><td class="label">&nbsp;</td></tr>
                        <% end if %>
                            </table>
                            </td>
                        </tr>
    				</table>
    			</td>
    		</tr>
    		<script language="JavaScript">
    			function tutti() {
    				for(var i=0; i < form1.elements.length; i++)
    					if (form1.elements(i).id.substring(0, 4) == "prm_")
    						form1.elements(i).checked = true
    			}
    	
    			function nessuno() {
    				for(var i=0; i < form1.elements.length; i++)
    					if (form1.elements(i).id.substring(0, 4) == "prm_")
    						form1.elements(i).checked = false
    			}
    		</script>
    <%	end if
    End Sub
    
    'salva i permessi dato il form, una connessione e l'ID della sezione
    'da richiamare in Gestione_Relazioni_record
    'conn:			connessione al dbLayer
    'sezID:			ID della sezione
    'sezPadreID:	CIntero(ID della sezione padre)
    Sub SalvaPermessi(sezID, sezPadreID)
    	'cancello i precedenti permessi per sovrascriverli
    	conn.execute("DELETE FROM rel_index_admin WHERE ria_index_id="& cIntero(sezID))
    	
    	dim campo
    	for each campo in request.form
    		if left(campo, 4) = "prm_" AND request.form(campo) <> "" then
    			conn.execute(" INSERT INTO rel_index_admin(ria_index_id, ria_admin_id) " & _
    				  		 " VALUES ("& cIntero(sezID) &", "& cIntero(right(campo, len(campo)-4)) &")")
    		end if
    	next
    End Sub
    
    'VERIFICA PERMESSI (controllare anche ObjContent.SQLPermessi)
    'restituisce true se l'utente ha i permessi per l'azione indicati dal parametro di ingresso
    'prm:			permesso da verificare, corrispondente ad una delle costanti
    'sezione:		sezione su cui controllare i permessi, obbligatoria solo per il permesso prm_pagine_altera
    Function ChkPrm(prm, sezID)
    	if session("WEB_ADMIN") = "" AND session("WEB_POWER") = "" AND session("WEB_USER") = "" then	'se sono in un applicativo ext
    		'vedo se l'admin corrente ha permessi sul nextweb
    		dim rs
    		set rs = conn.execute(" SELECT * FROM rel_admin_sito"& _
    							  " WHERE admin_id = "& CIntero(session("ID_ADMIN")) &" AND sito_id = "& NEXTWEB5 & _
    							  " ORDER BY rel_as_permesso")
    		if not rs.eof then
    			SELECT CASE rs("rel_as_permesso")
    				CASE 1
    					session("WEB_ADMIN") = "ext"
    				CASE 2
    					session("WEB_POWER") = "ext"
    				CASE 3
    					session("WEB_USER") = "ext"
    			END SELECT
    		end if
    		
    		set rs = nothing
    	end if
    	
    	if session("WEB_ADMIN") <> "" then
    		ChkPrm = true
    	elseif session("WEB_POWER") <> "" then
    		SELECT CASE prm
    			CASE prm_indice_permessi, prm_stili_accesso, prm_plugin_accesso, prm_strumenti_accesso, prm_siti_gestione, _
    				 prm_pubblicazioni_accesso, prm_immaginiFormati_accesso
    													'permessi di modifica dei permessi dei WEB_USER, area stili, area oggetti,
    													'area strumenti, creazione / modifica / cancellazione siti,
														'pubblicazioni automatiche, area formati immagini
    				ChkPrm = false
    			CASE ELSE
    				ChkPrm = true
    		END SELECT
    	else
    		SELECT CASE prm
    			CASE prm_pagine_altera					'permessi di creazione / mofifica / cancellazione pagine
    				dim sql
    				sql = " SELECT COUNT(*) FROM rel_index_admin"& _
    					  " WHERE ria_admin_id = "& CIntero(session("ID_ADMIN"))
    				if CIntero(sezID) > 0 then
    					sql = sql &" AND ria_index_id = "& cIntero(sezID)
    				end if
    				
    				ChkPrm = (CIntero(GetValueList(conn, NULL, sql)) >= 1)
    			CASE ELSE
    				ChkPrm = false
    		END SELECT
    	end if
    End Function

End Class
%>