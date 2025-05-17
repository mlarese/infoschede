<%
'************************************************************************************************
'COSTANTI PER LA DEFINIZIONE DELLE POSIZIONI NELL'ARRAY DELLE 
const RELAZIONI_POS_TABLE   = 0
const RELAZIONI_POS_PK      = 1
const RELAZIONI_POS_FK      = 2
const RELAZIONI_POS_DES     = 3

'************************************************************************************************
Class ObjCategorie

'nome categoria
Public nomeSingolare					'nome da visualizzare (es.: "categoria", "tipologia", ...)
Public nomePlurale

'tabella principale
Public tabella
Public prefisso							'prefisso del nome dei campi (prefisso_)

'gestione delle tabelle correlate
Private Relazioni()
Private RelazioniCount
'gestione della relazione principale che vincola i descrittori
Private RelazioneDescrittori_Table
Private RelazioneDescrittori_Pkfield
Private RelazioneDescrittori_FkField

'gestione della eventuale relazione con caratteristiche tecniche
Public tabellaRelCaratteristiche
Public chiaveEsternaRelCaratteristiche
Public idCarRelCaratteristiche
Public ordineRelCaratteristiche
Public lockedRelCaratteristiche
'gestione delle eventuali caratteristiche tecniche (descrittori)
Public tabellaCaratteristiche
Public idCaratteristiche
Public nomeCaratteristiche
Public tipoCaratteristiche
'gestione della eventuale relazione tra caratteristiche e gruppi
Public tabellaGruppiCaratteristiche
Public idGruppiCaratteristiche
Public idRelGruppiCaratteristiche
Public nomeGruppiCaratteristiche
Public ordineGruppiCaratteristiche
'gestione della eventuale relazione tra tabella correlata e caratteristiche tecniche
Public tabellaRelCorCaratteristiche
Public idArtRelCorCaratteristiche
Public idCarRelCorCaratteristiche
'gestione della eventuale relazione tra categorie e utenti dell'area riservata
Public tabellaRelUtenti
Public idUtenteRelUtenti
Public idCategoriaRelUtenti

'gestione eventuali relazioni esterne
Public RelazioniEsterne_Label
Public RelazioniEsterne_Link

'abilita/disabilita campi
Public abilitaBlocchiEsterni			'parametro che abilita la gestione dei blocchi esterni su _external_source, _external_id, _locked
Public abilitaLogo
Public abilitaFoto
Public abilitaDescrittori
Public categorieBloccate
Public filtroCategorieBase              'parametro che filtra la visualizzazione delle categorie a partire dalla lista di nodi indicata
Public abilitaPermessiAreaRiservata

Public attivaCKEditorPerDescrizione 	'se true attiva CKEditor per il campo descrizione della categoria

'connessione DB
Public conn

'parametri generali di funzionamento
Public isB2B							'true se sono nelle categorie del NEXT-b2b
Public GestioneCategorieMiste			'se false: i record categorizzati possono essere associati solo alle foglie.
Public prefissoPagine					'prefisso delle pagine di gestione delle categorie
Public blocchiTotali					'se true e categoria bloccata non e possibile immettere sottocategorie o relazioni

Public CategorieAlternative             'indica se sono gestite le categorie alternative abilitando i flag per la divisione in categorie principali ed alternative.

Private oIndex							'oggetto per la gestione dei contenuti e dell'indice generale.

'altre prorieta e variabili interne
Public OrdineLenght						'# max cifre che puo avere il campo ordine 
Private classSalva


'******************************************************************************************************************************************
'******************************************************************************************************************************************
'COSTRUTTORI CLASSE
'******************************************************************************************************************************************

Private Sub Class_Initialize()
	nomeSingolare = ChooseValueByAllLanguages(Session("LINGUA"), "categoria", "category", "", "", "", "", "", "")
	nomePlurale = ChooseValueByAllLanguages(Session("LINGUA"), "categorie", "categories", "", "", "", "", "", "")
	isB2B = false
	OrdineLenght = 4
	abilitaBlocchiEsterni = false
	abilitaLogo = true
	abilitaFoto = true
	abilitaDescrittori = true
	categorieBloccate = false
	GestioneCategorieMiste = false
	prefissoPagine = ""
	blocchiTotali = false
	abilitaPermessiAreaRiservata = false
	attivaCKEditorPerDescrizione = false
	
	Set conn = Server.CreateObject("ADODB.connection")
	conn.open Application("DATA_ConnectionString")
    
    'impostazione dati relazioni
    RelazioniCount = 0
    'impostazione della relazione principale
    RelazioneDescrittori_Table = ""
    RelazioneDescrittori_Pkfield = ""
    RelazioneDescrittori_FkField = ""
End Sub

Private Sub Class_Terminate()
	
End Sub


'******************************************************
'Oggetto di gestione dell'indice
Public Property Get Index()
    set Index = oIndex
End Property

Public Property Let Index(obj)
	if isObjectCreated(obj) then
		set oIndex = obj
		set conn = oIndex.conn
	else
		oIndex = obj
	end if
End Property

'******************************************************************************************************************************************
'******************************************************************************************************************************************
'FUNZIONI INTERNE DELLA CLASSE
'******************************************************************************************************************************************


'.................................................................................................
'funzione che restituisce il codice per filtrare tutti i record discendenti per le categorie base 
'indicate dalla lista contenuta nel relativo parametro.
'   InAnd       parametro che indica se alla condizione SQL deve essere inserito direttamente il
'               l'operatore "AND"
'.................................................................................................
public function SQL_FiltroCategorieBase(InAnd)
    dim sql, ListaCategorie, CatId
    
    ListaCategorie = split(cString(filtroCategorieBase), ",")
    sql = ""
    for each CatId in ListaCategorie
        sql = sql & SQL_IdListSearch(conn, tabella & "." & prefisso & "_tipologie_padre_lista", CatId) & " OR "
    next

    if sql <> "" then
        SQL_FiltroCategorieBase = IIF(InAnd, " AND ", "") & "( " & left(sql, len(sql) - 4) & " ) "
    else
        SQL_FiltroCategorieBase = ""
    end if
end function


'.................................................................................................
'funzione che restituisce il codice per filtrare i record in base alla tipologia di categoria 
'principale o alternativa del padre
'   InAnd       parametro che indica se alla condizione SQL deve essere inserito direttamente il
'               l'operatore "AND"
'.................................................................................................
public function SQL_FiltroCategorieAlternative(Principali, Alternative, SqlPrefix, InAnd)
    dim sql
    if CategorieAlternative then
        sql = SqlPrefix & prefisso & "_tipologia_padre_base IN (SELECT " & prefisso & "_id FROM " & tabella & " WHERE "
        if Principali then
            sql = sql & SQL_IsTrue(conn, prefisso & "_principale")
        end if
        if Alternative then
            sql = sql & IIF(Principali, " OR ", "") & _
                  SQL_IsTrue(conn, prefisso & "_alternativa")
        end if
        sql = sql & ")"
        SQL_FiltroCategorieAlternative = IIF(InAnd, " AND ", "") & sql
    else
        SQL_FiltroCategorieAlternative = ""
    end if
end function


'.................................................................................................
'.. funzione che restituisce il tipo (principale / altrenativa) della categoria indicata dal parametro id
'.................................................................................................
public function GetTipoCategoria(id)
    dim sql, rs, Principale, Alternativa
    if CategorieAlternative then
        set rs = Server.CreateObject("ADODB.RecordSet")
        sql = "SELECT figlio." & prefisso & "_padre_id, " + _
                    " figlio." & prefisso & "_principale AS FiglioPrincipale, " + _
                    " figlio." & prefisso & "_alternativa AS FiglioAlternativa, " + _
                    " padre." & prefisso & "_principale AS PadrePrincipale, " + _
                    " padre." & prefisso & "_alternativa AS PadreAlternativa " + _
              " FROM " & tabella & " figlio LEFT JOIN " & tabella & " padre ON figlio." & prefisso & "_tipologia_padre_base = padre." & prefisso & "_id " + _
              " WHERE figlio." & prefisso & "_id = " & id
        rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
        
        if not rs.eof then
            if cIntero(rs(prefisso & "_padre_id"))>0 then
                'categoria interna
                Principale = rs("PadrePrincipale")
                Alternativa = rs("PadreAlternativa")
            else
                'categoria base
                Principale = rs("FiglioPrincipale")
                Alternativa = rs("FiglioAlternativa")
            end if
        else
            Principale = false
            Alternativa = false
        end if
        
        rs.close
        
        if Principale OR Alternativa then
            GetTipoCategoria = IIF(Principale, "<strong>principale</strong>" + IIF(Alternativa, ", ", ""), "") + _
                               IIF(Alternativa, "alternativa", "")
        else
            GetTipoCategoria = "non impostato"
        end if
        set rs = nothing
    else
        GetTipoCategoria = ""
    end if
end function


'.................................................................................................
'   aggiunge la relazione all'elenco delle relazioni da gestire nei vincoli di integrita' 
'   per impedire la cancellazione dei record delle categorie
'   Table                   Nome della tabella correlata
'   PkField                 Nome del campo chiave primaria della tabella correlata
'   FkField                 Nome del campo chiave esterna della tabella correlata che punta alla tabella delle categorie.
'   RelazioneDescrittori    Indica se la relazione e' collegata con i descrittori e ne vincola il funzionamento
'                           ATTENZIONE: solo una della lista puo' essere la relazione dei descrittori
'.................................................................................................
Public sub AddRelazione(Table, PkField, FkField, RelazioneDescrittori)
    'aggiunge nuova relazione all'array
    RelazioniCount = RelazioniCount + 1
    redim preserve Relazioni(4, RelazioniCount)
    
    'imposta i dati della relazione
    Relazioni(RELAZIONI_POS_TABLE, RelazioniCount-1)      = Table
    Relazioni(RELAZIONI_POS_PK, RelazioniCount-1)         = PkField
    Relazioni(RELAZIONI_POS_FK, RelazioniCount-1)         = FkField
    Relazioni(RELAZIONI_POS_DES, RelazioniCount-1)        = RelazioneDescrittori
    
    'imposta relazione principale che vincola i descrittori
    if RelazioneDescrittori then
        RelazioneDescrittori_Table = Table
        RelazioneDescrittori_Pkfield = PkField
        RelazioneDescrittori_FkField = FkField
    end if
end sub


'.................................................................................................
'   funzione che verifica se almeno una delle relazioni blocca il record
'   id:     id del record della categoria da verificare
'.................................................................................................
Public function ConRelazioni(id)
    dim i, sql, rs
    ConRelazioni = NULL
    if cInteger(id)>0 then
        Set rs = server.CreateObject("ADODB.recordset")
        sql = ""
        'scorre record e verifica relazioni del record
        For i = 0 to RelazioniCount-1
            'cerca relazione per relazione
            sql = " SELECT COUNT(*) FROM "& Relazioni(RELAZIONI_POS_TABLE, i) &" WHERE "& Relazioni(RELAZIONI_POS_FK, i) &"=" & ID
            if cInteger(GetValueList(conn, rs, sql))>0 then
                ConRelazioni = true
                i = RelazioniCount  'alla prima che trova interrompe la ricerca.
            end if
        next
        set rs = nothing
    end if
    if IsNull(ConRelazioni) then
        ConRelazioni = false
    end if
end function


'.................................................................................................
'   funzione che restituisce la condizione WHERE per includere o escludere i le categorie
'   con record relazionati
'   SqlTableName        variabile che indica il nome nella query della tabella delle categorie.
'   ConRelazioni        Variabile che indica se la condizione deve restituire i record con relazioni (true)
'                       o senza relazioni (false)
'.................................................................................................
Public function SQL_Relazioni(SQLTableName, ConRelazioni)
    dim i, sql
    sql = ""
    
    for i = 0 to RelazioniCount-1
        sql = sql & " " & SQLTableName & "."& prefisso &"_id " & IIF(ConRelazioni, "", " NOT ") & _
                    " IN (SELECT " & Relazioni(RELAZIONI_POS_FK, i) & " FROM " & Relazioni(RELAZIONI_POS_TABLE, i) & " ) AND "
    next
    sql = "( " & left(sql, len(sql)-4) & ")"
end function


'******************************************************************************************************************************************
'******************************************************************************************************************************************
'FUNZIONI GENERICHE DELLA CLASSE
'******************************************************************************************************************************************

'.................................................................................................
'..			restituisce gli ID dei discendenti data la tipologia separati da ","
'..			IdList				ID del nodo di partenza
'.................................................................................................
Public Function DiscendentiID(IdList)
	dim sql, Lista, Id
	
	'filtra per "categoria base visualizzabile"
    sql = SQL_FiltroCategorieBase(true)

	If cString(IdList) <> "" then
		Lista = split(cString(IdList), ",")
		
		for each Id in Lista
			id = cIntero(Id)
			if id > 0 then
				sql = sql & " OR (" & SQL_IdListSearch(conn, prefisso & "_tipologie_padre_lista", Id) & _
                    		" AND " & prefisso & "_id<>" & Id &")"
			end if
		next
		
	end if
	
	if sql<>"" then
		'sql = " WHERE " & right(sql, len(sql) - 3)
		sql = " WHERE " & right(sql, len(Trim(sql)) - 3)
	end if
	
	sql = "SELECT " & prefisso & "_id FROM " & tabella & sql
    DiscendentiID = GetValueList(conn, NULL, sql)

End Function


'.................................................................................................
'..			restituisce gli ID delle foglie data la tipologia separati da ","
'..			ID				ID del nodo padre
'.................................................................................................
Public Function FoglieID(ID)
	dim sql
    ID = cIntero(ID)

    'filtra per "categoria base visualizzabile"
    sql = SQL_FiltroCategorieBase(true)
    
    'mostra solo le categorie alle quali e' possibile associare un record
    if not GestioneCategorieMiste then
        sql = sql & " AND " & SQL_IsTrue(conn, prefisso & "_foglia")
    end if
    
    'filtra per nodo, se richiesto
    if cIntero(ID)>0 then
        sql = sql & " AND " & SQL_IdListSearch(conn, prefisso & "_tipologie_padre_lista", ID)
    end if
    
    sql = " SELECT " & prefisso & "_id FROM " & tabella & _
          IIF(sql<>"", " WHERE " & right(sql, len(sql) - 5), "")
          
    FoglieID = GetValueList(conn, NULL, sql)
End Function


'.................................................................................................
'..			restituisce la query sql che genera l'elenco delle categorie con relative sottocategorie
'..			SoloFoglie		se true visualizza solo le categorie "foglie"
'.................................................................................................
Public Function QueryElenco(SoloFoglie, condition)
    QueryElenco = QueryElencoFiltrato(false, false, SoloFoglie, condition)
end function 


'.................................................................................................
'..			restituisce la query sql che genera l'elenco delle categorie con relative sottocategorie
'..			SoloFoglie		        se true visualizza solo le categorie "foglie"
'..         SoloCatPrincipali       se true visualizza solo le categorie principali
'..         SoloCatAlternative      se true visualizza solo le categorie alternative
'.. N.B.:   se entrambi i parametri SoloCatPrincipali e SoloCatAlternative sono a false il filtro viene omesso.
'.................................................................................................
Public Function QueryElencoFiltrato(SoloCatPrincipali, SoloCatAlternative, SoloFoglie, condition)
	dim sql, level, WHERE_sql, rs
	sql = "SELECT "& prefisso &"_livello FROM "& tabella &" GROUP BY "& prefisso &"_livello"
	Set rs = server.CreateObject("ADODB.recordset")
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	
    if SQL_FiltroCategorieBase(true)<>"" AND cString(condition)="" then
        'aggiunge condizione "inutile" per fare in modo che il ciclo interno arrivi ad aggiungere 
        'comunque tutte le condizioni (anche quelle di filtro per categoria base)
        condition = " (1=1) "
    end if
    
    WHERE_sql = ""
	
    'filtro su categorie finale per foglie
	if SoloFoglie then
		WHERE_sql = WHERE_sql & " AND " & SQL_IsTrue(conn, "TIP_L0."& prefisso &"_foglia")
	end if
    'filtro per categorie principali / alternative
    if CategorieAlternative AND (SoloCatPrincipali OR SoloCatAlternative) then
        WHERE_sql = WHERE_sql & _
                    " AND TIP_L0." & prefisso & "_tipologia_padre_base IN (" & _
                    " SELECT " & prefisso & "_id FROM " & tabella & " WHERE "
        if SoloCatPrincipali then
            WHERE_sql = WHERE_sql & SQL_IsTrue(conn, prefisso & "_principale")
        end if
        if SoloCatAlternative then
            WHERE_sql = WHERE_sql & IIF(SoloCatPrincipali, " OR ", "") & _
                        SQL_IsTrue(conn, prefisso & "_alternativa")
        end if
        WHERE_sql = WHERE_sql & ")"
    end if

    sql = ""
	while not rs.eof
		sql = sql & "SELECT TIP_L0."& prefisso &"_id, TIP_L0."& prefisso &"_codice, TIP_L0."& prefisso &"_visibile, TIP_L0."& prefisso &"_livello, (" 
		for level = rs(prefisso &"_livello") to 1 step -1 
			sql = sql & "TIP_L" & level & "."& ChooseValueByAllLanguages(Session("LINGUA"), prefisso & "_nome_it ", prefisso & "_nome_en ", "", "", "", "", "", "") & SQL_concat(conn) & " ' - ' " & SQL_concat(conn)
		next
		sql = sql & " TIP_L0."& prefisso &"_nome_it) AS NAME, (TIP_L0."& prefisso &"_nome_it) AS NAME_it, (TIP_L0."& prefisso &"_nome_en) AS NAME_en, TIP_L0."& prefisso &"_nome_it FROM " & IIF(cInteger(rs(prefisso &"_livello"))>0, String(rs(prefisso &"_livello"), "("), "") & " "& tabella &" TIP_L0 " 
		for level = 1 to rs(prefisso &"_livello")
			sql = sql & " INNER JOIN "& tabella &" TIP_L" & level & " ON TIP_L" & (level-1) & "."& prefisso &"_padre_id = TIP_L" & level & "."& prefisso &"_id ) "
		next
		sql = sql & "WHERE TIP_L0."& prefisso &"_livello=" & rs(prefisso &"_livello") & WHERE_sql
		if condition <> "" then
			sql = sql & " AND ( "
			for level = 0 to rs(prefisso &"_livello")
				if level > 0 then sql = sql & " OR "
				sql = sql & " ( " & replace(" (" & condition & ") " & SQL_FiltroCategorieBase(true), tabella, "TIP_L" & level ) & " ) "
			next
			sql = sql & " ) "
		end if
		rs.movenext
		if not rs.eof then
			sql = sql & " UNION "
		end if
	wend
	if sql <> "" then
		sql = sql & " ORDER BY " & IIF(rs.recordcount > 1, ChooseValueByAllLanguages(Session("LINGUA"), "NAME", "NAME", "", "", "", "", "", ""), ChooseValueByAllLanguages(Session("LINGUA"), prefisso & "_nome_it", prefisso & "_nome_en", "", "", "", "", "", ""))
	else
		sql = " SELECT *, ("& prefisso &"_nome_it) AS NAME FROM "& tabella &" TIP_L0 " + _
              " WHERE (1=1) " & WHERE_sql & SQL_FiltroCategorieBase(true) & _
			  " ORDER BY ("& ChooseValueByAllLanguages(Session("LINGUA"), prefisso & "_nome_it", prefisso & "_nome_en", "", "", "", "", "", "") & ")"
	end if
	rs.close

	QueryElencoFiltrato = sql
'response.write sql
	Set rs = nothing
End Function


'.................................................................................................
'..			restituisce il percorso ed il nome completo della categoria
'..			tip_id 			id della tipologia di cui recuperare il nome
'.................................................................................................
Public Function NomeCompleto(tip_id)
    NomeCompleto = NomeCategoria(tip_id, Session("LINGUA"))
End Function


'..................................................................................................
'..		PER LA PARTE VISIBILE
'..		ritorna il nome completo di percorso della tipologia corrente
'..		tip_id			id della tipologia di cui reperire il percorso ed il nome
'..................................................................................................
Public Function NomeCompletoVisibile(tip_id)
    NomeCompletoVisibile = NomeCategoria(tip_id, Session("LINGUA"))
end function


'..................................................................................................
'..		ritorna il nome completo di percorso della tipologia corrente in base alla lingua
'..		tip_id			id della tipologia di cui reperire il percorso ed il nome
'..     lingua          lingua nella quale estrarre il nome
'..................................................................................................
Private Function NomeCategoria(tip_id, lingua)
    dim rs, sql
	Set rs = server.CreateObject("ADODB.recordset")
    
    NomeCategoria = ""

	sql = "SELECT " & prefisso & "_tipologie_padre_lista FROM "& tabella &" WHERE "& prefisso &"_id=" & tip_id

	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    if not rs.eof then
        sql = "SELECT * FROM "& tabella & _
              " WHERE "& prefisso &"_id IN (" & rs(prefisso & "_tipologie_padre_lista") & ") " & SQL_FiltroCategorieBase(true) & _
              " ORDER BY " & prefisso & "_livello"
        rs.close
        rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        while not rs.eof
            NomeCategoria = IIF(NomeCategoria<>"", NomeCategoria & " - ", "") & CBLL(rs, prefisso &"_nome", lingua)
            rs.movenext
        wend
    end if
    rs.close
	
	set rs = nothing
    
end function


'.................................................................................................
'..		scrive il sistema di input per la selezione di una categoria
'..		FormName		Nome del form in cui viene generato l'input
'..		InputName 		Nome dell'input generato
'..		InputValue		Valore/categoria selezionata
'..		OnlyLeaf		Indica se vengono visualizzate solo le foglie (TRUE) o tutte le categorie
'..		DisplayReduced	Indica se viene visualizzato un input 
'.						ridotto (TRUE, per selezione nei motori di ricerca) o esteso (con link testuali)
'.................................................................................................
Public Sub WritePicker(FormName, InputName, InputValue, OnlyLeaf, DisplayReduced, InputSize)
    CALL WritePickerEx(FormName, InputName, InputValue, OnlyLeaf, DisplayReduced, InputSize, false)
end sub


'.................................................................................................
'..		scrive il sistema di input per la selezione di una categoria
'..		FormName		Nome del form in cui viene generato l'input
'..		InputName 		Nome dell'input generato
'..		InputValue		Valore/categoria selezionata
'..		OnlyLeaf		Indica se vengono visualizzate solo le foglie (TRUE) o tutte le categorie
'..		DisplayReduced	Indica se viene visualizzato un input 
'..						ridotto (TRUE, per selezione nei motori di ricerca) o esteso (con link testuali)
'..     Mandatory       indica se l'input &egrave; obbligatorio o meno, se obbligatorio non visualizza il pulsante di reset
'.................................................................................................
Public Sub WritePickerEx(FormName, InputName, InputValue, OnlyLeaf, DisplayReduced, InputSize, Mandatory)
	dim ViewName, ViewValue, rs
	Set rs = server.CreateObject("ADODB.recordset")
	if cInteger(InputValue)>0 then
		'recupera valore dell'input
		ViewValue = NomeCompleto(InputValue)
	else
		ViewValue = ""
	end if
	ViewName = "view_" & InputName %>
	<input type="hidden" name="<%= InputName %>" value="<%= InputValue %>">
	<table cellpadding="0" cellspacing="0">
		<tr>
			<td <%= IIF(DisplayReduced, " colspan=""2"" ", " style=""padding-top:2px;"" ") %>>
				<input READONLY type="text" name="<%= ViewName %>" id="<%= ViewName %>" value="<%= ViewValue %>" style="padding-left:3px;" size="<%= InputSize %>" title="<%=ChooseValueByAllLanguages(Session("LINGUA"), "Apre l'elenco " & nomePlurale & " per selezionarne una.", "It opens " & nomePlurale & " list to choose one.", "", "", "", "", "", "")%>" onmouseover="<%= FormName %>_<%= InputName %>_UpdateTitle(this)" onclick="<%= FormName %>_<%= InputName %>_ApriFinestra()">
			</td>
		<% if DisplayReduced then %>	
			</tr>
			<tr>
		<% end if %>
			<td style="<%= IIF(DisplayReduced, "width:68%; padding-bottom:2px;", "padding-top:1px;") %>" nowrap>
				<a class="<%= IIF(DisplayReduced, "button_input_bottom", "button_input") %>" id="link_scegli_<%= InputName %>" href="javascript:void(0)" onclick="<%= FormName %>.<%= ViewName %>.onclick();" title="<%=ChooseValueByAllLanguages(Session("LINGUA"), "Apre l'elenco " & nomePlurale & " per selezionarne una.", "It opens " & nomePlurale & " list to choose one.", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %> style="width:auto; display:block;">
					<%= IIF(DisplayReduced, ChooseValueByAllLanguages(Session("LINGUA"), "SCEGLI ", "CHOOSE ", "", "", "", "", "", "") & UCase(nomeSingolare), ChooseValueByAllLanguages(Session("LINGUA"), "SCEGLI ", "CHOOSE ", "", "", "", "", "", "")) %>
				</a>
			</td>
            <% if not Mandatory then %>
    			<td style="<%= IIF(DisplayReduced, "width:32%; padding-bottom:3px;", "padding-top:1px;") %>">
					<a class="<%= IIF(DisplayReduced, "button_input_bottom", "button_input") %>" id="link_reset_<%= InputName %>" style="border-left:0px;" href="javascript:void(0)" onclick="<%= FormName %>.<%= InputName %>.value='';<%= FormName %>.<%= ViewName %>.value='';" title="<%=ChooseValueByAllLanguages(Session("LINGUA"), "cancella la selezione eseguita", "clear selection", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %> style="width:auto; display:block;">RESET</a></td>
            <% else %>
                <td>&nbsp;(*)</td>
            <% end if %>
		</tr>
	</table>
    <script language="JavaScript" type="text/javascript">
        function <%= FormName %>_<%= InputName %>_ApriFinestra(){
            OpenAutoPositionedScrollWindow('<%= prefissoPagine %>CategorieSeleziona.asp?formname=<%= FormName %>&inputname=<%= InputName %>&<%= IIF(OnlyLeaf, "SoloFoglie=1&", "") %>selected=' + <%= FormName %>.<%= InputName %>.value, 'selezione_categoria', 760, 450, true)
            <%= FormName %>_<%= InputName %>_UpdateTitle(<%= FormName %>.<%= ViewName %>)
        }
        
        function <%= FormName %>_<%= InputName %>_UpdateTitle(viewInput){
            viewInput.title = viewInput.value;
        }
        
        <%= FormName %>_<%= InputName %>_UpdateTitle(<%= FormName %>.<%= ViewName %>)
    </script>
	<%
	set rs = nothing
End Sub


'.................................................................................................
'..			procedura che blocca la categoria impostando i parametri per i blocchi esterni
'..			ID 				id della tipologia da bloccare
'..			ExternalSource	sorgente del blocco
'..			ExternalId		id del record sorgente del blocco
'..			Locked			indica se deve essere o meno bloccata anche la gestione categoria
'.................................................................................................
Public Sub Lock(byval ID, byval ExternalSource, byval ExternalId)
	dim sql, rs
	Set rs = server.CreateObject("ADODB.recordset")
	sql = "SELECT * FROM "& tabella &" WHERE "& prefisso &"_id=" & ID
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

	if not rs.eof then
		if cString(rs(prefisso + "_external_source"))="" OR _
		   lCase(cString(ExternalSource)) = lCase(cString(rs(prefisso + "_external_source"))) then
			rs(prefisso + "_external_ID") = IdList_ADD(rs(prefisso + "_external_ID"), ExternalId)
			rs(prefisso + "_external_source") = ExternalSource
			rs.update
		elseif lCase(cString(ExternalSource)) <> lCase(cString(rs(prefisso + "_external_source"))) then
			Session("WARNING") = NomeSingolare + " gi&agrave; collegata!"
		end if
	else
		Session("WARNING") = NomeSingolare + " non trovata!"
	end if
	rs.close

	set rs = nothing
end sub


'.................................................................................................
'..			procedura che sblocca la categoria gestita da blocchi esterni
'..			ExternalSource	sorgente del blocco
'..			ExternalId		id del record sorgente del blocco
'.................................................................................................
Public Sub UnLock(ExternalSource, ExternalId)
	dim sql, rs
	Set rs = server.CreateObject("ADODB.recordset")
	sql = "SELECT * FROM "& tabella &" WHERE " + _
		  prefisso &"_external_source LIKE '" & ExternalSource & "' AND " + _
		  "( " & prefisso &"_external_ID LIKE '" & ExternalId & "' OR " & prefisso &"_external_ID LIKE '%(" & ExternalId & ")%') "
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	if not rs.eof then
		if lCase(cString(ExternalSource)) <> lCase(cString(rs(prefisso + "_external_source"))) then
			Session("WARNING") = NomeSingolare + " collegata con altra sorgente!"
		else
			rs(prefisso + "_external_ID") = IdList_REMOVE(rs(prefisso + "_external_ID"), ExternalId)
			if IsNull(rs(prefisso + "_external_ID")) then
				rs(prefisso + "_external_source") = NULL
			end if
			rs.update
		end if
	else
		Session("WARNING") = NomeSingolare + " non trovata!"
	end if
	rs.close
	
	set rs = nothing
end sub



'******************************************************************************************************************************************
'******************************************************************************************************************************************
'GESTIONE ELENCO CATEGORIE
'******************************************************************************************************************************************


'visualizza l'elenco delle categorie
Public Sub Elenco()
	dim sql, rs, rsc, Pager
	
		set Pager = new PageNavigator
	
	'imposta ricerca
	if Request.ServerVariables("REQUEST_METHOD")="POST" then
		Pager.Reset()
		CALL SearchSession_Reset(prefisso &"_")
		if not(request("tutti")<>"") then
			CALL SearchSession_Set(prefisso &"_")
		end if
	end if
	
	set rs = Server.CreateObject("ADODB.RecordSet")
	set rsc = Server.CreateObject("ADODB.RecordSet")
	
    'imposta eventuale filtro di base per gestione categorie
	sql = SQL_FiltroCategorieBase(true)
		
		
	'filtra per nome
	if Session(prefisso &"_nome")<>"" then
		sql = sql & " AND " & SQL_FullTextSearch(Session(prefisso &"_nome"), FieldLanguageList(prefisso &"_nome_"))
	end if
	
	'filtra per codice
	if Session(prefisso &"_codice")<>"" then
        sql = sql & " AND " & SQL_FullTextSearch(Session(prefisso &"_codice"), prefisso &"_codice")
	end if
	
	'filtra per livello
	if session(prefisso &"_livello")<>"" then
		sql = sql & " AND "& prefisso &"_livello=" & session(prefisso &"_livello")
	end if
	
	'ricerca per stato pubblicazione
	if Session(prefisso &"_visibile")<>"" then
		if not (instr(1, Session(prefisso &"_visibile"), "1", vbTextCompare)>0 AND _
			    instr(1, Session(prefisso &"_visibile"), "0", vbTextCompare)>0 ) then
			if instr(1, Session(prefisso &"_visibile"), "1", vbTextCompare)>0 then
				'visibile
				sql = sql & " AND "& SQL_IsTrue(conn, prefisso &"_visibile")
			elseif instr(1, Session(prefisso &"_visibile"), "0", vbTextCompare)>0 then
				'non visibile
				sql = sql & " AND NOT "& SQL_IsTrue(conn, prefisso &"_visibile")
			end if
		end if
	end if
	
	'ricerca per categorie foglie
	if Session(prefisso &"_posizione")<>"" then
        dim sql_posizione
        sql_posizione = ""
        if instr(1, Session(prefisso &"_posizione"), "0", vbTextCompare)>0 then
            'categorie base "root"
            sql_posizione = sql_posizione + IIF(sql_posizione <> "", " OR ", "") + _
                            prefisso & "_livello = 0"
        end if
        if instr(1, Session(prefisso &"_posizione"), "1", vbTextCompare)>0 then
            'categorie intermedie
            sql_posizione = sql_posizione + IIF(sql_posizione <> "", " OR ", "") + _
                            "( NOT " & SQL_IsTrue(conn, prefisso &"_foglia") & " AND NOT " & SQL_IsNull(conn, prefisso & "_padre_id") & ") "
        end if
        if instr(1, Session(prefisso &"_posizione"), "2", vbTextCompare)>0 then
            'categorie finali "foglie"
            sql_posizione = sql_posizione + IIF(sql_posizione <> "", " OR ", "") + _
                            SQL_IsTrue(conn, prefisso &"_foglia")
        end if
        if sql_posizione <> "" then
            sql = sql & " AND (" & sql_posizione & ") "
        end if
	end if
	
    'ricerca per categorie principali ed alternative
	if Session(prefisso &"_tipo")<>"" then
        sql = sql & SQL_FiltroCategorieAlternative(instr(1, Session(prefisso &"_tipo"), "0", vbTextCompare)>0, _
                                                   instr(1, Session(prefisso &"_tipo"), "1", vbTextCompare)>0, _
                                                   "", true)
	end if
    
	'ricerca per categoria padre
	if Session(prefisso &"_categoria")<>"" then	
        if cInteger(session(prefisso &"_categoria_TipoFiltro"))=1 then
            'recupera solo categorie associate direttamente (sottocategorie)
    		sql = sql & " AND "& prefisso &"_padre_id = " & cInteger(Session(prefisso &"_categoria"))
        else
            'recupera anche le sottocategorie discendenti
            sql = sql & " AND ',' + "& prefisso &"_tipologie_padre_lista + ',' LIKE '%," & cIntero(Session(prefisso &"_categoria")) & ",%' "
        end if
	end if
	
	'ricerca per raggruppamento
	if Session(prefisso &"_raggruppamento") <> "" then
		sql = sql & " AND "& prefisso &"_id IN (SELECT rag_tipologia_id FROM gtb_tipologie_raggruppamenti WHERE " + _
			  		SQL_FullTextSearch(Session(prefisso &"_raggruppamento"), FieldLanguageList("rag_nome_")) + " )"
	end if
	
	'filtro per descrittore
	if session(prefisso &"_descrittore") <> "" then
		sql = sql & " AND "& prefisso &"_id IN (SELECT " & chiaveEsternaRelCaratteristiche & " FROM " & tabellaRelCaratteristiche & " WHERE " & idCarRelCaratteristiche & "=" & session(prefisso &"_descrittore") & ") "
	end if
	
	if isB2B then
		sql = " SELECT *, (SELECT COUNT(*) FROM gtb_tipologie figli WHERE figli.tip_padre_id=gtb_tipologie.tip_id) AS N_FIGLI, " & _
			  " (SELECT COUNT(*) FROM gtb_tipologie_raggruppamenti WHERE rag_tipologia_id=gtb_tipologie.tip_id) AS N_GRUPPI " & _
			  " FROM gtb_tipologie " + _
			  " WHERE (1=1) " + sql + " ORDER BY tip_nome_it"
	else
		sql = " SELECT *, (SELECT COUNT(*) FROM "& tabella &" figli WHERE figli."& prefisso &"_padre_id=" & tabella & "."& prefisso &"_id) AS N_FIGLI, " & _
			  " 0 AS N_GRUPPI " & _
			  " FROM "& tabella & _
			  " WHERE (1=1) " + sql + " ORDER BY "& prefisso &"_nome_it"
	end if
	Session(prefisso + "CATEGORIE_SQL") = sql
	Session("LIVELLO_BASE") = 999999
	CALL Pager.OpenSmartRecordset(conn, rs, sql, 10) %>
	<div id="content">
		<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
		  		<td style="width:27%;" valign="top">
	<!-- BLOCCO DI RICERCA -->
					<form action="" method="post" id="ricerca" name="ricerca">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td>
								<table cellspacing="1" cellpadding="0" class="tabella_madre">
									<caption><%= ChooseValueByAllLanguages(Session("LINGUA"), "Opzioni di ricerca", "Search options", "", "", "", "", "", "")%></caption>
									<tr>
										<td class="footer" colspan="2">
											<input type="submit" name="cerca" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "CERCA", "SEARCH", "", "", "", "", "", "")%>" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "VEDI TUTTI", "VIEW ALL", "", "", "", "", "", "")%>" style="width: 49%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg(prefisso &"_codice") %>><%= ChooseValueByAllLanguages(Session("LINGUA"), "CODICE", "CODE", "", "", "", "", "", "")%></th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_codice" value="<%= TextEncode(session(prefisso &"_codice")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg(prefisso &"_nome") %>><%= ChooseValueByAllLanguages(Session("LINGUA"), "NOME", "NAME", "", "", "", "", "", "")%></th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_nome" value="<%= TextEncode(session(prefisso &"_nome")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg(prefisso &"_visibile") %>><%= ChooseValueByAllLanguages(Session("LINGUA"), "STATO PUBBLICAZIONE", "PUBLICATION", "", "", "", "", "", "")%></th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="checkbox" class="checkbox" name="search_visibile" value="1" <%= chk(instr(1, Session(prefisso &"_visibile"), "1", vbTextCompare)>0) %>>
											<strong><%= ChooseValueByAllLanguages(Session("LINGUA"), "visibile", "visible", "", "", "", "", "", "")%></strong>
										</td>
									</tr>
									<tr>
										<td class="content" colspan="2">
											<input type="checkbox" class="checkbox" name="search_visibile" value="0" <%= chk(instr(1, session(prefisso &"_visibile"), "0", vbTextCompare)>0) %>>
											<%= ChooseValueByAllLanguages(Session("LINGUA"), "non visibile", "invisible", "", "", "", "", "", "")%>
										</td>
									</tr>
                                    <% if CategorieAlternative then %>
                                        <tr><th colspan="2" <%= Search_Bg(prefisso &"_tipo") %>><%= ChooseValueByAllLanguages(Session("LINGUA"), "TIPO " & UCase(nomeSingolare), UCase(nomeSingolare) & " TYPE", "", "", "", "", "", "")%></th></tr>
    									<tr>
    										<td class="content_b" colspan="2">
    											<input type="checkbox" class="checkbox" name="search_tipo" value="0" <%= chk(instr(1, session(prefisso &"_tipo"), "0", vbTextCompare)>0) %>>
    											<%= ChooseValueByAllLanguages(Session("LINGUA"), nomePlurale & " principali", "main " & nomePlurale , "", "", "", "", "", "")%>
    										</td>
    									</tr>
    									<tr>
    										<td class="content" colspan="2">
    											<input type="checkbox" class="checkbox" name="search_tipo" value="1" <%= chk(instr(1, Session(prefisso &"_tipo"), "1", vbTextCompare)>0) %>>
    											<%= ChooseValueByAllLanguages(Session("LINGUA"), nomePlurale & " alternative", "alternative " & nomePlurale , "", "", "", "", "", "")%>
    										</td>
    									</tr>
                                    <% end if %>
									<tr><th colspan="2" <%= Search_Bg(prefisso &"_categoria") %>><%= ChooseValueByAllLanguages(Session("LINGUA"), UCase(nomeSingolare) & " SUPERIORE", "HIGHER " & UCase(nomeSingolare), "", "", "", "", "", "")%></th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WritePicker("ricerca", "search_categoria", session(prefisso &"_categoria"), false, true, 32) %>
										</td>
									</tr>
									<tr>
										<td class="content" colspan="2">
											<input type="radio" class="checkbox" name="search_categoria_TipoFiltro" value="0" <%= chk(cInteger(session(prefisso &"_categoria_TipoFiltro"))=0) %>>
											<%= ChooseValueByAllLanguages(Session("LINGUA"), "associate anche a discendenti", "associated also with descendant" , "", "", "", "", "", "")%>
										</td>
									</tr>
                                    <tr>
										<td class="content" colspan="2">
											<input type="radio" class="checkbox" name="search_categoria_TipoFiltro" value="1" <%= chk(cInteger(session(prefisso &"_categoria_TipoFiltro"))=1) %>>
											<%= ChooseValueByAllLanguages(Session("LINGUA"), "solo direttamente associate", "only directly associated", "", "", "", "", "", "")%>
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg(prefisso &"_posizione") %>><%= ChooseValueByAllLanguages(Session("LINGUA"), "POSIZIONE " & UCase(nomeSingolare), UCase(nomeSingolare) & " POSITION" , "", "", "", "", "", "")%></th></tr><tr>
										<td class="content" colspan="2">
											<input type="checkbox" class="checkbox" name="search_posizione" value="0" <%= chk(instr(1, Session(prefisso &"_posizione"), "0", vbTextCompare)>0) %>>
											<%= ChooseValueByAllLanguages(Session("LINGUA"), nomePlurale & " base", "root " & nomePlurale, "", "", "", "", "", "")%>
										</td>
									</tr>
									<tr>
										<td class="content" colspan="2">
											<input type="checkbox" class="checkbox" name="search_posizione" value="1" <%= chk(instr(1, Session(prefisso &"_posizione"), "1", vbTextCompare)>0) %>>
											<%= ChooseValueByAllLanguages(Session("LINGUA"), nomePlurale & " intermedie", "in-between " & nomePlurale, "", "", "", "", "", "")%>
										</td>
									</tr>
									<tr>
										<td class="content visibile" colspan="2">
											<input type="checkbox" class="checkbox" name="search_posizione" value="2" <%= chk(instr(1, session(prefisso &"_posizione"), "2", vbTextCompare)>0) %>>
											<%= ChooseValueByAllLanguages(Session("LINGUA"), nomePlurale & " finali", "final " & nomePlurale, "", "", "", "", "", "")%>
										</td>
									</tr>
									<tr><th colspan="2" <%= Search_Bg(prefisso &"_livello") %>><%= ChooseValueByAllLanguages(Session("LINGUA"), "LIVELLO " & UCase(nomeSingolare), UCase(nomeSingolare) & " LEVEL", "", "", "", "", "", "")%></th></tr>
									<tr>
										<td class="content" colspan="2">
											<% sql = "SELECT MAX("& prefisso &"_livello) FROM "& tabella
											dim levels, i
											set levels = Server.CreateObject("Scripting.Dictionary")
											CALL levels.Add("0", ChooseValueByAllLanguages(Session("LINGUA"), "Categorie di base", "Root category", "", "", "", "", "", ""))
											for i=1 to cInteger(GetValueList(conn, NULL, sql))
												CALL levels.Add(cString(i), ChooseValueByAllLanguages(Session("LINGUA"), "Livello ", "Level ", "", "", "", "", "", "") & i)
											next
											CALL DropDownDictionary(levels, "search_livello", Session(prefisso &"_livello"), false, "style=""width:100%;""", Session("LINGUA"))%>
										</td>
									</tr>
									<% 	if isB2B then %>
									<tr><th colspan="2" <%= Search_Bg(prefisso &"_raggruppamento") %>><%= ChooseValueByAllLanguages(Session("LINGUA"), "RAGGRUPPAMENTO", "GROUP", "", "", "", "", "", "")%></th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_raggruppamento" value="<%= TextEncode(session(prefisso &"_raggruppamento")) %>" style="width:100%;">
										</td>
									</tr>
									<% 	end if 
									
									if abilitaDescrittori then
										sql = " SELECT TOP 1 * FROM " & tabellaCaratteristiche & " ORDER BY " & nomeCaratteristiche
										if cString(GetValueList(conn, NULL, sql))<>"" then %>
											<tr><th colspan="2" <%= Search_Bg(prefisso &"_descrittore") %>><%= ChooseValueByAllLanguages(Session("LINGUA"), "DESCRITTORE ASSOCIATO", "ASSOCIATE DESCRIBER", "", "", "", "", "", "")%></th></tr>
											<tr>
												<td class="content" colspan="2">
													<% sql = Replace(sql, "TOP 1", "")
													CALL dropDown(conn, sql, idCaratteristiche, nomeCaratteristiche, "search_descrittore", session(prefisso &"_descrittore"), false, " style=""width:100%;""", Session("LINGUA")) %>
												</td>
											</tr>
										<% end if
									end if %>
									<tr>
										<td class="footer" colspan="2">
											<input type="submit" name="cerca" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "CERCA", "SEARCH", "", "", "", "", "", "")%>" class="button" style="width: 49%;">
											<input type="submit" class="button" name="tutti" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "VEDI TUTTI", "VIEW ALL", "", "", "", "", "", "")%>" style="width: 49%;">
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
					</form>
				</td>
				<td width="1%">&nbsp;</td>
				<td valign="top">
	<!-- BLOCCO RISULTATI -->
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
						<caption class="border">
							<%= ChooseValueByAllLanguages(Session("LINGUA"), "Albero " & nomePlurale, nomePlurale & " tree", "", "", "", "", "", "")%>
						</caption>
						<tr>
							<td class="content">
								<%= ChooseValueByAllLanguages(Session("LINGUA"), "Visualizza " & nomePlurale & " nella gerarchia ad albero:", "View " & nomePlurale & " in the hierarchy tree:", "", "", "", "", "", "")%>								
							</td>
							<td class="content_right">
								<a class="button" href="<%= prefissoPagine %>CategorieAlbero.asp" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Apre la visualizzazione ad albero.", "Open tree viewing.", "", "", "", "", "", "")%>">
									<%= ChooseValueByAllLanguages(Session("LINGUA"), "VISUALIZZA COME ALBERO", "TREE VIEWING", "", "", "", "", "", "")%>
								</a>
							</td>
						</tr>
					</table>
					<table cellspacing="1" cellpadding="0" class="tabella_madre">
						<caption>
							<%= ChooseValueByAllLanguages(Session("LINGUA"), "Elenco " & nomePlurale, nomePlurale & " list", "", "", "", "", "", "")%>
						</caption>
						<% 	if not rs.eof then %>
							<tr><th><%= ChooseValueByAllLanguages(Session("LINGUA"), "Trovati n&ordm; " & Pager.recordcount & " record in n&ordm; " & Pager.PageCount & " pagine ", Pager.recordcount & " records found in " & Pager.PageCount & " pages ", "", "", "", "", "", "")%></th></tr>
						<%	rs.AbsolutePage = Pager.PageNo
							dim lock, HasRelazioni
							lock = false
							while not rs.eof and rs.AbsolutePage = Pager.PageNo
                                HasRelazioni = ConRelazioni(rs(prefisso &"_id"))
								if abilitaBlocchiEsterni then
									lock = (CString(rs(prefisso &"_external_id")) <> "")
								end if %>
								<tr>
									<td class="body">
										<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
											<tr>
												<td class="header<%= IIF(rs(prefisso &"_foglia"), " visibile", "") %>" colspan="4" <% if not rs(prefisso &"_visibile") then%>style="font-weight:normal;"<% end if %>>
													<table border="0" cellspacing="0" cellpadding="0" align="right">
														<tr>
															<td style="font-size:1px;">
																<% if IsObject(oIndex) then
																	CALL index.WriteButton(tabella, rs(prefisso &"_id"), POS_ELENCO)
																end if%>
																<a class="button" href="<%= prefissoPagine %>CategorieMod.asp?ID=<%= rs(prefisso &"_id") %>&from=<%= FROM_ELENCO %>">
																	<%= ChooseValueByAllLanguages(Session("LINGUA"), "MODIFICA", "MODIFY", "", "", "", "", "", "")%>
																</a>
																<% if NOT categorieBloccate then
                                                                    if not InIdList(filtroCategorieBase, rs(prefisso &"_id")) then
                                                                        if lock then %>
                                                                            &nbsp;
                                                                            <a class="button_disabled" href="javascript:void(0);" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Impossibile cancellare la " & nomeSingolare & ": gestione da un applicativo esterno", "Unable to delete this " & nomeSingolare & ": it is managed by an external application", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
                                                                                <%= ChooseValueByAllLanguages(Session("LINGUA"), "CANCELLA", "DELETE", "", "", "", "", "", "")%>
                                                                            </a>
					    											    <% elseif HasRelazioni OR rs("N_FIGLI") > 0 OR rs("N_GRUPPI")>0 then %>
                                                                            &nbsp;
                                                                            <a class="button_disabled" href="javascript:void(0);" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Impossibile cancellare la " & nomeSingolare & ": sono presenti sotto" & nomePlurale & " o record associati", "Unable to delete this " & nomeSingolare & ": there are associated sub" & nomePlurale & " or records", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
                                                                                <%= ChooseValueByAllLanguages(Session("LINGUA"), "CANCELLA", "DELETE", "", "", "", "", "", "")%>
                                                                            </a>
                                                                        <% else %>
                                                                            &nbsp;
                                                                            <a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('<%= prefissoPagine %>CATEGORIE','<%= rs(prefisso &"_id") %>');">
                                                                                <%= ChooseValueByAllLanguages(Session("LINGUA"), "CANCELLA", "DELETE", "", "", "", "", "", "")%>
                                                                            </a>
                                                                        <% end if %>
                                                                    <% else %>
                                                                        &nbsp;
                                                                        <a class="button_disabled" href="javascript:void(0);" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Impossibile cancellare la " & nomeSingolare & ": categoria bloccata", "Unable to delete this " & nomeSingolare & ": locked category", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
                                                                            <%= ChooseValueByAllLanguages(Session("LINGUA"), "CANCELLA", "DELETE", "", "", "", "", "", "")%>
                                                                        </a>
                                                                    <% end if
                                                                end if %>
															</td>
														</tr>
													</table>
													<%= CBLE(rs, prefisso &"_nome_it", Session("LINGUA")) %>
												</td>
											</tr>
											<tr>
												<td class="label" style="width:20%;">id:</td>
												<td class="content"><%= rs(prefisso &"_id") %></td>
												<td class="content_right" colspan="2">
													<% if (NOT categorieBloccate OR rs("N_FIGLI") > 0) then
                                                        if lock AND blocchiTotali AND rs("N_FIGLI") = 0 then %>
                                                            <a class="button_l2_disabled" href="javascript:void(0);" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Impossibile creare sotto" & nomePlurale & ": gestione da un applicativo esterno", "Impossible to create sub" & nomePlurale, "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
                                                                <%= ChooseValueByAllLanguages(Session("LINGUA"), "SOTTO", "SUB", "", "", "", "", "", "")%><%= UCase(nomePlurale) %>
                                                            </a>
                                                        <% elseif (not GestioneCategorieMiste AND HasRelazioni) OR rs("N_GRUPPI")>0 then %>
                                                            <a class="button_l2_disabled" href="javascript:void(0);" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Impossibile creare sotto" & nomePlurale & ": sono gi&agrave; presenti dei record associati", "Impossible to create sub" & nomePlurale, "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
        														<%= ChooseValueByAllLanguages(Session("LINGUA"), "SOTTO", "SUB", "", "", "", "", "", "")%><%= UCase(nomePlurale) %>
        													</a>
													    <% else %>
        													<a class="button_l2" href="<%= prefissoPagine %>CategorieSottocategorie.asp?ID=<%= rs(prefisso &"_id") %>" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Apre elenco delle sotto" & nomePlurale, "Open sub" & nomePlurale & " list", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
        														<%= ChooseValueByAllLanguages(Session("LINGUA"), "SOTTO", "SUB", "", "", "", "", "", "")%><%= UCase(nomePlurale) %>
        													</a>
													    <% end if
                                                    end if
                                                    
                                                    if isB2B then
                                                        if rs(prefisso &"_foglia") then %>
        													<a class="button_l2" style="margin-left:4px;" href="<%= prefissoPagine %>CategorieRaggruppamenti.asp?ID=<%= rs(prefisso &"_id") %>" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Apre l'elenco dei raggruppamenti della categoria.", "Open catecategory groups list.", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
        														<%= ChooseValueByAllLanguages(Session("LINGUA"), "RAGGRUPPAMENTI", "GROUPS", "", "", "", "", "", "")%>
        													</a>
													    <% else %>
        													<a class="button_l2_disabled" style="margin-left:4px;" href="javascript:void(0);" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Impossibile creare raggruppamenti: la categoria &egrave; intermedia.", "Impossible to create groups: the categpry is in-between.", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
        														<%= ChooseValueByAllLanguages(Session("LINGUA"), "RAGGRUPPAMENTI", "GROUPS", "", "", "", "", "", "")%>
        													</a>
													    <% end if
                                                    end if
                                                    
                                                    if RelazioniEsterne_Label<>"" AND RelazioniEsterne_Link<>"" then %>
                                                        <a class="button_L2" style="margin-left:4px;" href="<%= RelazioniEsterne_Link %>?ID=<%= rs(prefisso &"_id") %>" 
														   onclick="OpenAutoPositionedScrollWindow('', this.target, 760, 450)"
														   target="RelazioniEsterne_<%= rs(prefisso &"_id") %>" title="Apre <%= lCase(RelazioniEsterne_Label) %>." <%= ACTIVE_STATUS %>>
															<%= RelazioniEsterne_Label %>
														</a>
											        <% end if %>
												</td>
											</tr>
											<tr>
												<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "percorso completo:", "complete path:", "", "", "", "", "", "")%></td>
												<td class="content" colspan="3"><%= NomeCompleto(rs(prefisso &"_id")) %></td>
											</tr>
											<tr>
												<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "codice:", "code:", "", "", "", "", "", "")%></td>
												<td class="content"><%= rs(prefisso &"_codice") %></td>
												<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "livello:", "level:", "", "", "", "", "", "")%></td>
												<td class="content" style="width:30%;">
													<% if rs(prefisso &"_livello")=0 then %>
														<%= ChooseValueByAllLanguages(Session("LINGUA"), nomeSingolare & " di base", "root " & nomeSingolare, "", "", "", "", "", "")%>
													<% elseif rs(prefisso &"_foglia") then %>
														<%= ChooseValueByAllLanguages(Session("LINGUA"), nomeSingolare & " finale - livello: ", "final " & nomeSingolare & " - level: ", "", "", "", "", "", "")%><%= rs(prefisso &"_livello") %>
													<% else %>
														<%= ChooseValueByAllLanguages(Session("LINGUA"), nomeSingolare & " intermedia - livello: ", "in-between " & nomeSingolare & " - level: ", "", "", "", "", "", "")%><%= rs(prefisso &"_livello") %>
													<% end if %>
												</td>
											</tr>
											<tr>
												<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "ordine:", "order:", "", "", "", "", "", "")%></td>
												<td class="content"><%= rs(prefisso &"_ordine") %></td>
												<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "pubblicata:", "published:", "", "", "", "", "", "")%></td>
												<td class="content"><input type="checkbox" class="checkbox" disabled <%= chk(rs(prefisso &"_visibile")) %>></td>
											</tr>
                                            <% if CategorieAlternative then %>
                                                <tr>
                                                    <td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "tipo:", "type:", "", "", "", "", "", "")%></td>
                                                    <td class="content" colspan="3">
                                                        <%= GetTipoCategoria(rs(prefisso & "_id")) %>
                                                    </td>
                                                </tr>
                                            <% end if %>
										</table>
									</td>
								</tr>
								<%  if rs(prefisso &"_livello") < Session("LIVELLO_BASE") then
										Session("LIVELLO_BASE") = rs(prefisso &"_livello")
									end if
									
								rs.moveNext
							wend %>
							<tr>
								<td class="footer" style="border-top:0px; text-align:left;">
									<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
								</td>
							</tr>
						<%else%>
							<tr><td class="noRecords"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Nessun record trovato", "No record found", "", "", "", "", "", "")%></th></tr>
						<% end if %>
					</table>
				</td> 
			</tr>
			<tr><td>&nbsp;</td></tr>
		</table>		
	</div>
	</body>
	</html>
	<% rs.close
	set rs = nothing
	set rsc = nothing
End Sub


'******************************************************************************************************************************************
'******************************************************************************************************************************************
'GESTIONE ALBERO DELLE CATEGORIE
'******************************************************************************************************************************************
'genera albero delle categorie
Public Sub Albero()
    %>
    <div id="content">
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
			<caption class="border">
				<%= ChooseValueByAllLanguages(Session("LINGUA"), "Albero " & nomePlurale, nomePlurale & " tree", "", "", "", "", "", "")%>
			</caption>
			<tr>
				<td class="content">					
					<%= ChooseValueByAllLanguages(Session("LINGUA"), "Visualizza " & nomePlurale & " ad elenco:", nomePlurale & " list wiewing", "", "", "", "", "", "")%>
				</td>
				<td class="content_right">
					<a class="button" href="<%= prefissoPagine %>Categorie.asp" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Apre la visualizzazione come elenco.", "Open list viewing", "", "", "", "", "", "")%>">
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "VISUALIZZA COME ELENCO", "LIST VIEWING", "", "", "", "", "", "")%>
					</a>
				</td>
			</tr>
		</table>
        <% 
		
        dim oTree
        set oTree = new ObjJsTree
        otree.Name = nomePlurale
        oTree.TableCaption = ChooseValueByAllLanguages(Session("LINGUA"), "Albero " & nomePlurale & ":", nomePlurale & " tree:", "", "", "", "", "", "")
        oTree.Title = UCase(nomePlurale)		
		if cString(filtroCategorieBase)<>"" then	
			oTree.LivelloBase = GetValueList(NULL, NULL, "SELECT TOP 1 " & prefisso & "_livello FROM " & tabella & " WHERE " & prefisso & "_id = " & cInteger(filtroCategorieBase))
			
		end if
		
        'aggiunge nodi all'albero
        CALL Albero_Explore(oTree, 0)
       
        if NOT categorieBloccate AND cString(filtroCategorieBase)="" then
            CALL oTree.AddNodeNew(0, 0, prefissoPagine & "CategorieNew.asp?FROM=" & FROM_ALBERO, ChooseValueByAllLanguages(Session("LINGUA"), "NUOVA", "NEW", "", "", "", "", "", ""))
		end if

        CALL oTree.Write()
 
        set oTree = nothing
        %>
	</div>
	</body>
	</html>
    <%
end sub


''''----- DA SPOSTARE
Public Sub ElencoPerSelezione(soloFoglie, urlAction, nameInput)
	dim sql, rs, maxLivello, i, livNextNode, livPrevNode
	Set rs = server.CreateObject("ADODB.recordset")
	maxLivello = cIntero(GetValueList(conn, rs, "SELECT MAX("&prefisso&"_livello) FROM "&tabella))
	
	dim tdLastNode, tdVerticalLine, tdNode, tdEmpty, tdHelp
	tdNode = "<img width=""16"" height=""22"" src=""../../amministrazione/grafica/filemanager/ftv2node.gif"" />"
	tdVerticalLine = "<img width=""16"" height=""22"" src=""../../amministrazione/grafica/filemanager/ftv2vertline.gif"" />"
	tdLastNode = "<img width=""16"" height=""22"" src=""../../amministrazione/grafica/filemanager/ftv2lastnode.gif"" />"
	tdEmpty = "&nbsp;"
	
	tdHelp = "<span style=""background:red; width:100%; display:block;"">&nbsp;</span>"
	%>
	<div id="content">
		<table cellspacing="0" cellpadding="0" class="tabella_madre" style="">
			<caption class="border">
				<%= ChooseValueByAllLanguages(Session("LINGUA"), "Seleziona la "&nomeSingolare&" dall'albero " & nomePlurale, "Select "&nomeSingolare&" from the "&nomePlurale & " tree", "", "", "", "", "", "")%>
			</caption>
			<tr>
				<td>
					<%
					sql = " SELECT * FROM " & tabella & _
						  " ORDER BY "&prefisso&"_ordine_assoluto, "&prefisso&"_nome_"&Session("LINGUA")
					rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
					while not rs.eof
						%>
						<form action="<%=urlAction%>" method="post" style="margin:0px;" name="form_<%=rs(prefisso&"_id")%>">
							<input type="hidden" name="<%=nameInput%>" value="<%=rs(prefisso&"_id")%>">
							<table cellspacing="0" cellpadding="0" class="tabella_madre" style="border:0px; border-top:1px solid #ffffff;">
								<tr>
									<%
									'calcolo il livello del nodo precedente
									if rs.AbsolutePosition > 1 then
										rs.MovePrevious
										livPrevNode = rs(prefisso&"_livello")
										rs.MoveNext
									else
										livPrevNode = 1
									end if
									'calcolo il livello del nodo successivo
									if rs.AbsolutePosition < rs.RecordCount then
										rs.moveNext
										livNextNode = rs(prefisso&"_livello")
										rs.MovePrevious
									else
										livNextNode = rs.RecordCount + 1
									end if
									%>
									<% for i = 0 to rs(prefisso&"_livello") %>
										<td class="content" style="padding:0px; margin:0px; width:16px; height:22p;">
											<% 											
											'if i = rs(prefisso&"_livello") then
											'	if rs(prefisso&"_foglia") and rs(prefisso&"_livello") <> livNextNode then
											'		response.write tdLastNode
											'	else
											'		response.write tdNode
											'	end if
											'else
											'	if i = (rs(prefisso&"_livello") - 1) and rs(prefisso&"_foglia") and _
											'			(rs(prefisso&"_livello") <> livNextNode OR rs(prefisso&"_livello") <> livPrevNode) then
											'		response.write tdEmpty
											'	else
											'		response.write tdVerticalLine
											'	end if
											'end if
											response.write tdEmpty
											%>
										</td>
									<% next %>
									<td class="content">
										<% if rs(prefisso&"_foglia") OR not soloFoglie then %>
											<a href="javascript:void(0)" onclick="form_<%=rs(prefisso&"_id")%>.submit();"><%=rs(prefisso&"_nome_"&Session("LINGUA"))%></a>
										<% else %>
											<span><%=rs(prefisso&"_nome_"&Session("LINGUA"))%></span>
										<% end if %>
									</td>
									<td class="content_right" style="width:80px;">
										<% if rs(prefisso&"_foglia") OR not soloFoglie then %>
											<a class="button" href="javascript:void(0)" title="" onclick="form_<%=rs(prefisso&"_id")%>.submit();">
												<%= ChooseValueByAllLanguages(Session("LINGUA"), "SELEZIONA", "SELECT", "", "", "", "", "", "")%>
											</a>
										<% else %>
											&nbsp;
										<% end if %>
									</td>
								</tr>
							</table>
						</form>
						<%
						rs.moveNext
					wend
					rs.close	
					%>
				</td>
			</tr>
		</table>
	</div>
	</body>
	</html>
    <%
end sub
''''------------------


'scorre l'albero delle categorie
Public sub Albero_Explore(oTree, padre_id)
    dim rs, sql, lock, nome, BaseLink
    BaseLink = prefissoPagine & "CategorieMod.asp?FROM=" & FROM_ALBERO & "&ID="
  
    sql = "SELECT * FROM "& tabella
    if cString(filtroCategorieBase)<>"" AND cIntero(padre_id)=0 then
        'mostra solo i rami abilitati
        sql = sql & " WHERE " & prefisso & "_id IN (" & filtroCategorieBase & ") "
    else
        'mostra tutti i rami
         sql = sql & " WHERE "& SQL_IfIsNull(conn, prefisso &"_padre_id", "0") &"="& padre_id
	end if
    sql = sql + " ORDER BY "& prefisso &"_nome_it"
    set rs = conn.Execute(sql)
 
 
    while not rs.Eof
        nome = JSEncode(rs(prefisso &"_nome_it"), """")
        
        if rs(prefisso & "_foglia") then 
            CALL oTree.AddLeaf(rs(prefisso &"_livello"), nome, ChooseValueByAllLanguages(Session("LINGUA"), "codice: ", "code: ", "", "", "", "", "", "") & rs(prefisso &"_codice") & ChooseValueByAllLanguages(Session("LINGUA"), " - ordine:", " - order:", "", "", "", "", "", "") & rs(prefisso &"_ordine"), BaseLink & rs(prefisso &"_id"))
        else
            CALL oTree.AddNode(rs(prefisso &"_livello"), nome, ChooseValueByAllLanguages(Session("LINGUA"), "codice: ", "code: ", "", "", "", "", "", "") & rs(prefisso &"_codice") & ChooseValueByAllLanguages(Session("LINGUA"), " - ordine:", " - order:", "", "", "", "", "", "") & rs(prefisso &"_ordine"), BaseLink & rs(prefisso &"_id"), rs(prefisso &"_id"))
        end if
        
        if oTree.IsNodeExpanded(rs(prefisso &"_id")) then
            CALL Albero_Explore(oTree, rs(prefisso &"_id"))
        end if

		'inserisco il ramo NUOVO
		lock = false
		if abilitaBlocchiEsterni AND blocchiTotali then
			lock = (CString(rs(prefisso &"_external_id")) <> "")
		end if
        
		if not rs(prefisso &"_foglia") AND NOT categorieBloccate AND NOT lock then 
            CALL oTree.AddNodeNew(rs(prefisso &"_livello") + 1, rs(prefisso &"_id"), prefissoPagine & "CategorieNew.asp?FROM=" & FROM_ALBERO & "&tfn_" & prefisso & "_padre_id=", ChooseValueByAllLanguages(Session("LINGUA"), "NUOVA", "NEW", "", "", "", "", "", ""))
        end if
		
		rs.MoveNext
	wend
	
	set rs = nothing
end sub

'******************************************************************************************************************************************
'******************************************************************************************************************************************
'GESTIONE INSERIMENTO NUOVA CATEGORIA
'******************************************************************************************************************************************

'visualizza il form per creare una nuova categoria
Public Sub Nuova(dicitura)
	if Request.ServerVariables("REQUEST_METHOD")="POST" then
		Server.Execute(prefissoPagine &"CategorieSalva.asp")
	end if

	if not dicitura is nothing then
		dicitura.puls_new = ChooseValueByAllLanguages(Session("LINGUA"), "INDIETRO;", "BACK;", "", "", "", "", "", "") & dicitura.puls_new
		dicitura.link_new = prefissoPagine & IIF(request("FROM")=FROM_ALBERO, "CategorieAlbero.asp", "Categorie.asp") & ";" & dicitura.link_new
		dicitura.scrivi_con_sottosez()
	end if

	dim sql, i, rs, rsc, ordine
	set rs = server.CreateObject("ADODB.recordset")
	set rsc = server.CreateObject("ADODB.recordset") %>
	<div id="content">
		<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_<%= prefisso %>_albero_visibile" value="1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption><%= ChooseValueByAllLanguages(Session("LINGUA"), "Inserimento nuova ", "Add new ", "", "", "", "", "", "")%><%= nomeSingolare %></caption>
			<tr><th colspan="4"><%= ChooseValueByAllLanguages(Session("LINGUA"), "DATI DELLA " & UCase(nomeSingolare), UCase(nomeSingolare) & " INFORMATIONS", "", "", "", "", "", "")%></th></tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
				<% 	if i = 0 then %>
					<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>"><%= ChooseValueByAllLanguages(Session("LINGUA"), "nome:", "name:", "", "", "", "", "", "")%></td>
				<% 	end if %>
					<td class="content" colspan="2">
						<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
						<input type="text" class="text" name="tft_<%= prefisso %>_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_"& prefisso &"_nome_"& Application("LINGUE")(i)) %>" maxlength="50" size="75">
						<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
					</td>
				</tr>
			<%next %>
			<tr>
				<td class="label" nowrap><%= ChooseValueByAllLanguages(Session("LINGUA"), nomeSingolare & " superiore:", "upper " & nomeSingolare, "", "", "", "", "", "")%></td>
				<td class="content" colspan="2">
				<% 	sql = ""
					i = ""
					if NOT GestioneCategorieMiste then
						'sottocategorie creabili solo dove non presenti record correlati
                        dim j
                        sql = ""
                        'esclude record correlati in almeno una categorizzazione
                        sql = SQL_Relazioni("TIP_L0", false)
						i = ChooseValueByAllLanguages(Session("LINGUA"), " di livello intermedio o non associate", " intermediate or non-associated ", "", "", "", "", "", "")
				 	end if
					if abilitaBlocchiEsterni AND blocchiTotali then
						'sottocategorie creabili solo se non gestite esternamente
						if sql <> "" then
							sql = sql &" AND (TIP_L0."& prefisso &"_external_id = '' OR "& SQL_IsNull(conn, "TIP_L0."& prefisso &"_external_id") &")"
							i = i &" o "
						end if
						sql = sql &""
						i = i & ChooseValueByAllLanguages(Session("LINGUA"), "non gestite da applicativi esterni", "not managed by external applications", "", "", "", "", "", "")
					end if
					CALL dropDown(conn, QueryElenco(false, sql), prefisso &"_id", "NAME", "tfn_"& prefisso &"_padre_id", request("tfn_"& prefisso &"_padre_id"), CString(filtroCategorieBase) <> "", IIF(CategorieAlternative, " onClick='GestioneCategorieAlternative();'", ""), Session("LINGUA"))
					if i <> "" then %>
					<br>
					<span class="note">
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "L'elenco contiene solo " & nomePlurale & i & ".", "List contains only " & i & nomePlurale & ".", "", "", "", "", "", "")%>
					</span>
				<% 	end if %>
				</td>
			</tr>
            <% if CategorieAlternative then %>
                <tr>
				    <td class="label" rowspan="2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "tipo:", "type:", "", "", "", "", "", "")%></td> 
                    <td class="content" colspan="2">
                        <input type="checkbox" class="checkbox" value="1" name="chk_<%= prefisso %>_principale" <%= chk(request.servervariables("REQUEST_METHOD")<>"POST" OR request("chk"& prefisso &"_principale")<>"") %>>
                        <%= ChooseValueByAllLanguages(Session("LINGUA"), "categoria principale", "main category", "", "", "", "", "", "")%>
                    </td>
                </tr>
                <tr>
                    <td class="content" colspan="2">
                        <input type="checkbox" class="checkbox" value="1" name="chk_<%= prefisso %>_alternativa" <%= chk(request("chk"& prefisso &"_alternativa")<>"") %>>
                        <%= ChooseValueByAllLanguages(Session("LINGUA"), "categoria alternativa", "alternative category", "", "", "", "", "", "")%>
                    </td>
                </tr>
                <script language="JavaScript" type="text/javascript">
                    function GestioneCategorieAlternative(){
                        DisableControl(form1.chk_<%= prefisso %>_principale, (form1.tfn_<%= prefisso %>_padre_id.selectedIndex!=0));
                        DisableControl(form1.chk_<%= prefisso %>_alternativa, (form1.tfn_<%= prefisso %>_padre_id.selectedIndex!=0));
                    }
                    
                    GestioneCategorieAlternative();
                </script>
            <% end if %>
			<tr>
				<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "codice:", "code:", "", "", "", "", "", "")%></td>
				<td class="content" colspan="2">
					<input type="text" class="text" name="tft_<%= prefisso %>_codice" value="<%= request("tft_"& prefisso &"_codice") %>" size="50">
				</td>
			</tr>
			<tr>
				<td class="label" rowspan="2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "dati pubblicazione:", "pubblication info:", "", "", "", "", "", "")%></td>
				<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "pubblicata:", "published:", "", "", "", "", "", "")%></td>
				<td class="content">
					<input type="radio" class="checkbox" value="1" name="tfn_<%= prefisso %>_visibile" <%= chk(request.servervariables("REQUEST_METHOD")<>"POST" OR cInteger(request("tfn_"& prefisso &"_visibile"))>0) %>>
					<%= ChooseValueByAllLanguages(Session("LINGUA"), "si", "yes", "", "", "", "", "", "")%>
					<input type="radio" class="checkbox" value="0" name="tfn_<%= prefisso %>_visibile" <%= chk(request("tfn_"& prefisso &"_visibile")<>"" AND cInteger(request("tfn_"& prefisso &"_visibile"))=0) %>>
					<%= ChooseValueByAllLanguages(Session("LINGUA"), "no", "no", "", "", "", "", "", "")%>
				</td>
			</tr>
            <% ordine = cIntero(request.form("tfn_"& prefisso &"_ordine"))
            if ordine = 0 then
                sql = "SELECT MAX(" & prefisso & "_ordine) FROM " + tabella
                ordine = cInteger(GetValueList(conn, rs, sql)) + 1
            end if %>
			<tr>
				<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "ordine:", "order:", "", "", "", "", "", "")%></td>
				<td class="content">
					<input type="text" class="text" name="tfn_<%= prefisso %>_ordine" value="<%= ordine %>" maxlength="4" size="4">
				</td>
			</tr>
			<% 	if abilitaLogo then %>
			<tr>
				<td class="label">logo:</td>
				<td class="content" colspan="2">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_"& prefisso &"_logo", request.form("tft_"& prefisso &"_logo"), "width:329px;", FALSE) %>
				</td>
			</tr>
			<% 	end if %>
			<% 	if abilitaFoto then %>
			<tr>
				<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "foto:", "picture:", "", "", "", "", "", "")%></td>
				<td class="content" colspan="2">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_"& prefisso &"_foto", request.form("tft_"& prefisso &"_foto"), "width:329px;", FALSE) %>
				</td>
			</tr>
			<% 	end if %>
			<tr><th colspan="3"><%= ChooseValueByAllLanguages(Session("LINGUA"), "DESCRIZIONE", "DESCRIPTION", "", "", "", "", "", "")%></th></tr>
			<% 
			dim prefissoCampo
			if attivaCKEditorPerDescrizione then
				prefissoCampo = "tfh_"
			else
				prefissoCampo = "tft_"
			end if
			%>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
					<td class="content" colspan="3">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>
								<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<td><textarea style="width:100%;" rows="3" name="<%= prefissoCampo%><%= prefisso %>_descr_<%= Application("LINGUE")(i) %>"><%= request(prefissoCampo & prefisso &"_descr_" & Application("LINGUE")(i)) %></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
				<% if attivaCKEditorPerDescrizione then %>
					<% CALL activateCKEditor(prefissoCampo&prefisso&"_descr_"&Application("LINGUE")(i))%>
				<% end if %>
			<%next %>
			<% 	if abilitaDescrittori then %>
				<tr><th colspan="3"><%= ChooseValueByAllLanguages(Session("LINGUA"), "CARATTERISTICHE ASSOCIATE", "ASSOCIATED FEATURES", "", "", "", "", "", "")%></th></tr>
				<tr>
					<td colspan="3">
						<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
							<tr>
								<td class="label_no_width">
									<%= ChooseValueByAllLanguages(Session("LINGUA"), "L'associazione delle caratteristiche &egrave; possibile dopo aver salvato.", "You can associate features only after saving.", "", "", "", "", "", "")%>									
								</td>
							</tr>
						</table>
					</td>
				</tr>
			<% 	end if %>
			<% 	if abilitaPermessiAreaRiservata then %>
				<tr><th colspan="3"><%= ChooseValueByAllLanguages(Session("LINGUA"), "PERMESSI DELL'AREA RISERVATA", "RESERVED AREA PERMISSIONS", "", "", "", "", "", "")%></th></tr>
				<tr>
					<td colspan="3">
						<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
							<tr>
								<td class="label_no_width">
									<%= ChooseValueByAllLanguages(Session("LINGUA"), "La selezione dei permessi &egrave; possibile dopo aver salvato.", "Permits selection is possible only after saving.", "", "", "", "", "", "")%>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			<% 	end if %>
			<tr>
				<td class="footer" colspan="3">
					<%= ChooseValueByAllLanguages(Session("LINGUA"), "(*) Campi obbligatori.", "(*) Mandatory fields.", "", "", "", "", "", "")%>
				    <% if abilitaDescrittori OR abilitaPermessiAreaRiservata then %>
					    <input type="submit" class="button" name="salva" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "SALVA", "SAVE", "", "", "", "", "", "")%> &gt;&gt;">
				    <% else %>
					    <input type="submit" class="button" name="elenco" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "SALVA", "SAVE", "", "", "", "", "", "")%>">
				    <% end if %>
                    <input type="submit" style="width:22%;" class="button" name="elenco" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "SALVA & TORNA ALL'ELENCO", "SAVE & GO BACK TO THE LIST", "", "", "", "", "", "")%>">
				</td>
			</tr>
		</table>
		&nbsp;
		</form>
	</div>
	</body>
	</html>
	<%set rs = nothing
	set rsc = nothing
End Sub


'******************************************************************************************************************************************
'******************************************************************************************************************************************
'GESTIONE MODIFICA CATEGORIA
'******************************************************************************************************************************************

'visualizza il form per modificare la categoria
Public Sub Modifica(dicitura)

	if Request.ServerVariables("REQUEST_METHOD")="POST" then
		Server.Execute(prefissoPagine &"CategorieSalva.asp")
	end if
		
	dim sql, rs, rsc, i, disabled
	set rs = server.CreateObject("ADODB.recordset")
	set rsc = server.CreateObject("ADODB.recordset")

	if request("goto")<>"" then
		CALL GotoRecord(conn, rs, session(prefisso + "CATEGORIE_SQL"), prefisso &"_id", prefissoPagine &"CategorieMod.asp")
	end if
	
	if IsObject(oIndex) AND cInteger(request("ID"))>0 AND not dicitura is nothing then
		CALL dicitura.InitializeIndex(oIndex, tabella, request("ID"))
	end if
	
	if not dicitura is nothing then
		dicitura.puls_new = ChooseValueByAllLanguages(Session("LINGUA"), "INDIETRO;", "BACK;", "", "", "", "", "", "") & dicitura.puls_new
		dicitura.link_new = prefissoPagine & IIF(request("FROM")=FROM_ALBERO, "CategorieAlbero.asp", "Categorie.asp") & ";" & dicitura.link_new
		dicitura.scrivi_con_sottosez()
	end if

	sql = " SELECT * FROM "& tabella &" t WHERE "& prefisso &"_id="& cIntero(request("ID"))
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if abilitaBlocchiEsterni then
		if CString(rs(prefisso &"_external_id")) <> "" then
			disabled = "disabled"
		end if
	end if %>
	<div id="content">
		<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>	
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td class="caption"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Modifica dati " & nomeSingolare, nomeSingolare & " modify", "", "", "", "", "", "")%></td>
						<td align="right" style="font-size: 1px;">
							<a class="button" href="?FROM=<%= request("FROM") %>&ID=<%= request("ID") %>&goto=PREVIOUS" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), nomeSingolare & " precedente", "previous " & nomeSingolare, "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
								&lt;&lt; <%= ChooseValueByAllLanguages(Session("LINGUA"), "PRECEDENTE", "PREVIOUS", "", "", "", "", "", "")%>
							</a>
							&nbsp;
							<a class="button" href="?FROM=<%= request("FROM") %>&ID=<%= request("ID") %>&goto=NEXT" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), nomeSingolare & " successiva", "next " & nomeSingolare, "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
								<%= ChooseValueByAllLanguages(Session("LINGUA"), "SUCCESSIVA", "NEXT", "", "", "", "", "", "")%> &gt;&gt;
							</a>
						</td>
					</tr>
				</table>
			</caption>
			<tr><th colspan="3"><%= ChooseValueByAllLanguages(Session("LINGUA"), "DATI DELLA " & Ucase(nomeSingolare), UCase(nomeSingolare) & " INFORMATIONS", "", "", "", "", "", "")%></th></tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>"><%= ChooseValueByAllLanguages(Session("LINGUA"), "nome:", "name:", "", "", "", "", "", "")%></td>
			<% 	end if %>
				<td class="content" colspan="2">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" <%= disabled %> class="text" name="tft_<%= prefisso %>_nome_<%= Application("LINGUE")(i) %>" value="<%= CBR(rs, prefisso &"_nome_"& Application("LINGUE")(i), "tft_") %>" maxlength="50" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then
							response.write "(*)"
							if disabled <> "" then %>
					<input type="hidden" name="tft_<%= prefisso %>_nome_it" value="<%= CBR(rs, prefisso &"_nome_it", "tft_") %>" maxlength="50" size="75">
					<%		end if
						end if %>
				</td>
			</tr>
			<%next %>
			<tr>
				<td class="label" nowrap><%= ChooseValueByAllLanguages(Session("LINGUA"), nomesingolare & " superiore:", "upper " & nomeSingolare, "", "", "", "", "", "")%></td>
				<td class="content" colspan="2">
				<% 	sql = ""
					i = ""
					if NOT GestioneCategorieMiste then
						'sottocategorie creabili solo dove non presenti record correlati
                        sql = SQL_Relazioni("TIP_L0", false)
						i = ChooseValueByAllLanguages(Session("LINGUA"), " di livello intermedio o non associate", " intermediate or non-associated ", "", "", "", "", "", "")
				 	end if
					if abilitaBlocchiEsterni AND blocchiTotali then
						'sottocategorie creabili solo se non gestite esternamente
						if sql <> "" then
							sql = sql &" AND (TIP_L0."& prefisso &"_external_id = '' OR "& SQL_IsNull(conn, "TIP_L0."& prefisso &"_external_id") &")"
							i = i &" o "
						end if
						sql = sql &""
						i = i & ChooseValueByAllLanguages(Session("LINGUA"), "non gestite da applicativi esterni", "not managed by external applications", "", "", "", "", "", "")
					end if
					
					'tolgo la categoria attuale dalla lista di padri
					if sql <> "" then
						sql = sql &" AND "
					end if
					
					sql = sql &"NOT "&  SQL_IdListSearch(conn, "TIP_L0."& prefisso & "_tipologie_padre_lista", cIntero(request("ID")))
					
					
					if Session("LIVELLO_BASE") = rs(prefisso & "_livello") AND filtroCategorieBase<>"" then
						response.write GetValueList(conn, NULL, "SELECT " & prefisso & "_nome_it FROM " & tabella & " WHERE " & prefisso & "_id = (SELECT " & prefisso & "_padre_id FROM " & tabella & " WHERE " & prefisso & "_id = " & cIntero(request("ID")) & ")")
						%>
						<input type="hidden" name="tfn_<%= prefisso %>_padre_id" value="<%= rs(prefisso & "_padre_id") %>" maxlength="50" size="75">
						<%
					else
						CALL dropDown(conn, QueryElenco(false, sql), prefisso &"_id", "NAME", "tfn_"& prefisso &"_padre_id", _
									  CBR(rs, prefisso &"_padre_id", "tfn_"), IIF(filtroCategorieBase<>"", true, false), _
									  disabled & IIF(CategorieAlternative, " onClick='GestioneCategorieAlternative();'", "") & IIF(abilitaPermessiAreaRiservata, "onchange=form1.submit();", ""), _
									  Session("LINGUA"))
					end if
					
					if i <> "" then %>
					<br>
					<span class="note">
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "L'elenco contiene solo " & i & nomePlurale & ".", "List contains only " & i & nomePlurale & ".", "", "", "", "", "", "")%>
					</span>
				<% 	end if %>
				</td>
			</tr>
            <% if CategorieAlternative then %>
                <tr>
				    <td class="label" rowspan="2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "tipo:", "type:", "", "", "", "", "", "")%></td> 
                    <td class="content" colspan="2">
                        <input type="checkbox" class="checkbox" value="1" name="chk_<%= prefisso %>_principale" <%= chk(CBR(rs, prefisso & "_principale", "chk_")) %>>
                        <%= ChooseValueByAllLanguages(Session("LINGUA"), "categoria principale", "main category", "", "", "", "", "", "")%>
                    </td>
                </tr>
                <tr>
                    <td class="content" colspan="2">
                        <input type="checkbox" class="checkbox" value="1" name="chk_<%= prefisso %>_alternativa" <%= chk(CBR(rs, prefisso & "_alternativa", "chk_")) %>>
                        <%= ChooseValueByAllLanguages(Session("LINGUA"), "categoria alternativa", "alternative category", "", "", "", "", "", "")%>
                    </td>
                </tr>
                <script language="JavaScript" type="text/javascript">
                    function GestioneCategorieAlternative(){
                        DisableControl(form1.chk_<%= prefisso %>_principale, (form1.tfn_<%= prefisso %>_padre_id.selectedIndex!=0));
                        DisableControl(form1.chk_<%= prefisso %>_alternativa, (form1.tfn_<%= prefisso %>_padre_id.selectedIndex!=0));
                    }
                    
                    GestioneCategorieAlternative();
                </script>
            <% end if %>
			<tr>
				<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "codice:", "code:", "", "", "", "", "", "")%></td>
				<td class="content" colspan="2">
					<input type="text" <%= disabled %> class="text" name="tft_<%= prefisso %>_codice" value="<%= CBR(rs, prefisso &"_codice", "tft_") %>" size="50">
				</td>
			</tr>
			<tr>
				<td class="label" rowspan="2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "dati pubblicazione:", "pubblication info:", "", "", "", "", "", "")%></td>
				<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "pubblicata:", "published:", "", "", "", "", "", "")%></td>
				<td class="content">
					<input type="radio" <%= disabled %> class="checkbox" value="1" name="tfn_<%= prefisso %>_visibile" <%= chk(CBRV(rs(prefisso &"_visibile"), request.form("tfn_"& prefisso &"_visibile") = "1")) %>>
					<%= ChooseValueByAllLanguages(Session("LINGUA"), "si", "yes", "", "", "", "", "", "")%>
					<input type="radio" <%= disabled %> class="checkbox" value="0" name="tfn_<%= prefisso %>_visibile" <%= chk(CBRV(not rs(prefisso &"_visibile") OR IsNull(rs(prefisso &"_visibile")), request.form("tfn_"& prefisso &"_visibile") = "0")) %>>
					<%= ChooseValueByAllLanguages(Session("LINGUA"), "no", "no", "", "", "", "", "", "")%>
				</td>
			</tr>
			<tr>
				<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "ordine:", "order:", "", "", "", "", "", "")%></td>
				<td class="content">
					<input type="text" <%= disabled %> class="text" name="tfn_<%= prefisso %>_ordine" value="<%= CBR(rs, prefisso &"_ordine", "tfn_") %>" maxlength="4" size="4">
				</td>
			</tr>
			<% 	if abilitaLogo then %>
			<tr>
				<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "logo:", "logo:", "", "", "", "", "", "")%></td>
				<td class="content" colspan="2">
				<% 	if disabled = "" then
						CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_"& prefisso &"_logo", CBR(rs, prefisso &"_logo", "tft_"), "width:329px;", FALSE)
					else %>
					<input type="text" disabled class="text" name="tft_<%= prefisso %>_logo" value="<%= rs(prefisso &"_logo") %>" style="width: 50%;">
				<% 	end if %>
				</td>
			</tr>
			<% 	end if %>
			<% 	if abilitaFoto then %>
			<tr>
				<td class="label"><%= ChooseValueByAllLanguages(Session("LINGUA"), "foto:", "picture:", "", "", "", "", "", "")%></td>
				<td class="content" colspan="2">
				<% 	if disabled = "" then
						CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_"& prefisso &"_foto", CBR(rs, prefisso &"_foto", "tft_"), "width:329px;", FALSE)
					else %>
					<input type="text" disabled class="text" name="tft_<%= prefisso %>_foto" value="<%= rs(prefisso &"_foto") %>" style="width: 50%;">
				<% 	end if %>
				</td>
			</tr>
			<% 	end if %>
			<tr><th colspan="3"><%= ChooseValueByAllLanguages(Session("LINGUA"), "DESCRIZIONE", "DESCRIPTION", "", "", "", "", "", "")%></th></tr>
			<% 
			dim prefissoCampo
			if attivaCKEditorPerDescrizione then
				prefissoCampo = "tfh_"
			else
				prefissoCampo = "tft_"
			end if
			%>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
					<td class="content" colspan="3">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>
								<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<%
								dim textAreaValue
								textAreaValue = CBR(rs, prefisso &"_descr_" & Application("LINGUE")(i), prefissoCampo)
								if attivaCKEditorPerDescrizione then
									textAreaValue = MakeAbsoluteLink(textAreaValue)
								end if
								%>
								<td><textarea style="width:100%;" <%= disabled %> rows="3" name="<%=prefissoCampo%><%= prefisso %>_descr_<%= Application("LINGUE")(i) %>"><%= textAreaValue %></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
				<% if attivaCKEditorPerDescrizione then %>
					<% CALL activateCKEditor(prefissoCampo&prefisso&"_descr_"&Application("LINGUE")(i))%>
				<% end if %>
			<%next
            
            if abilitaDescrittori then %>
    			<tr><th colspan="3"><%= ChooseValueByAllLanguages(Session("LINGUA"), "CARATTERISTICHE ASSOCIATE", "ASSOCIATED FEATURES", "", "", "", "", "", "")%></th></tr>
    			<tr>
    				<td colspan="3">
    				<%	dim value, gruppo
    					gruppo = ""
    					sql = " SELECT *, (SELECT COUNT("& RelazioneDescrittori_Pkfield &") FROM "& tabellaRelCorCaratteristiche &" INNER JOIN "& RelazioneDescrittori_Table &" ON "& tabellaRelCorCaratteristiche &"."& idArtRelCorCaratteristiche &" = "& RelazioneDescrittori_Table &"."& RelazioneDescrittori_Pkfield &" " + _
    						  "		 	   WHERE "& RelazioneDescrittori_Table &"."& RelazioneDescrittori_FkField &" = " & rs(prefisso &"_id") & " AND "& tabellaRelCorCaratteristiche &"."& idCarRelCorCaratteristiche &"= descrittori."& idCaratteristiche &") AS N_ARTICOLI " + _
    						  " FROM ("& tabellaCaratteristiche &" descrittori "& _
    						  " LEFT JOIN "& tabellaRelCaratteristiche &" relazione ON (descrittori."& idCaratteristiche &" = relazione."& idCarRelCaratteristiche &" AND relazione."& chiaveEsternaRelCaratteristiche &" = "& rs(prefisso &"_id") &"))"
    					if tabellaGruppiCaratteristiche <> "" then
    						sql = sql &" LEFT JOIN "& tabellaGruppiCaratteristiche &" gruppi ON descrittori."& idRelGruppiCaratteristiche &" = gruppi."& idGruppiCaratteristiche & _
    							  	   " ORDER BY "& ordineGruppiCaratteristiche & ", " & NomeGruppiCaratteristiche & ", " & IdGruppiCaratteristiche & ", " & nomeCaratteristiche
    					else
    						sql = sql &" ORDER BY "& nomeCaratteristiche
    					end if

    					rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText 
						%>
    					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
    						<% if rsc.eof then %>
    							<tr>
    								<td class="label_no_width">
    									<%= ChooseValueByAllLanguages(Session("LINGUA"), "Nessuna caratteristica trovata", "No features found", "", "", "", "", "", "")%>
    								</td>
    							</tr>
    						<% else %>
    							<tr>
    								<th class="l2_center" width="6%"><%= ChooseValueByAllLanguages(Session("LINGUA"), "associa", "connect", "", "", "", "", "", "")%></th>
    								<th class="l2_center" width="7%"><%= ChooseValueByAllLanguages(Session("LINGUA"), "ordine", "order", "", "", "", "", "", "")%></th>
    								<th class="L2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "caratteristica", "feature", "", "", "", "", "", "")%></th>
    								<th class="L2" width="20%"><%= ChooseValueByAllLanguages(Session("LINGUA"), "tipo di dato", "data type", "", "", "", "", "", "")%></th>
    							</tr>
    							<% while not rsc.eof
    									if tabellaGruppiCaratteristiche <> "" then
    										if gruppo <> CString(rsc(nomeGruppiCaratteristiche)) then
    											gruppo = rsc(nomeGruppiCaratteristiche) %>
    							<tr>
    								<th class="l2" colspan="4"><%= gruppo %></th>
    							</tr>
    							<%			end if
    									end if %>
    								<tr>
    									<td class="content_center">
    										<% 	disabled = cInteger(rsc("N_ARTICOLI"))>0
    											if abilitaBlocchiEsterni then
    												disabled = disabled OR rsc(lockedRelCaratteristiche)
    											end if
    											if disabled then
    												value = true%>
    											<input type="checkbox" checked class="checked" id="caratteristiche_associate_<%= rsc(idCaratteristiche) %>" disabled onclick="set_state_<%= rsc(idCaratteristiche) %>(this)" title="<%= IIF(cInteger(rsc("N_ARTICOLI"))>0, ChooseValueByAllLanguages(Session("LINGUA"), "Sono presenti valori per questa caratteristica negli articoli della categoria.", "There are some values in the category articles for this features", "", "", "", "", "", ""), ChooseValueByAllLanguages(Session("LINGUA"), "Il descrittore &egrave; gestito da un applicativo esterno.", "The descriptor is managed by an external application", "", "", "", "", "", "")) %>">
    											<input type="hidden" <% if abilitaBlocchiEsterni then response.write Disable(rsc(lockedRelCaratteristiche)) end if %> name="caratteristiche_associate" value=" <%= rsc(idCaratteristiche) %> ">
    										<% 	else
    												value = CBRV(not IsNull(rsc(chiaveEsternaRelCaratteristiche)), InStr(request.form("caratteristiche_associate"), " "& rsc(idCaratteristiche)) > 0) %>
    											<input type="checkbox" name="caratteristiche_associate" id="caratteristiche_associate_<%= rsc(idCaratteristiche) %>" value=" <%= rsc(idCaratteristiche) %> " <%= chk(value) %> class="<%= IIF(value, "checked", "checkbox") %>" onclick="set_state_<%= rsc(idCaratteristiche) %>(this)">
    										<% 	end if %>
    									</td>
    									<td class="content_center"><input <% if abilitaBlocchiEsterni then if rsc(lockedRelCaratteristiche) then response.write Disable(rsc(lockedRelCaratteristiche)) & "disabled class=""text_disabled""" else response.write "class=""text""" end if else response.write "class=""text""" end if %> type="text" size="2" name="rel_ordine_<%= rsc(idCaratteristiche) %>" value="<%= rsc(ordineRelCaratteristiche) %>"></td>
    									<td class="content"><%= rsc(nomeCaratteristiche) %></td>
    									<td class="content"><%= DesVisTipo(rsc(tipoCaratteristiche)) %></td>
    								</tr>
    								<script language="JavaScript" type="text/javascript">
    									function set_state_<%= rsc(idCaratteristiche) %>(chk){
    										EnableIfChecked(chk, form1.rel_ordine_<%= rsc(idCaratteristiche) %>);
    										if (chk.checked){
    											form1.rel_ordine_<%= rsc(idCaratteristiche) %>.title = "<%= ChooseValueByAllLanguages(Session("LINGUA"), "Inserisci l'ordine di visualizzazione nella scheda", "Insert visualization order in the form", "", "", "", "", "", "")%>";
    										}
    										else{
    											form1.rel_ordine_<%= rsc(idCaratteristiche) %>.title = "<%= ChooseValueByAllLanguages(Session("LINGUA"), "Selezionare il flag di associazione prima di inserire l'ordine di visualizzazione nella scheda", "Select the flag before inserting visualization order in the form", "", "", "", "", "", "")%>";
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
			<% end if %>
			
			<% 	if abilitaPermessiAreaRiservata then %>
				<tr><th colspan="3"><%= ChooseValueByAllLanguages(Session("LINGUA"), "PERMESSI DELL'AREA RISERVATA", "RESERVED AREA PERMISSIONS", "", "", "", "", "", "")%></th></tr>
				<tr>
					<td colspan="3">
						<script language="JavaScript">
			    			function tutti() {
			    				for(var i=0; i < form1.elements.length; i++)
			    					if (form1.elements(i).id.substring(0, 5) == "prmR_")
			    						form1.elements(i).checked = true
			    			}
			    	
			    			function nessuno() {
			    				for(var i=0; i < form1.elements.length; i++)
			    					if (form1.elements(i).id.substring(0, 5) == "prmR_")
			    						form1.elements(i).checked = false
			    			}
			    		</script>
						<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
							<%	dim listaPadri
								listaPadri = CIntero(CBR(rs, prefisso &"_padre_id", "tfn_"))
								if listaPadri > 0 then
									sql = " SELECT "& prefisso &"_tipologie_padre_lista FROM "& tabella &" WHERE "& prefisso &"_id = "& listaPadri
									listaPadri = CIntero(GetValueList(conn, rsc, sql))
								end if
								sql = " SELECT *, (SELECT COUNT(*) FROM "& tabellaRelUtenti &" WHERE "& idUtenteRelUtenti &" = ut_id AND "& idCategoriaRelUtenti &" IN ("& listaPadri &")) AS N_PRM"& _
									  " FROM (tb_utenti u"& _
									  " INNER JOIN tb_indirizzario i ON u.ut_NextCom_ID = i.IDElencoIndirizzi)"& _
									  " LEFT JOIN "& tabellaRelUtenti &" r ON (u.ut_id = r."& idUtenteRelUtenti & _
									  "		AND r."& idCategoriaRelUtenti &" = "& cIntero(request("ID")) &")"& _
									  " ORDER BY modoRegistra"
								rsc.open sql, conn, adOpenStatic, adLockOptimistic
								if rsc.eof then %>
							<tr><td class="label_no_width"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Nessun utente trovato", "No user found", "", "", "", "", "", "")%></td></tr>
							<%	else %>
							<tr>
	    						<td>
	    							<table cellpadding="0" cellspacing="0" width="100%">
	    								<tr>
	    									<td class="label" nowrap><%= ChooseValueByAllLanguages(Session("LINGUA"), "Elenco degli utenti", "User list", "", "", "", "", "", "")%></td>
	    									<td class="content_right" style="font-size: 1px; padding-right:1px;">
	    										<a id="tutti" class="button_L2" href="javascript:void(0);" onclick="tutti()" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "seleziona tutti gli utenti elencati", "select all listed user", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
	    											<%= ChooseValueByAllLanguages(Session("LINGUA"), "ABILITA TUTTI", "ENABLE ALL", "", "", "", "", "", "")%>
	    										</a>
	    										&nbsp;
	    										<a id="nessuno" class="button_L2" href="javascript:void(0);" onclick="nessuno()" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "toglie la selezione a tutti gli utenti elencati", "deselect all listed user", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
	    											<%= ChooseValueByAllLanguages(Session("LINGUA"), "DISABILITA TUTTI", "DISABLE ALL", "", "", "", "", "", "")%>
	    										</a>
	    									</td>
	    								</tr>
	    							</table>
	    						</td>
	    					</tr>
							<tr>
								<td>
									<table cellpadding="0" cellspacing="1" align="left">
		                            	<tr>
		                                	<th class="l2_center" style="width:1%;"><%= ChooseValueByAllLanguages(Session("LINGUA"), "abilita", "enable", "", "", "", "", "", "")%></th>
							              	<th class="l2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "cognome e nome", "surname and name", "", "", "", "", "", "")%></th>
		                            	</tr>
							<%		while not rsc.eof %>
										<tr>
			        						<td class="content_center">
							<%			if rsc("N_PRM") > 0 then %>
												<input type="checkbox" disabled class="checkbox" checked>
							<%			else
											i = CBRV(NOT IsNull(rsc(idCategoriaRelUtenti)), request("prmR_"& rsc("ut_id")) <> "") %>
			        							<input type="checkbox" readonly class="<%= IIF(i, "checked", "checkbox") %>" value="1"
													   id="prmR_<%= rsc("ut_ID") %>" name="prmR_<%= rsc("ut_ID") %>" <%= Chk(i) %>>
							<%			end if %>
			        						</td>
			        						<td class="label"><%= ContactName(rsc) %></td>
			        					</tr>
							<%			rsc.movenext
									wend %>
									</table>
								</td>
							</tr>
							<%	end if
								rsc.close %>
						</table>
					</td>
				</tr>
			<% 	end if %>
			
			<tr>
				<td class="footer" colspan="3">
					<input type="submit" style="float:left;" class="button_L2" name="propaga_descrittori_nei_figli" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "ASSOCIA QUESTI DESCRITTORI ALLE CATEGORIE FIGLIE", "", "", "", "", "", "", "")%>">
					
					<%= ChooseValueByAllLanguages(Session("LINGUA"), "(*) Campi obbligatori.", "(*) Mandatory fields.", "", "", "", "", "", "")%>
					<input type="submit" class="button" name="salva" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "SALVA", "SAVE", "", "", "", "", "", "")%>">
					<input type="submit" style="width:24%;" class="button" name="elenco" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "SALVA & TORNA ALL'ELENCO", "SAVE & GO BACK TO THE LIST", "", "", "", "", "", "")%>">
				</td>
			</tr>
		</table>
		&nbsp;
		</form>
	</div>
	</body>
	</html>
	<%set rs = nothing
	set rsc = nothing
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
		FormatOrdine = String(OrdineLenght - Len(FormatOrdine), "0") & FormatOrdine
	end if
End Function

'definizione eventuali operazioni su relazioni	
Public Sub Gestione_Relazioni_record(rs, ID)
	dim livello, sql, padre, descr, i, rs_cat, rs_desc
	
	if request.form("propaga_descrittori_nei_figli") <> "" AND abilitaDescrittori then
		'elenco categorie figlie di quella selezionata
		sql = " SELECT "&prefisso&"_id FROM "&tabella&" WHERE ',' + "&prefisso&"_tipologie_padre_lista + ',' LIKE '%,"&ID&",%'" & _
			  " AND "&prefisso&"_livello > (SELECT "&prefisso&"_livello FROM "&tabella&" WHERE "&prefisso&"_id = "&ID&")"
		set rs_cat = server.CreateObject("ADODB.recordset")
		rs_cat.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		'elenco descrittori collegati alla categoria selezionata
		sql = " SELECT DISTINCT "&idCarRelCaratteristiche&","&ordineRelCaratteristiche&_
			  " FROM "&tabellaRelCaratteristiche&" WHERE "&chiaveEsternaRelCaratteristiche&" = "&ID
		set rs_desc = server.CreateObject("ADODB.recordset")
		rs_desc.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		if not rs_cat.eof AND not rs_desc.eof then
			while not rs_cat.eof
				while not rs_desc.eof
					'controllo che non ci sia già l'associazione...
					sql = " SELECT TOP 1 "&idCarRelCaratteristiche&" FROM "&tabellaRelCaratteristiche&" WHERE "&idCarRelCaratteristiche&" = "&rs_desc(idCarRelCaratteristiche)&_
						  " AND "&chiaveEsternaRelCaratteristiche&" = "&rs_cat(prefisso&"_id")
					if cIntero(GetValueList(conn, NULL, sql)) = 0 then
						'...se non c'è la aggiungo
						sql = " INSERT INTO "&tabellaRelCaratteristiche&"("&idCarRelCaratteristiche&","&chiaveEsternaRelCaratteristiche&","&ordineRelCaratteristiche&")"&_
							  "	VALUES("&rs_desc(idCarRelCaratteristiche)&","&rs_cat(prefisso&"_id")&","&rs_desc(ordineRelCaratteristiche)&")"
						'response.write sql & "<br>"
						CALL conn.execute(sql, , adExecuteNoRecords)
					end if
					rs_desc.moveNext
				wend
				rs_desc.moveFirst
				rs_cat.moveNext
			wend
		end if
		rs_desc.close
		rs_cat.close
		set rs_desc = nothing
		set rs_cat = nothing
	end if
	
	'controllo lunghezza campo ordine
	if Len(Replace(request.form("tfn_"& prefisso &"_ordine"), " ", "")) > OrdineLenght then
		session("ERRORE") = "La lunghezza massima del campo ordine &egrave; di "& OrdineLenght &" caratteri"
		Exit Sub
	end if
	
	CALL operazioni_ricorsive_tipologia(ID)
	
	if abilitaDescrittori then
		'GESTIONE CARATTERISTICHE TECNICHE
		if request("ID")<>"" then
			dim CaratList, Carat
			
			'cancella relazioni precedenti
			sql = "DELETE FROM "& tabellaRelCaratteristiche &" WHERE "& chiaveEsternaRelCaratteristiche &"=" & cIntero(ID)
			if abilitaBlocchiEsterni then
				sql = sql &" AND NOT "& SQL_IsTrue(conn, lockedRelCaratteristiche)
			end if
			CALL conn.execute(sql, 0, adExecuteNoRecords)
			
			'gestione categorie associate
			CaratList = split(replace(request("caratteristiche_associate"), " ", ""), ",")
			
			for each Carat in CaratList
				'recupera dati ordine
				sql = "INSERT INTO "& tabellaRelCaratteristiche &"("& idCarRelCaratteristiche &", "& chiaveEsternaRelCaratteristiche &", "& ordineRelCaratteristiche &") " + _
					  " VALUES (" & cIntero(Carat) & ", " & cIntero(ID) & ", " & cIntero(request("rel_ordine_" & Carat)) & ")"
				CALL conn.execute(sql, , adExecuteNoRecords)
			next
		end if
	end if
	
	if abilitaPermessiAreaRiservata then
		dim prm
		sql = "DELETE FROM "& tabellaRelUtenti &" WHERE "& idCategoriaRelUtenti &" = "& ID
		conn.Execute(sql)
		
		sql = " INSERT INTO "& tabellaRelUtenti &"("& idCategoriaRelUtenti &", "& idUtenteRelUtenti &")"& _
			  " VALUES ("& cIntero(ID) &", "
		for each prm in request.form
			if Left(prm, 5) = "prmR_" then
				if request.form(prm) <> "" then
					i = Right(prm, Len(prm) - 5)
					conn.Execute(sql & cIntero(i) & ")")
				end if
			end if
		next
	end if
	
'..............................................................................
	'sincronizzazione con i contenuti e l'indice
	if IsObject(oIndex) then
		CALL Index_UpdateItem(conn, tabella, ID, false)
	end if
'..............................................................................
	
	'imposta parametri per passare alla pagina successiva
	classSalva.isReport = false
	if request.form("elenco")<>"" then
		if request("FROM") = FROM_ALBERO then
			classSalva.Next_Page = prefissoPagine &"CategorieAlbero.asp"
		else
			classSalva.Next_Page = prefissoPagine &"Categorie.asp"
		end if
	else
		classSalva.Next_Page = prefissoPagine &"CategorieMod.asp?ID=" & ID & "&FROM=" & request("FROM")
	end if
End Sub


'funzione che esegue tutte le operazioni ricorsive sulle categorie e relative foglie
Public Sub operazioni_ricorsive_tipologia(tip_id)
	dim rs, sql, tip_padre_id, tip_ordine
	set rs = server.CreateObject("ADODB.recordset")

	'leggo nodo corrente
	'..........................................................................
	sql = "SELECT * FROM "& tabella &" WHERE "& prefisso &"_id=" & tip_id
	rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdtext

	tip_ordine = FormatOrdine(rs(prefisso &"_ordine"))
	if tip_ordine<>"" then
		tip_ordine = tip_ordine & "-"
	end if
	tip_padre_id = cInteger(rs(prefisso &"_padre_id"))
	
	if rs(prefisso &"_padre_id")=0 then
		rs(prefisso &"_padre_id") = NULL
		rs.update
	end if
	
	rs.close
	
	'legge dati del nodo padre
	'..........................................................................
	sql = "SELECT * FROM "& tabella &" WHERE "& prefisso &"_id=" & tip_padre_id
	rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdtext

	sql = "UPDATE "& tabella &" SET "
	if rs.eof then
		'il nodo corrente e' di livello 0, nessun padre
		sql = sql + prefisso &"_livello=0, " + _
					prefisso &"_tipologia_padre_base=" & cIntero(tip_id) & ", " + _
					prefisso &"_albero_visibile = " & prefisso &"_visibile, " & _
					prefisso &"_ordine_assoluto='" & ParseSql(tip_ordine, adChar) & "', " & _
					prefisso &"_tipologie_padre_lista='"& ParseSql(tip_id, adChar) &"'"
	else
		sql = sql + prefisso &"_livello=" & (cInteger(rs(prefisso &"_livello")) + 1) & ", " + _
					prefisso &"_tipologia_padre_base=" & rs(prefisso &"_tipologia_padre_base") & ", " + _
					prefisso &"_tipologie_padre_lista='"& rs(prefisso &"_tipologie_padre_lista") &","& tip_id &"', "& _
					prefisso &"_albero_visibile=" & IIF(rs(prefisso &"_visibile") AND rs(prefisso &"_albero_visibile"), "1", "0") & ", " + _
					prefisso &"_ordine_assoluto='"
		if tip_ordine<>"" AND rs(prefisso &"_ordine_assoluto")<>"" then		'se l'ordine vuoto lo annulla per tutto il ramo seguente
			sql = sql & rs(prefisso &"_ordine_assoluto") & tip_ordine
		end if
		sql = sql + "' "
		
		if rs(prefisso &"_foglia") then
			'rimuove eventuale flag foglia sul padre (Se e' padre: non e' foglia!)
			rs(prefisso &"_foglia") = false
			rs.update
		end if
	end if
	sql = sql + " WHERE "& prefisso &"_id="& tip_id
	rs.close
	
	'imposta operazioni tipologia e tipologie figlie
	if DB_Type(conn) = DB_Access then
		conn.committrans
		conn.begintrans
	end if
	'aggiorna dati nodo corrente
	CALL conn.execute(sql, , adExecuteNoRecords)
	
	'imposta i dati dei nodi figli
	'..........................................................................
	sql = "SELECT "& prefisso &"_id FROM "& tabella &" WHERE "& prefisso &"_padre_id=" & tip_id
	rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdtext
	if rs.eof then
		sql = "UPDATE "& tabella &" SET "& prefisso &"_foglia=1 WHERE "& prefisso &"_id=" & cIntero(tip_id)
		CALL conn.execute(sql, , adExecuteNoRecords)
	else
		'Giacomo, 14/05/2013, aggiunto query che setta a 0 la colonna _foglia, invece che lasciarla a NULL
		sql = "UPDATE "& tabella &" SET "& prefisso &"_foglia=0 WHERE "& prefisso &"_id=" & cIntero(tip_id)
		CALL conn.execute(sql, , adExecuteNoRecords)
		'-------------------------------------------------------------------------------------------------
		while not rs.eof
			CALL operazioni_ricorsive_tipologia(rs(prefisso &"_id"))
			rs.movenext
		wend
	end if
	rs.close
	
	set rs = nothing
End Sub


'salva i dati dei form nuovo e modifica nel DB e redirige
Public Sub Salva()
	if request.form("salva") <> "" OR request.form("elenco") <> "" OR request.form("propaga_descrittori_nei_figli") <> "" then
		'se cambio padre guarda se quello vecchio e diventato foglia
		if CIntero(request("ID")) > 0 then
			dim rs, sql
			set rs = server.createobject("adodb.recordset")
			sql = "SELECT "& prefisso &"_padre_id FROM "& tabella &" WHERE "& prefisso &"_id = "& cIntero(request("ID"))
			rs.open sql, conn, AdOpenStatic, adLockOptimistic
			if CIntero(rs(prefisso &"_padre_id")) > 0 AND rs(prefisso &"_padre_id") <> request.form("tfn_"& prefisso &"_padre_id") then
				sql = "SELECT COUNT(*) FROM "& tabella &" WHERE "& prefisso &"_padre_id = "& rs(prefisso &"_padre_id")
				if CIntero(GetValueList(conn, null, sql)) = 1 then
					conn.Execute("UPDATE "& tabella &" SET "& prefisso &"_foglia = 1 WHERE "& prefisso &"_id = "& cIntero(rs(prefisso &"_padre_id")))
				end if
			end if
			rs.close
			set rs = nothing
		end if
		
		Set classSalva = New OBJ_Salva
		'Impostazione parametri
		conn.BeginTrans
		Set classSalva.conn = conn
		classSalva.ConnectionString 		= ""
		classSalva.Requested_Fields_List	= "tft_"& prefisso &"_nome_IT"
		classSalva.Checkbox_Fields_List 	= IIF(CategorieAlternative, "chk_" & prefisso & "_principale;chk_" & prefisso & "_alternativa", "")
		classSalva.Page_Ins_Form			= ""
		classSalva.Page_Mod_Form			= ""
		classSalva.Next_Page				= ""	'impostata nella gestione delle relazioni
		classSalva.Next_Page_ID				= FALSE
		classSalva.Table_Name				= tabella
		classSalva.id_Field					= prefisso &"_id"
		classSalva.Read_New_ID				= TRUE
		classSalva.isReport 				= TRUE
		classSalva.Gestione_Relazioni 		= TRUE
		
		'salvataggio/modifica dati
		classSalva.Salva()
	end if
End Sub

'******************************************************************************************************************************************
'******************************************************************************************************************************************
'GESTIONE CANCELLAZIONE CATEGORIA
'******************************************************************************************************************************************

'deve essere richiamata su ClassDelete.Delete_Relazioni()
Public Sub Delete(ID)
	'se ho cancellato l'ultimo figlio il padre diventa foglia
	dim rs, sql, padre
	set rs = server.CreateObject("ADODB.Recordset")
	
	sql = "SELECT "& prefisso &"_padre_id FROM "& tabella &" WHERE "& prefisso &"_id="& cIntero(ID)
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	if not rs.eof then
		padre = cInteger(rs(prefisso &"_padre_id"))
	end if
	rs.close
	if padre > 0 then
		sql = " UPDATE " & tabella & " SET " & prefisso & "_foglia=1 " + _
			  " WHERE " & prefisso & "_id=" & cIntero(padre) & _
			  "		AND (SELECT COUNT(*) FROM " & tabella & " WHERE " & prefisso & "_padre_id=" & cIntero(padre) & ") = 1"
		CALL conn.execute(sql, ,adExecuteNoRecords)
	end if
	set rs = nothing
End Sub


'******************************************************************************************************************************************
'******************************************************************************************************************************************
'GESTIONE SOTTOCATEGORIE
'******************************************************************************************************************************************

'visualizza la pagina che elenca le sottocategorie
Public Sub ElencoSottoCategorie()
	dim rs, rsc, sql, i, lock, HasRelazioni
	set rs = server.CreateObject("ADODB.recordset")
	set rsc = server.CreateObject("ADODB.recordset")

    'recupera dati record principale
	sql = " SELECT * FROM "& tabella &" t WHERE "& prefisso &"_id="& cIntero(request("ID"))
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText %>
	
	<div id="content">
		<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Sotto<%= nomePlurale %> di "<%= NomeCompleto(rs(prefisso &"_id")) %>"</caption>
			<tr><th colspan="7">ELENCO SOTTO<%= UCase(nomePlurale) %></th></tr>
			<%'recupera dati dei figli
			if isB2B then
				sql = " SELECT *, (SELECT COUNT(*) FROM gtb_tipologie WHERE tip_padre_id=t.tip_id) AS N_FIGLI, " & _
				  	 " (SELECT COUNT(*) FROM gtb_tipologie_raggruppamenti WHERE rag_tipologia_id=t.tip_id) AS N_GRUPPI " & _
				  	 " FROM gtb_tipologie t " + _
				   	 " WHERE tip_padre_id=" & cIntero(request("ID")) & " ORDER BY tip_ordine, tip_nome_it"
			else
				sql = " SELECT *, (SELECT COUNT(*) FROM "& tabella &" WHERE "& prefisso &"_padre_id=t."& prefisso &"_id) AS N_FIGLI, " & _
					  " 0 AS N_GRUPPI " & _
					  " FROM "& tabella &" t " + _
				   	  " WHERE "& prefisso &"_padre_id=" & cIntero(request("ID")) & " ORDER BY "& prefisso &"_ordine, "& prefisso &"_nome_it"
			end if
			rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			
			lock = false
			if abilitaBlocchiEsterni then
				lock = (CString(rs(prefisso &"_external_id")) <> "")
			end if %>
			<tr>
				<td colspan="3" class="label_no_width">
					<% if rsc.eof then %>
						Nessuna sotto<%= nomeSingolare %> presente.
					<% else %>
						Trovate n&ordm; <%= rsc.recordcount %>&nbsp;<%= nomePlurale %>
					<% end if %>
				</td>
				<td colspan="4" class="content_right" style="padding-right:0px;">
				<% 	if NOT categorieBloccate AND NOT (lock AND blocchiTotali) then %>
					<a class="button_L2" target="_blank" href="<%= prefissoPagine %>CategorieNew.asp?FROM=<%= FROM_ELENCO %>&tfn_<%= prefisso %>_padre_id=<%= cIntero(request("ID")) %>">
						NUOVA <%= UCase(nomeSingolare) %>
					</a>
				<% 	end if %>
				</td>
			</tr>
            <% if not rsc.eof then %>
				<tr>
					<th>NOME</th>
					<th>CODICE</th>
					<th class="center">ORDINE</th>
					<th class="center" width="1%" colspan="<%= IIF(isB2B, "4", "3") %>">OPERAZIONI</th>
				</tr>
				<% while not rsc.eof
                    HasRelazioni = ConRelazioni(rsc(prefisso &"_id")) %>
					<tr>
						<td class="content"><%= rsc(prefisso &"_nome_it") %></td>
						<td class="content"><%= rsc(prefisso &"_codice") %></td>
						<td class="content_center"><%= rsc(prefisso &"_ordine") %></td>
						<td class="content_center">
                            <%if NOT categorieBloccate OR rsc("N_FIGLI") > 0 then
                                if lock AND blocchiTotali AND rsc("N_FIGLI") = 0 then %>
							        <a class="button_L2_disabled" href="javascript:void(0);" title="Impossibile creare sotto<%= nomePlurale %>: gestione da un applicativo esterno" <%= ACTIVE_STATUS %>>
                                        SOTTO<%= UCase(nomePlurale) %>
                                    </a>
                                <% elseif (not GestioneCategorieMiste AND HasRelazioni) OR rsc("N_GRUPPI")>0 then %>
                                    <a class="button_L2_disabled" href="javascript:void(0);" title="Impossibile creare sotto<%= nomePlurale %>: sono gi&agrave; presenti record associati" <%= ACTIVE_STATUS %>>
                                        SOTTO<%= UCase(nomePlurale) %>
                                    </a>
                                <% else %>
                                    <a class="button_L2" target="_blank" href="<%= prefissoPagine %>CategorieSottocategorie.asp?ID=<%= rsc(prefisso &"_id") %>" title="Apre elenco delle sotto<%= nomePlurale %>." <%= ACTIVE_STATUS %>>
                                        SOTTO<%= UCase(nomePlurale) %>
                                    </a>
                                <%end if
                            end if %>
						</td>
						<% 	if isB2B then %>
    						<td class="content_center">
    							<%if rsc(prefisso &"_foglia") then %>
    								<a class="button_L2" target="_blank" href="<%= prefissoPagine %>CategorieRaggruppamenti.asp?ID=<%= rsc(prefisso &"_id") %>" title="Apre l'elenco dei raggruppamenti della categoria." <%= ACTIVE_STATUS %>>
    									RAGGRUPPAMENTI
    								</a>
    							<% else %>
    								<a class="button_L2_disabled" href="javascript:void(0);" title="Impossibile creare raggruppamenti: la categoria &egrave; intermedia." <%= ACTIVE_STATUS %>>
    									RAGGRUPPAMENTI
    								</a>
    							<% end if %>
    						</td>
						<% 	end if %>
						<td class="content_center">
							<a class="button_L2" target="_blank" href="<%= prefissoPagine %>CategorieMod.asp?FROM=<%= FROM_ELENCO %>&ID=<%= rsc(prefisso &"_id") %>">
								MODIFICA
							</a>
						</td>
						<td class="content_center">					
    						<%if NOT categorieBloccate then
                                if lock then %>
    							    <a class="button_L2_disabled" href="javascript:void(0);" title="Impossibile cancellare la <%= nomeSingolare %>: gestione da un applicativo esterno" <%= ACTIVE_STATUS %>>
    							    	CANCELLA
    							    </a>
    						    <%elseif HasRelazioni OR rsc("N_FIGLI") > 0 OR rsc("N_GRUPPI")>0 then %>
    							    <a class="button_L2_disabled" href="javascript:void(0);" title="Impossibile cancellare la <%= nomeSingolare %>: sono presenti record associati" <%= ACTIVE_STATUS %>>
    								    CANCELLA
    							    </a>
    						    <% else %>
    							    <a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('<%= prefissoPagine %>CATEGORIE','<%= rsc(prefisso &"_id") %>');">
    							    	CANCELLA
    							    </a>
    						    <% end if
                            end if %>
						</td>
					</tr>
					<% rsc.movenext
				wend
			end if
			rsc.close %>
		</table>
		&nbsp;
		</form>
	</div>
	</body>
	</html>
	<%set rs = nothing
	set rsc = nothing
End Sub


'******************************************************************************************************************************************
'******************************************************************************************************************************************
'GESTIONE SELETTORE CATEGORIE
'******************************************************************************************************************************************

Public Sub Seleziona()
	dim Pager
		set Pager = new PageNavigator

	'imposta le variabili iniziali
	if request.querystring("formname")<>"" AND request.servervariables("REQUEST_METHOD") <> "POST" then
		Pager.Reset()
		'imposta parametri iniziali per apertura elenco
		Session("FormName") = request.querystring("formname")
		Session("InputName") = request.querystring("InputName")
		Session("SoloFoglie") = (request("SoloFoglie")<>"")
		Session("Selected") = request.querystring("selected")
		response.redirect prefissoPagine &"CategorieSeleziona.asp"
	end if

	if Request.ServerVariables("REQUEST_METHOD")="POST" then
		Pager.Reset()
		CALL SearchSession_Reset("Sel"& prefisso &"_")
		if not(request("tutti")<>"") then
			CALL SearchSession_Set("Sel"& prefisso &"_")
		end if
	end if
	
	dim conn, sql, rs
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open Application("DATA_ConnectionString")
	set rs = Server.CreateObject("ADODB.RecordSet")
	
	sql = ""
	if session("Sel"& prefisso &"_codice") <> "" then
        sql = sql & SQL_FullTextSearch(session("Sel"& prefisso &"_codice"), tabella &"."& prefisso &"_codice")
	end if
	
	if session("Sel"& prefisso &"_nome") <> "" then
		if sql <> "" then sql = sql & " AND "
		sql = sql & SQL_FullTextSearch(Session("Sel"& prefisso &"_nome"), FieldLanguageList(tabella &"."& prefisso &"_nome_"))
	end if
	
	if session("Sel"& prefisso &"_livello") <> "" then
		if sql <> "" then sql = sql & " AND "
		sql = sql & "  TIP_L0."& prefisso &"_livello=" & session("Sel"& prefisso &"_livello")
	end if
    
	sql = QueryElencoFiltrato(instr(1, Session("Sel"& prefisso &"_tipo"), "0", vbTextCompare)>0, _
                              instr(1, Session("Sel"& prefisso &"_tipo"), "1", vbTextCompare)>0, _
                              Session("SoloFoglie"), sql)

	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	rs.PageSize = 20 %>
	<script language="JavaScript" type="text/javascript">
		function Selezione(objID, objNome){
			opener.<%= Session("FormName") %>.<%= Session("InputName") %>.value=objID.value;
			opener.<%= Session("FormName") %>.view_<%= Session("InputName") %>.value=objNome.value;
            try{
                 opener.<%= Session("FormName") %>_<%= Session("InputName") %>_UpdateTitle(<%= Session("FormName") %>.view_<%= Session("InputName") %>)
            } catch(except){ }
			window.close();
		}
	</script>
	<div id="content_ridotto">
	<form action="" method="post" id="ricerca" name="ricerca">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
		<caption>
			<table border="0" cellspacing="0" cellpadding="1" align="right">
				<tr>
					<td style="font-size: 1px; padding-right:1px;" nowrap>
						<input type="submit" name="cerca" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "CERCA", "SEARCH", "", "", "", "", "", "")%>" class="button">
						&nbsp;
						<input type="submit" name="tutti" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "VEDI TUTTI", "VIEW ALL", "", "", "", "", "", "")%>" class="button">
					</td>
				</tr>
			</table>
			<%= ChooseValueByAllLanguages(Session("LINGUA"), "Opzioni di ricerca", "Search options", "", "", "", "", "", "")%>
		</caption>
		<tr>
			<th<%= Search_Bg("Sel"& prefisso &"_nome") %>><%= ChooseValueByAllLanguages(Session("LINGUA"), "NOME", "NAME", "", "", "", "", "", "")%></th>
			<th<%= Search_Bg("Sel"& prefisso &"_codice") %>><%= ChooseValueByAllLanguages(Session("LINGUA"), "CODICE", "CODE", "", "", "", "", "", "")%></th>
            <% if categorieAlternative then %>
                <th<%= Search_Bg("Sel"& prefisso &"_tipo") %>><%= ChooseValueByAllLanguages(Session("LINGUA"), "TIPO", "TYPE", "", "", "", "", "", "")%></th>
            <% end if %>
			<th<%= Search_Bg("Sel"& prefisso &"_livello") %>><%= ChooseValueByAllLanguages(Session("LINGUA"), "LIVELLO", "LEVEL", "", "", "", "", "", "")%></th>
		</tr>
		<tr>
			<td class="content"><input type="text" class="text" name="search_nome" value="<%= session("Sel"& prefisso &"_nome") %>" maxlength="50" style="width=98%"></td>
			<td class="content" width="20%"><input type="text" class="text" name="search_codice" value="<%= session("Sel"& prefisso &"_codice") %>" maxlength="50" style="width=98%"></td>
             <% if categorieAlternative then %>
                <td class="content" width="25%">
                    <input type="checkbox" class="checkbox" name="search_tipo" value="0" <%= chk(instr(1, session("Sel"& prefisso &"_tipo"), "0", vbTextCompare)>0) %>>
                    <%= ChooseValueByAllLanguages(Session("LINGUA"), "principali", "main", "", "", "", "", "", "")%>
                    &nbsp;&nbsp;
                    <input type="checkbox" class="checkbox" name="search_tipo" value="1" <%= chk(instr(1, Session("Sel"& prefisso &"_tipo"), "1", vbTextCompare)>0) %>>
    				<%= ChooseValueByAllLanguages(Session("LINGUA"), "alternative", "alternative", "", "", "", "", "", "")%>
                </td>
             <% end if %>
			<td class="content" width="10%">
				<% sql = "SELECT MAX("& prefisso &"_livello) FROM "& tabella
				dim levels, i
				set levels = Server.CreateObject("Scripting.Dictionary")
				CALL levels.Add("0", ChooseValueByAllLanguages(Session("LINGUA"), "Categorie di base", "Root category", "", "", "", "", "", ""))
				for i=1 to cInteger(GetValueList(conn, NULL, sql))
					CALL levels.Add(cString(i), ChooseValueByAllLanguages(Session("LINGUA"), "Livello ", "Level ", "", "", "", "", "", "") & i)
				next
				CALL DropDownDictionary(levels, "search_livello", Session("Sel"& prefisso &"_livello"), false, "", Session("LINGUA"))%>
			</td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
		<caption class="border"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Elenco " & nomePlurale , nomePlurale & " list", "", "", "", "", "", "")%></caption>
		<tr>
			<td>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<td class="label_no_width" colspan="5">
							<% if rs.eof then %>
								<%= ChooseValueByAllLanguages(Session("LINGUA"), "Nessuna " & nomeSingolare & " trovata." , "No " & nomeSingolare & " found.", "", "", "", "", "", "")%>
							<% else %>
								<%= ChooseValueByAllLanguages(Session("LINGUA"), "Trovate n&ordm; " & rs.recordcount & " " & nomePlurale & " in n&ordm; " & rs.PageCount & " pagine." , "N&ordm; " & rs.recordcount & " " & nomePlurale & " found in n&ordm; " & rs.PageCount & " pages.", "", "", "", "", "", "")%>
							<% end if %>
						</td>
					</tr>
					<%if not rs.eof then %>
						<tr>
							<th class="l2_center">SEL.</th>
							<th class="L2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "NOME " & UCase(nomeSingolare), UCase(nomeSingolare) & " NAME ", "", "", "", "", "", "")%></th>
							<th class="L2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "CODICE", "CODE", "", "", "", "", "", "")%></th>
							<th class="L2">VIS.</th>
						</tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
							<input type="hidden" name="NAME_<%= rs(prefisso &"_id") %>" value="<%= TextEncode(rs("NAME")) %>">
							<tr>
								<td width="4%" class="content_center">
									<input type="radio" name="seleziona" class="checkbox" value="<%= rs(prefisso &"_id") %>" <%= Chk(cInteger(Session("SELECTED")) = rs(prefisso &"_id")) %>
											   title="Click per selezionare la categoria" onclick="Selezione(this, NAME_<%= rs(prefisso &"_id") %>)">
									</td>
								<td class="<%= IIF(rs(prefisso &"_visibile"), "content", "content_disabled"" title=""categoria non visibile") %>"><%= CBLE(rs, "NAME", Session("LINGUA")) %></td>
								<td class="content"><%= rs(prefisso &"_codice") %></td>
								<td class="content"><input type="checkbox" class="checkbox" disabled <%= chk(rs(prefisso &"_visibile")) %>></td>
							</tr>
							<% rs.MoveNext
						wend%>
						<tr>
							<td colspan="5" class="footer">
								<table width="100%" cellpadding="0" cellspacing="0">
									<tr>
										<td><% 	CALL Pager.Render_GroupNavigator(10, rs.PageCount, "", "button", "button_disabled")%></td>
										<td align="right">
											<a class="button" href="javascript:window.close();" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "chiudi la finestra", "close the window", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
												<%= ChooseValueByAllLanguages(Session("LINGUA"), "CHIUDI", "CLOSE", "", "", "", "", "", "")%></a>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					<% end if %>
				</table>
			</td>
		</tr>
	</form>
	</table>
	</div>
	</body>
	</html>
	<script language="JavaScript" type="text/javascript">
	<!--
		FitWindowSize(this);
	//-->
	</script>
	<%rs.close
	set rs = nothing
End Sub

End Class
%>
