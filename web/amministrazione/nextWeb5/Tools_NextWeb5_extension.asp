
<%
'Riporto su questo file alcune funzioni che erano presenti su Tools_NextWeb5.asp.
'Questo perchè le seguenti funzioni mi servono su Update__library__framework_core.asp, 
'e includendo Tools_NextWeb5.asp mi dava un errore dovuto alla "doppia inclusione" (il nome ... è ridefinito)
%>

<%
'....................................................................................................
'		funzione che controlla se le pagine di stage esistono, altrimenti le crea.
'		crea le pagine per tutte le lingue in modo da precalcolare tutti i link dell'indice.
'		conn:			connessione attiva al database
'		rs:				recordset aperto su tb_paginesito
'....................................................................................................
sub Ceck_page_exists(conn, rs)
	dim lingua, i, val
	for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
		lingua = Application("LINGUE")(i)
		'if Session("LINGUA_" & lingua ) then
			if cInteger(rs("id_pagDyn_" & lingua))=0 OR cInteger(rs("id_pagStage_" & lingua))=0 then
				'verifica se la pagina di stage e' stata creata	
				if cInteger(rs("id_pagStage_" & lingua))=0 then
					'pagina mancante: la crea
					val = Create_page(conn, IIF(rs("nome_ps_" & lingua)<>"", rs("nome_ps_" & lingua), rs("nome_ps_IT")), rs("id_web"), rs("id_pagineSito"), lingua)
					rs("id_pagStage_" & lingua) = cIntero(val)
				end if
				
				'verifica se la pagina pubblica e' stata creata
				if cInteger(rs("id_pagDyn_" & lingua))=0 then
					'pagina mancante: la crea
					rs("id_pagDyn_" & lingua) = Create_page(conn, IIF(rs("nome_ps_" & lingua)<>"", rs("nome_ps_" & lingua), rs("nome_ps_IT")), rs("id_web"), rs("id_pagineSito"), lingua)
				end if
				rs.Update
			end if
		'end if
	next
end sub




'....................................................................................................
'		funzione che aggiorna il nome delle pagine nella tabella tb_pages copiandolo dalla tabella 
'		tb_pagineSito.
'		conn:				connessione attiva al database
'       IDpagSito:     		ID della pagina sito
'....................................................................................................
sub PaginaSitoUpdatePages(conn, IDpagSito)
	dim sql, i, lingua, rs
	set rs = server.createobject("adodb.recordset")
	if cInteger(IDpagSito)>0 then
		for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
			lingua = Application("LINGUE")(i)
			if Session("LINGUA_" & lingua) then 
				sql = "SELECT nome_ps_" & lingua & " FROM tb_pagineSito WHERE id_pagineSito=" & IDpagSito
				sql = " UPDATE tb_pages SET nomepage=" + SQL_UTF8Qualifier(conn) + "'" & ParseSQL(GetValueList(conn, rs, sql), adChar) & "' WHERE id_paginaSito=" & IDpagSito & " AND lingua like '" & lingua & "'"
				CALL conn.execute(sql, 0, adExecuteNoRecords)		
			end if
		next
	end if
	set rs = nothing
end sub




'....................................................................................................
'		funzione che crea una pagina e ne restituisce l'ID
'		conn:			connessione attiva al database
'		nomepage		nome della pagian da creare
'		id_webs			id del sito a cui e' associata la pagina
'....................................................................................................
function Create_page(conn, nomepage, id_webs, id_paginaSito, lingua)
	Create_page = Create_page_New(conn, nomepage, id_webs, id_paginaSito, lingua, 0)
end function


'....................................................................................................
'		funzione che crea una pagina e ne restituisce l'ID
'		conn:			connessione attiva al database
'		nomepage		nome della pagian da creare
'		id_webs			id del sito a cui e' associata la pagina
'		idTemplate		id del template associato alla pagina
'....................................................................................................
function Create_page_New(conn, nomepage, id_webs, id_paginaSito, lingua, idTemplate)
	dim rs, sql
	set rs = server.createObject("ADODB.Recordset")
	
	'inserisce nuova pagina
	sql = "SELECT * FROM tb_pages WHERE id_webs=" & id_webs
	rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
	rs.AddNew
	rs("id_webs") = id_webs
	rs("id_PaginaSito") = id_paginaSito
'response.write "<br>id pagina:" & id_paginaSito & "-nomepage:" & nomepage & "-" & vbCrlf
 	rs("nomepage") = nomepage
	rs("lingua") = lingua
	rs("contatore") = 0
	rs("contUtenti") = 0
	rs("contCrawler") = 0
	rs("contAltro") = 0
	rs("ContRes") = Date
	rs("template") = false
	if cIntero(idTemplate)>0 then
		rs("id_template") = idTemplate
	end if
	CALL SetUpdateParamsRS(rs, "page_", true)
	rs.Update
	
	'imposta id pagina appena inserita
	Create_page_New = cIntero(rs("id_page"))

    'aggiorna data di ultima modifica della struttura delle pagine del sito
    CALL UpdateSitoDataModificaPagine(conn, rs("id_webs"))
    
	rs.close
	set rs = nothing
    
end function


'....................................................................................................
'		funzione che aggiorna data di ultima modifica della struttura delle pagine del sito
'		conn:			connessione attiva al database
'		webs_id			id del sito le cui pagine sono state variate
'....................................................................................................
sub UpdateSitoDataModificaPagine(conn, webs_id)
    dim sql
    sql = "UPDATE tb_webs SET webs_modData_pagine = " & SQL_Now(conn) & " WHERE id_webs=" & webs_id
    CALL conn.execute(sql, ,adExecuteNoRecords)
end sub 


%>