<!--#INCLUDE FILE="../site_tools.asp"-->
<!--#INCLUDE FILE="../../tools.asp"-->
<%
'**************************************************************************************************************************************
'funzione che reinizializza l'applicazione.
'**************************************************************************************************************************************
function Applicazione_INIT(conn, rs)
	dim a, sql
	'assume come impostato Application("AZ_ID")
	
	'apre dati del sito
	sql = "SELECT * FROM tb_webs WHERE "
	'recupera sito se preimpostato
	if Session("AZ_ID")<>"" then
		sql = sql & " id_webs=" & cIntero(Session("AZ_ID"))
	else
		sql = sql & " id_webs=" & cIntero(Application("AZ_ID"))
	end if
	rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
	
	'imposta lingua iniziale, se necessario
	if cString(Session("LINGUA"))="" then
		Session("LINGUA") = cString(rs("lingua_iniziale"))
		if cString(Session("LINGUA"))="" then
			Session("LINGUA") = "it"
		end if
	end if
	
	'imposta home page
	Applicazione_INIT = rs("id_home_page")
	
	'imposta indice sito
	Session("AZ_ID") = rs("id_webs")
	
	rs.Close
	
	'ricostruisce array delle pagine
	a = array()
	sql = "SELECT id_pagineSito, id_pagDyn_" & Session("LINGUA") &_
	  	  " FROM tb_pagineSito WHERE id_web=" & cIntero(Session("AZ_ID")) & _
	  	  " ORDER BY id_pagineSito DESC"

	rs.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adAsyncFetch
	if not rs.EOF then
		'imposta dimensione array pagine
		redim a(rs(0)+1)
	
		'salva nell'array l'elenco delle pagine attive del sito
		while not rs.EOF
			a(rs(0)) = rs(1)
			rs.MoveNext
		wend
	end if 
	rs.close
	
	'imposta vettore di sessione pagine
	Session("PAGINE") = a
	Session("VERSION") = 4
	
	CALL LogVisit_Application(conn)
end function
%>