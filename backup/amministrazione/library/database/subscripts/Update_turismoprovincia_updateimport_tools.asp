<%
'................................................................................................................
'
'  TOOLS PER AGGIORNAMENTO DATI DA PORTALI APT
'
'	Import necessario per importare i testi in lingua tedesca, francese e spagnola saltati nel primo import.
'
'.................................................................................................................


'**********************************************************************************************************************************************************************************************************************************
'FUNZIONI COMUNI
'**********************************************************************************************************************************************************************************************************************************


'.................................................................................................................
'procedura che aggiorna il valore del campo di destinazione al campo sorgente solo se questo e' vuoto.
'.................................................................................................................
sub update_import_field(destField, aptFieldValue)
	if cString(destField.value)="" then
		destField.Value = aptFieldValue
	end if
end sub




'**********************************************************************************************************************************************************************************************************************************
'FUNZIONI PER AGGIORNARE L'IMPORT DELLE ANAGRAFICHE
'**********************************************************************************************************************************************************************************************************************************


'.................................................................................................................
'procedura che aggiorna il valore del descrittore solo se i testi di destinazione sono vuoti
'.................................................................................................................
sub update_import_valore_descrittore(conn, rs, AnaId, ExternalSource, ExternalId, valore_fr, valore_de, valore_es)
	dim sql
	dim DesId, DesTipo, EsitoTest
	sql = " SELECT * FROM itb_anagrafiche_descrittori " + _
		  " WHERE and_external_id LIKE '%(" & ExternalId & ")%' AND and_external_source LIKE '" & ExternalSource & "' "
	rs.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
	DesId = cIntero(rs("and_id"))
	DesTipo = rs("and_tipo")
	rs.close
	
	if DesId > 0 AND (valore_fr & valore_de & valore_es) <> "" then
		'inserisce valore/i
		sql = " SELECT * FROM irel_anagrafiche_DescrTipi " + _
			  " WHERE rad_anagrafica_id=" & AnaId & " AND rad_descrittore_id=" & DesId
		rs.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
		if not rs.eof then
			Select Case DesTipo
				case adLongVarChar
					CALL update_import_field(rs("rad_memo_fr"), TextDecode(valore_fr))
					CALL update_import_field(rs("rad_memo_de"), TextDecode(valore_de))
					CALL update_import_field(rs("rad_memo_es"), TextDecode(valore_es))
				case else
					CALL update_import_field(rs("rad_valore_fr"), TextDecode(valore_fr))
					CALL update_import_field(rs("rad_valore_de"), TextDecode(valore_de))
					CALL update_import_field(rs("rad_valore_es"), TextDecode(valore_es))
				end select
			rs.update
		end if
		
		rs.close
	end if
	
end sub


'.................................................................................................................
'procedura che aggiorna l'import delle spiagge
'.................................................................................................................
sub update_import_SPIAGGE(AptCode)
    dim connApt, Apt_rs, rs, sql, AnaCodice
    
    set rs = Server.CreateObject("ADODB.Recordset")
    set Apt_rs = Server.CreateObject("ADODB.Recordset")
    
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
    
    response.write "<!-- Import spiagge " & AptCode & " -->" + vbCrLf
	sql = "SELECT * FROM Spiagge "
	Apt_rs.open sql, connApt, adOpenStatic, adLockOptimistic
    
	while not Apt_rs.eof
		
		AnaCodice = GetCodice(AptCode, CODE_SPIAGGE, Apt_rs("ID_spiagge"))
		%><!-- <%= AnaCodice %> - <%= Apt_rs.absoluteposition %> su <%= Apt_rs.recordcount %>--><%
		
		'verifica se record collegato esiste gia'
		sql = " SELECT * FROM itb_anagrafiche WHERE ana_codice LIKE '" & AnaCodice & "' "
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		if not rs.eof then
			CALL update_import_field(rs("ana_descr_fr"), TextDecode(Apt_rs("note_fra")))
			CALL update_import_field(rs("ana_descr_de"), TextDecode(Apt_rs("note_ted")))
			CALL update_import_field(rs("ana_descr_es"), TextDecode(Apt_rs("note_spa")))
			rs.update
		end if
		rs.close
		
        Apt_rs.movenext
	wend
	
	Apt_rs.close
    
    connApt.close
	
	set connApt = nothing
	set Apt_rs = nothing
	set rs = nothing
end sub


'.................................................................................................................
'procedura che aggiorna l'import dei luoghi
'.................................................................................................................
sub update_import_LUOGHI(AptCode)
    dim connApt, Apt_rs, rs, rst, sql, AnaCodice, AnaId
    
    set rs = Server.CreateObject("ADODB.Recordset")
    set rst = Server.CreateObject("ADODB.Recordset")
    set Apt_rs = Server.CreateObject("ADODB.Recordset")
    
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
    
    response.write "<!-- Import luoghi " & AptCode & " -->" + vbCrLf

    sql = "SELECT * FROM luoghi INNER JOIN TipoLuoghi ON Luoghi.id_tipo = TipoLuoghi.idl "
	Apt_rs.open sql, connApt, adOpenStatic, adLockOptimistic
	
	while not Apt_rs.eof
		
		AnaCodice = GetCodice(AptCode, CODE_LUOGHI, Apt_rs("id"))
		%><!-- <%= AnaCodice %> - <%= Apt_rs.absoluteposition %> su <%= Apt_rs.recordcount %>--><%
        
		'verifica se record collegato esiste gia'
		sql = " SELECT * FROM itb_anagrafiche WHERE ana_codice LIKE '" & AnaCodice & "' "
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		if not rs.eof then
			CALL update_import_field(rs("ana_descr_fr"), TextDecode(Apt_rs("note_fra")))
			CALL update_import_field(rs("ana_descr_de"), TextDecode(Apt_rs("note_ted")))
			CALL update_import_field(rs("ana_descr_es"), TextDecode(Apt_rs("note_spa")))
			rs.update
			
			CALL update_import_valore_descrittore(conn, rst, rs("ana_id"), "general", "apertura", Apt_rs("orar_fra"), Apt_rs("orar_ted"), Apt_rs("orar_spa"))
			CALL update_import_valore_descrittore(conn, rst, rs("ana_id"), "general", "accessibileDisabiliInfo", Apt_rs("info_disabili_fr"), Apt_rs("info_disabili_de"), Apt_rs("info_disabili_es"))
			CALL update_import_valore_descrittore(conn, rst, rs("ana_id"), "general", "ridottoPer", Apt_rs("rid_fra"), Apt_rs("rid_ted"), Apt_rs("rid_spa"))
			
		end if
		rs.close
            
		Apt_rs.movenext
    wend
    
    Apt_rs.close
    
    connApt.close
	
	set connApt = nothing
	set Apt_rs = nothing
	set rs = nothing
    set rst = nothing
end sub


'.................................................................................................................
'aggiornamento import delle notizie utili
'................................................................................................................
sub update_import_NOTIZIE_UTILI(AptCode)
    dim connApt, rsS, rsA, rst, sql, AnaCodice
    
    set rsS = Server.CreateObject("ADODB.Recordset")
	set rsA = Server.CreateObject("ADODB.Recordset")
	set rsT = Server.CreateObject("ADODB.Recordset")
    
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
    
    response.write "<!-- Import Notizie utili " & AptCode & " -->" + vbCrLf
	sql = "SELECT * FROM not_util INNER JOIN Tipi_notutil ON Not_util.tipo = tipi_notutil.id_tipoutil "
	rsS.open sql, connApt, adOpenStatic, adLockOptimistic
	
	while not rsS.eof
		AnaCodice = GetCodice(AptCode, CODE_LUOGHI, rsS("id_UTIL"))
		%><!-- <%= AnaCodice %> - <%= rsS.absoluteposition %> su <%= rsS.recordcount %>--><%
        'verifica se record collegato esiste gia'
		sql = " SELECT * FROM itb_anagrafiche WHERE ana_codice LIKE '" & AnaCodice & "' "
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		if not rs.eof then
			CALL update_import_field(rs("ana_descr_fr"), TextDecode(rsS("descr_fra")))
			CALL update_import_field(rs("ana_descr_de"), TextDecode(rsS("descr_ted")))
			CALL update_import_field(rs("ana_descr_es"), TextDecode(rsS("descr_spa")))
			rs.update
		end if
		rs.close
		
        rsS.movenext
    wend
    
    rsS.close
    
    connApt.close
	
	set connApt = nothing
	set rsS = nothing
	set rsA = nothing
    set rst = nothing
end sub



'.................................................................................................................
'aggiornamento import dei locali e servizi
'................................................................................................................
sub update_import_LOCALI_SERVIZI(AptCode)
    dim connApt, rsS, rsA, rst, sql, AnaCodice
    
    set rsS = Server.CreateObject("ADODB.Recordset")
	set rsA = Server.CreateObject("ADODB.Recordset")
	set rsT = Server.CreateObject("ADODB.Recordset")
    
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
    
    response.write "<!-- Import Locali e servizi " & AptCode & " -->" + vbCrLf
	sql = "SELECT * FROM LocaliEServizi INNER JOIN Tipi_LS ON LocaliEServizi.tipo = Tipi_LS.id_tipoutil "
	rsS.open sql, connApt, adOpenStatic, adLockOptimistic
	
	while not rsS.eof
		AnaCodice = GetCodice(AptCode, CODE_LUOGHI, rsS("id_Ls"))
		%><!-- <%= AnaCodice %> - <%= rsS.absoluteposition %> su <%= rsS.recordcount %>--><%
        'verifica se record collegato esiste gia'
		sql = " SELECT * FROM itb_anagrafiche WHERE ana_codice LIKE '" & AnaCodice & "' "
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		if not rs.eof then
			CALL update_import_field(rs("ana_descr_fr"), TextDecode(rsS("note_fra")))
			CALL update_import_field(rs("ana_descr_de"), TextDecode(rsS("note_ted")))
			CALL update_import_field(rs("ana_descr_es"), TextDecode(rsS("note_spa")))
			rs.update
			
            'registra dati aggiuntivi nei descrittori
			CALL update_import_valore_descrittore(conn, rst, rs("ana_id"), "general", "apertura", rsS("orar_fra"), rsS("orar_ted"), rsS("orar_spa"))
			CALL update_import_valore_descrittore(conn, rst, rs("ana_id"), "general", "ridottoPer", rsS("rid_fra"), rsS("rid_ted"), rsS("rid_spa"))
        end if
        rs.close
        rsS.movenext
    wend
    
    rsS.close
    
    connApt.close
	set connApt = nothing
	set rsS = nothing
	set rsA = nothing
    set rst = nothing
end sub




'**********************************************************************************************************************************************************************************************************************************
'FUNZIONI PER AGGIORNARE L'IMPORT DEGLI EVENTI
'**********************************************************************************************************************************************************************************************************************************


'.................................................................................................................
'funzione che salva valore del descrittore
'.................................................................................................................
sub update_import_EventiValoriDescrittori(conn, rs, evento, descrittore, valore_fr, valore_de, valore_es)
	dim sql
	dim DesId, DesTipo
	sql = " SELECT * FROM itb_eventi_descrittori " + _
		  " WHERE evd_id = "& descrittore
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	DesId = rs("evd_id")
	DesTipo = rs("evd_tipo")
	rs.close
	
	if DesId > 0 AND (valore_fr & valore_de & valore_es) <> "" then
		'inserisce valore/i
		sql = " SELECT * FROM irel_eventi_DescrCat " + _
			  " WHERE red_evento_id=" & evento & " AND red_descrittore_id=" & DesId
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		if not rs.eof then
			Select Case DesTipo
				case adLongVarChar
					CALL update_import_field(rs("red_memo_fr"), TextDecode(valore_fr))
					CALL update_import_field(rs("red_memo_de"), TextDecode(valore_de))
					CALL update_import_field(rs("red_memo_es"), TextDecode(valore_es))
				case else
					CALL update_import_field(rs("red_valore_fr"), TextDecode(valore_fr))
					CALL update_import_field(rs("red_valore_de"), TextDecode(valore_de))
					CALL update_import_field(rs("red_valore_es"), TextDecode(valore_es))
				end select
			rs.update
		end if
		rs.close
	end if
end sub


'.................................................................................................................
'import degli eventi
'.................................................................................................................
sub update_import_EVENTI(AptCode)
    dim connApt, rsS, rsA, sql, EveCodice
    
    set rsS = Server.CreateObject("ADODB.Recordset")
	set rsA = Server.CreateObject("ADODB.Recordset")
    
	Set connApt = Server.CreateObject("ADODB.connection")
	connApt.open Application("DATA_" + AptCode + "_ConnectionString")
    connApt.CommandTimeOut = 360
    
    response.write "<!-- Import Eventi " & AptCode & " -->" + vbCrLf
    'recupera elenco eventi da APT
	sql = " SELECT * FROM Eventi "
	rsS.open sql, connApt, adOpenStatic, adLockOptimistic
    
    while not rsS.eof 
        EveCodice = GetCodice(AptCode, CODE_EVENTI, rsS("id"))
		%><!-- <%= EveCodice %> - <%= rsS.absoluteposition %> su <%= rsS.recordcount %>--><%

		sql = "SELECT * FROM itb_eventi WHERE eve_codice LIKE '" & EveCodice & "'"
    	rsA.open sql, conn, adOpenKeyset, adLockOptimistic
    	if not rsA.eof then
			
			CALL update_import_field(rsA("eve_titolo_fr"), TextDecode(rsS("titol_fra")))
			CALL update_import_field(rsA("eve_titolo_de"), TextDecode(rsS("titol_ted")))
			CALL update_import_field(rsA("eve_titolo_es"), TextDecode(rsS("titol_spa")))
			CALL update_import_field(rsA("eve_descr_fr"), TextDecode(rsS("descr_fra")))
			CALL update_import_field(rsA("eve_descr_de"), TextDecode(rsS("descr_ted")))
			CALL update_import_field(rsA("eve_descr_es"), TextDecode(rsS("descr_spa")))
			CALL update_import_field(rsA("eve_ridotto_fr"), TextDecode(rsS("rid_fra")))
			CALL update_import_field(rsA("eve_ridotto_de"), TextDecode(rsS("rid_ted")))
			CALL update_import_field(rsA("eve_ridotto_es"), TextDecode(rsS("rid_spa")))
			CALL update_import_field(rsA("eve_info_fr"), TextDecode(rsS("info_fra")))
			CALL update_import_field(rsA("eve_info_de"), TextDecode(rsS("info_ted")))
			CALL update_import_field(rsA("eve_info_es"), TextDecode(rsS("info_spa")))
			rsA.update()
			
			'gestione descrittori
			'id descrittori:	1: 	gestione gruppi
			'					2:	prenotazioni
    		CALL update_import_EventiValoriDescrittori(conn, rs, rsA("eve_id"), 1, rsS("gegr_fra"), rsS("gegr_ted"), rsS("gegr_spa"))
    		CALL update_import_EventiValoriDescrittori(conn, rs, rsA("eve_id"), 2, rsS("pren_fra"), rsS("pren_ted"), rsS("pren_spa"))
            
        end if
        
		rsS.movenext
		rsA.close
	wend
	rsS.close
	connApt.close

	set connApt = nothing
	set rsS = nothing
	set rsA = nothing
end sub


%>
