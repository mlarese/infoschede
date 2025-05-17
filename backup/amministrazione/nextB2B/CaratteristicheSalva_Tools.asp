<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_ct_nome_IT"
	Classe.Checkbox_Fields_List 	= "chk_ct_per_ricerca;chk_ct_per_confronto"
	Classe.Page_Ins_Form			= "CaratteristicheNew.asp"
	Classe.Page_Mod_Form			= "CaratteristicheMod.asp?ID="& request("ID")
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_carattech"
	Classe.id_Field					= "ct_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, CatList, Cat
	
	'cancella relazioni precedenti
	if request("ID")<>"" then
		sql = "DELETE FROM gtb_tip_ctech WHERE rct_ctech_id=" & ID
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	end if
	
	'gestione categorie associate
	CatList = split(replace(request("categorie_associate"), " ", ""), ",")

	for each Cat in CatList
		'recupera dati ordine
		dim ordine
		ordine = cIntero(request("rel_ordine_" & Cat))
		if ordine = 0 then
			ordine = 999
		end if
		sql = "INSERT INTO gtb_tip_ctech(rct_ctech_id, rct_tipologia_id, rct_ordine) " + _
			  " VALUES (				 " & ID & ",   " & Cat & ", 	 " & ordine & ")"
		CALL conn.execute(sql, , adExecuteNoRecords)
	next

	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "Caratteristiche.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>