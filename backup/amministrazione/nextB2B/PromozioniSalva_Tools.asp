
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_promo_descrizione_it"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= "Promozioni.asp"
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_promozioni"
	Classe.id_Field					= "promo_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim c
	
	conn.Execute(" DELETE FROM grel_promo_articoli WHERE pa_promo_id = "& ID )
	
	for each c in Split(request.form("art"), ",")
		conn.Execute("INSERT INTO grel_promo_articoli (pa_promo_id, pa_art_id) VALUES ("& ID &", "& c &")")
	next
	
	Classe.isReport = FALSE
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
%>