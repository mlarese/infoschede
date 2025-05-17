
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_so_nome_it;tfn_so_ordine;tfn_so_stato_ordini"
	Classe.Checkbox_Fields_List 	= "chk_so_internet"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_stati_ordine"
	Classe.id_Field					= "so_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql
	
	if request("chk_so_internet")<>"" then
		sql = "UPDATE gtb_stati_ordine SET so_internet=0 WHERE so_id<>" & ID
		CALL conn.execute(sql,, adExecuteNoRecords)
	end if

	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "OrdiniStatiLavorazione.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>