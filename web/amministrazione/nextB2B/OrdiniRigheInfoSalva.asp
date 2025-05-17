<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_dod_nome_IT;tft_dod_codice"
	Classe.Checkbox_Fields_List 	= "chk_dod_qta_in_detrazione"
	Classe.Page_Ins_Form			= "OrdiniRigheInfoNew.asp"
	Classe.Page_Mod_Form			= "OrdiniRigheInfoMod.asp?ID="& request("ID")
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_dettagli_ord_des"
	Classe.id_Field					= "dod_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
    
'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, CatList, Cat
	
	'cancella relazioni precedenti
	if request("ID")<>"" then
		sql = "DELETE FROM grel_dettagli_ord_tipo_des WHERE rtd_descrittore_id=" & ID
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	end if
	
	'gestione categorie associate
	CatList = split(replace(request("categorie_associate"), " ", ""), ",")
	
	for each Cat in CatList
		'recupera dati ordine
		sql = "INSERT INTO grel_dettagli_ord_tipo_des(rtd_descrittore_id, rtd_tipo_id, rtd_ordine) " + _
			  " VALUES (				 " & ID & ",   " & Cat & ", 	 " & cIntero(request("rel_ordine_" & Cat)) & ")"
		CALL conn.execute(sql, , adExecuteNoRecords)
	next

	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "OrdiniRigheInfo.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>