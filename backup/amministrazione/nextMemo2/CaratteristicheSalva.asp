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
	Classe.Requested_Fields_List	= "tft_ct_nome_IT"
	Classe.Checkbox_Fields_List 	= "chk_ct_per_ricerca;chk_ct_per_confronto;chk_ct_principale"
	Classe.Page_Ins_Form			= "CaratteristicheNew.asp"
	Classe.Page_Mod_Form			= "CaratteristicheMod.asp?ID="& request("ID")
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "mtb_carattech"
	Classe.id_Field					= "ct_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, CatList, Cat
	
	'cancella relazioni precedenti
	if request("ID")<>"" then
		sql = "DELETE FROM mrel_categ_ctech WHERE rcc_ctech_id=" & ID
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	end if
	
	'gestione categorie associate
	CatList = split(replace(request("categorie_associate"), " ", ""), ",")
	
	for each Cat in CatList
		'recupera dati ordine
		sql = "INSERT INTO mrel_categ_ctech(rcc_ctech_id, rcc_categoria_id, rcc_ordine) " + _
			  " VALUES (				 " & ID & ",   " & Cat & ", 	 " & cIntero(request("rel_ordine_" & Cat)) & ")"
		CALL conn.execute(sql, , adExecuteNoRecords)
	next

	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "Caratteristiche.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>
