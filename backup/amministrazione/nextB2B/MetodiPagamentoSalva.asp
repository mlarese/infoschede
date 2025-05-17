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
	Classe.Requested_Fields_List	= "tft_mosp_nome_IT"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= TRUE
	Classe.Table_Name				= "gtb_modipagamento"
	Classe.id_Field					= "mosp_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= FALSE
	Classe.Gestione_Relazioni 		= TRUE

	if cIntero(request("tfn_mosp_ammontare_spsp"))=0 then
		CALL Classe.AddForcedValue("mosp_ammontare_spsp", 0)
	end if
	
'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "MetodiPagamento.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>