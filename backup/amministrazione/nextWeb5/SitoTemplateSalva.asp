<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_template_accesso, 0))

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_nomepage"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_pages"
	Classe.id_Field					= "id_page"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
	Classe.SetUpdateParams("page_")
    
    CALL Classe.AddForcedValue("lingua", NULL)

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	if isNumeric(request.Form("template_padre")) AND request.Form("template_padre")<>"" then
		'copia il template padre
		CALL Copy_page(conn, request.Form("template_padre"), ID, false)
	end if

	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	if request("salva_torna")<>"" then
		Classe.Next_Page = "SitoTemplate.asp"
	else
		Classe.Next_Page = "SitoTemplateMod.asp?ID=" & ID
	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>