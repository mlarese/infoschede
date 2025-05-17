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
	Classe.Requested_Fields_List	= "tft_sts_nome_it;"
	Classe.Checkbox_Fields_List 	= "chk_sts_visibile_admin;chk_sts_visibile_officina;chk_sts_visibile_centr_assist;" & _
									  "chk_sts_modifica_admin;chk_sts_modifica_officina;chk_sts_modifica_centr_assist;" & _
									  "chk_sts_elenco_ddt_da_ritirare;chk_sts_elenco_ddt_da_consegnare"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "sgtb_stati_schede"
	Classe.id_Field					= "sts_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql

	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "SchedeStatiLavorazione.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>