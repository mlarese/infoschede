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
	Classe.Requested_Fields_List	= "tft_des_nome_it"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= "DescrittoriNew.asp"
	Classe.Page_Mod_Form			= "DescrittoriMod.asp?ID="& request.form("ID")
	Classe.Next_Page				= "Descrittori.asp"
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "sgtb_descrittori"
	Classe.id_Field					= "des_id"
	Classe.Read_New_ID				= FALSE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	Classe.isReport = FALSE
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>