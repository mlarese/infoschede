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
	Classe.Requested_Fields_List	= "tft_icr_titolo_it"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= "ContattiDescrittoriGruppi.asp"
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_indirizzario_carattech_raggruppamenti"
	Classe.id_Field					= "icr_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim c
	
	conn.Execute(" UPDATE tb_indirizzario_carattech SET ict_raggruppamento_id = NULL WHERE ict_raggruppamento_id = "& ID )
	
	for each c in Split(request.form("car"), ",")
		conn.Execute("UPDATE tb_indirizzario_carattech SET ict_raggruppamento_id = "& ID &" WHERE ict_id = "& c)
	next
	
	Classe.isReport = FALSE
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
%>