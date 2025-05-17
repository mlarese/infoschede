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
	Classe.Requested_Fields_List	= "tft_rag_titolo_it"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= "DescrRag.asp"
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "sgtb_descrittori_raggruppamenti"
	Classe.id_Field					= "rag_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim c
	
	conn.Execute(" UPDATE sgtb_descrittori SET des_raggruppamento_id = 0"& _
				 " WHERE des_raggruppamento_id = "& ID)
	for each c in Split(request.form("car"), ",")
		conn.Execute("UPDATE sgtb_descrittori SET des_raggruppamento_id = "& ID &" WHERE des_id = "& c)
	next
	
	Classe.isReport = FALSE
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
%>