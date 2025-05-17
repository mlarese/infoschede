<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/ClassIndirizzarioSyncro.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_pro_nome_IT"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_profili"
	Classe.id_Field					= "pro_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
    Classe.SetUpdateParams("pro_")
    
	
'definizione eventuali operazioni su relazioni
Sub Gestione_Relazioni_record(conn, rs, ID)
    dim sql, IdRubrica
    
    'sincronizza dati rubrica collegata
    IdRubrica = UpdateSyncroRubricaGruppo(conn, rs, "(B2B) Clienti - Profilo " & request("tft_pro_nome_it"), "gtb_profili", "gtb_profili", ID, Application("NextCom_DefaultWorkGroup"))
    
	sql = "UPDATE gtb_profili SET pro_rubrica_id = "&IdRubrica&" WHERE pro_id = " & ID
	conn.Execute(sql)
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "ClientiProfili.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>