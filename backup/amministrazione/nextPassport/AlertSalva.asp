<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_Passport.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<%
'controllo accesso
CALL CheckAutentication(request("ID")<>"" OR session("PASS_ADMIN") <> "")

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_sev_nome_it;tft_sev_codice"
	Classe.Checkbox_Fields_List 	= "chk_sev_abilitato;chk_sev_multisito"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_siti_eventi"
	Classe.id_Field					= "sev_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql
	
	sql = " SELECT COUNT(*) FROM tb_siti_eventi"& _
		  " WHERE sev_id <> "& ID & _
		  " AND sev_codice LIKE '"& request.form("tft_sev_codice") &"'"
	'controllo univocita codice
	if CIntero(GetValueList(conn, rs, sql)) > 0 then
		session("ERRORE") = "Codice non univoco"
	end if
	
	if session("ERRORE") = "" then
		'imposta parametri per passare alla pagina successiva
		Classe.isReport = FALSE
		
	    if request("salva_elenco")<>"" then
			Classe.Next_Page = "Alert.asp"
		else
	        Classe.Next_Page = "AlertMod.asp?ID="& ID
		end if
	end if
end Sub
	
	'salvataggio/modifica dati
	Classe.Salva()
%>