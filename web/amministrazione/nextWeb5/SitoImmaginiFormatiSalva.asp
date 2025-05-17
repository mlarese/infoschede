<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="Tools_NEXTweb5.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_immaginiFormati_accesso, 0))

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_imf_nome"
	Classe.Checkbox_Fields_List 	= "chk_imf_suffissoFormato;chk_imf_dimensioniMax;chk_imf_salvaOriginale"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= "SitoImmaginiFormati.asp"
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_immaginiFormati"
	Classe.id_Field					= "imf_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
	Classe.SetUpdateParams("imf_")

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	if request.form("chk_imf_suffissoFormato") = "" AND request.form("tft_imf_dir") = "" AND request.form("tft_imf_suffisso") = "" then
		session("ERRORE") = "Selezionare il suffisso oppure una directory separata."
	elseif request.form("tfn_imf_width") = "" AND request.form("tfn_imf_height") = "" then
		session("ERRORE") = "Selezionare la larghezza o l'altezza."
	else
		'imposta parametri per passare alla pagina successiva
		Classe.isReport = FALSE
	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
%>