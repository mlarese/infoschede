<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_stili_accesso, 0))

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= ""
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_css_styles"
	Classe.id_Field					= "style_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
	
	'update campi gestione modifiche
	classe.SetUpdateParams("style_")

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim cssO, sql
	
	'rigenera file degli stili.
	set cssO = new CssManager
	CALL cssO.GenerateCssFile(conn, request("grp_id_webs"))
	CALL cssO.UpdateChecksum(conn, rs, request("style_grp_id"))
	set cssO = nothing
	sql = " UPDATE tb_css_groups SET"& _
		  " grp_modAdmin_id = "& session("ID_ADMIN") &","& _
		  " grp_modData = "& SQL_Date(conn, GetValueList(conn, rs, "SELECT style_modData FROM tb_css_styles WHERE style_id = "& ID)) & _
		  " WHERE grp_id = "& cIntero(request("style_grp_id"))
	conn.execute(sql)
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "SitoStili.asp"
end Sub

'salvataggio/modifica dati
Classe.Salva()

%>