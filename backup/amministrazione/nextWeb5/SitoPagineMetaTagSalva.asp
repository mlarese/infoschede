<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 10000000 %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
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
	Classe.Table_Name				= "tb_pagineSito"
	Classe.id_Field					= "id_pagineSito"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
	
	'update campi gestione modifiche
	classe.SetUpdateParams("ps_")

'definizione eventuali operazioni su relazioni
Sub Gestione_Relazioni_record(conn, rs, ID)
	
	'..............................................................................
	'sincronizzazione con i contenuti e l'indice
	CALL Index_UpdateItem(conn, "tb_pagineSito", ID, true)
	'..............................................................................
	
	classe.isReport = true %>
	<script language="JavaScript" type="text/javascript">
		window.close();
	</script>
<%end Sub

	'salvataggio/modifica dati
	Classe.Salva()
%>