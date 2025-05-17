<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<%

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_tq_nome;tft_tq_separatore;tft_tq_query"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_siti_tabelle_tag_query"
	Classe.id_Field					= "tq_id"
	Classe.Read_New_ID				= FALSE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

	
Sub Gestione_Relazioni_record(conn, rs, ID)
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = TRUE
%>
	<script language="JavaScript" type="text/javascript">
		opener.location.reload(true);
		window.close();
	</script>
<%	
end Sub

	
	'salvataggio/modifica dati
	Classe.Salva()
	
%>