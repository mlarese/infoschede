<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/ClassContent.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tfn_com_idx_id;tfn_com_contatto_id;tft_com_comment;"
	Classe.Checkbox_Fields_List 	= "chk_com_validate"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_comments"
	Classe.id_Field					= "com_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE


'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
end Sub

%>
		<script language="JavaScript" type="text/javascript">
				opener.location.reload(true);
			window.close();
		</script>
<%
	'salvataggio/modifica dati
	Classe.Salva()
	
	
%>