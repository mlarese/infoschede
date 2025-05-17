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
	Classe.Requested_Fields_List	= "tft_val_nome_it;tfn_val_ordine"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_valori"
	Classe.id_Field					= "val_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	'imposta parametri per salvare e chiudere la finestra corrente
	Classe.isReport = TRUE%>
	<script language="JavaScript" type="text/javascript">
		opener.location.reload(true);
		window.close();
	</script>
<%end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>