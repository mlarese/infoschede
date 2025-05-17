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
	Classe.Requested_Fields_List	= "tfn_sco_qta_da;"
	if cReal(request("tfn_sco_sconto"))>0 then 
		Classe.Requested_Fields_List = Classe.Requested_Fields_List + "tfn_sco_sconto;" 
	else
		Classe.Requested_Fields_List = Classe.Requested_Fields_List + "tfn_sco_prezzo;" 
	end if
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= "ScontiQ_dettagliNew.asp?EXTID="& request("EXT_ID")
	Classe.Page_Mod_Form			= "ScontiQ_dettagliMod.asp?EXTID="& request("EXT_ID") &"&ID="& request("ID")
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_scontiQ"
	Classe.id_Field					= "sco_id"
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