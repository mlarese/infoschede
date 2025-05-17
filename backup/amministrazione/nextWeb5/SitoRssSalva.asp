<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4admin.asp" -->
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_rss_titolo; tft_rss_file; tft_rss_descrizione; tft_rss_query; tfn_rss_freq_generazione"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_rss"
	Classe.id_Field					= "rss_id"
	Classe.Read_New_ID				= FALSE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni
Sub Gestione_Relazioni_record(conn, rs, ID)
	'aggiorna date di modifica del sito per attivare la rigenerazione della NEXT-Application .NET
	CALL UpdateSitoDataModifica(conn, request("tfn_rss_web_id"))
	
	
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