<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tfn_dtd_ddt_id;tft_dtd_articolo_nome"
	Classe.Checkbox_Fields_List 	= "chk_dtd_in_garanzia"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "sgtb_dettagli_ddt"
	Classe.id_Field					= "dtd_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql
	
	'imposta stato dell'articolo con accessori
    ' sql = "SELECT * FROM gtb_articoli WHERE art_id=" & cIntero(request("tfn_aa_art_id"))
	' rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    ' rs("art_ha_accessori") = true
    ' CALL SetUpdateParamsRS(rs, "art_", false)
	' rs.update
	' rs.close

	'imposta parametri per salvare e chiudere la finestra corrente
	Classe.isReport = TRUE%>
	<script language="JavaScript" type="text/javascript">
		opener.document.form1.submit();
		window.close();
	</script>
<%end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>