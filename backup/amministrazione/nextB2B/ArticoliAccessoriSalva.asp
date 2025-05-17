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
	Classe.Requested_Fields_List	= ""
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "grel_art_acc"
	Classe.id_Field					= "aa_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql
	
	'imposta propriet&agrave; dell'articolo accessorio
	sql = "SELECT * FROM gtb_articoli WHERE art_id=" & cIntero(request("tfn_aa_acc_id"))
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	rs("art_se_accessorio") = true
	rs("art_noVenSingola") = (request("VendibileSingolarmente")<>"")
    CALL SetUpdateParamsRS(rs, "art_", false)
	rs.update
	rs.close
	
	'imposta stato dell'articolo con accessori
    sql = "SELECT * FROM gtb_articoli WHERE art_id=" & cIntero(request("tfn_aa_art_id"))
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    rs("art_ha_accessori") = true
    CALL SetUpdateParamsRS(rs, "art_", false)
	rs.update
	rs.close

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