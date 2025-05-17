<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tfn_bun_quantita"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_bundle"
	Classe.id_Field					= "bun_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql
	
	if cInteger(request("tfn_bun_quantita"))<1 then
		Session("ERRORE") = "La quantit&agrave; del componente nel bundle deve essere maggiore di 0."
		exit sub
	end if
	
	'imposta proprieta' del componente
	sql = "SELECT * FROM gtb_articoli WHERE art_id=" & cIntero(request("COM_ID"))
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	if request("BUN_TYPE") = "B" then
		rs("art_in_bundle") = true
	else
		rs("art_in_confezione") = true
	end if
    CALL SetUpdateParamsRS(rs, "art_", false)
	rs.update
	rs.close
	
    'imposta date aggiornamento bundle/confezione
    CALL UpdateParams(conn, "gtb_articoli", "art_", "art_id", request("BUN_ID"), false)
    
	'ricalcola tutte le giacenze a magazzino
	sql = "SELECT mag_id FROM gtb_magazzini"
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	while not rs.eof
		CALL SetGiacenza_b(conn, request("tfn_bun_bundle_id"), QTA_IMPEGNATA, rs("mag_id"))
		CALL SetGiacenza_b(conn, request("tfn_bun_bundle_id"), QTA_ORDINATA, rs("mag_id"))
		CALL SetGiacenza_b(conn, request("tfn_bun_bundle_id"), QTA_GIACENZA, rs("mag_id"))
		rs.movenext
	wend
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