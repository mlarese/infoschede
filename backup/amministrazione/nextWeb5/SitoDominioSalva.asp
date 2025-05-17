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
	Classe.Requested_Fields_List	= "tft_dom_url;tft_dom_lingua;tft_dom_name"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_webs_domini"
	Classe.id_Field					= "dom_id"
	Classe.Read_New_ID				= FALSE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni
Sub Gestione_Relazioni_record(conn, rs, ID)
	'imposta parametri per salvare e chiudere la finestra corrente
	Classe.isReport = TRUE
	
	dim sql
	sql = "SELECT dom_url FROM " & Classe.Table_Name & " WHERE " & Classe.id_Field & " = " & ID
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	if not rs.eof then
		if inStr(rs("dom_url"), "http://") <> 1 AND inStr(rs("dom_url"), "https://") <> 1 then
			rs("dom_url") = "http://" & rs("dom_url")
			rs.update
		end if
	end if
	rs.close
	
	%>
	<script language="JavaScript" type="text/javascript">
		opener.location.reload(true);
		window.close();
	</script>
<%end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>