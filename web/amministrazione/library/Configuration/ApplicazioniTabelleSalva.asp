<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="../ClassSalva.asp" -->
<!--#INCLUDE FILE="../IndexContent/ClassContent.asp" -->
<!--#INCLUDE FILE="../database/Tools4Database.asp" -->
<!--#INCLUDE FILE="../../nextPassport/ToolsApplicazioni.asp" -->
<%
'verifica dei permessi
CALL VerificaPermessiUtente(true)

dim content
set content = new ObjContent
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= GetConfigurationConnectionstring()
	Classe.Requested_Fields_List	= "tfn_tab_sito_id;tft_tab_titolo;tft_tab_name;tft_tab_field_chiave;tft_tab_field_titolo_it;tft_tab_from_sql"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_siti_tabelle"
	Classe.id_Field					= "tab_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, rst, var
	set rst = server.CreateObject("ADODB.Recordset")
	
	'esegue una query di test per verificare se tutti i dati immessi sono corretti
	sql = " SELECT TOP 1 " + _
		  	ParseSQL(request("tft_tab_field_chiave"), adChar)
	
	for each var in request.form
		if instr(1, var, "tft_tab_field_", vbTextCompare)>0 AND _
		   instr(1, var, "_chiave", vbTextCompare)<1 then
			sql = sql + AddField(var)
		end if
	next
	
	sql = sql + " FROM " & ParseSQL(request("tft_tab_from_sql"), adChar) %>
	Apertura della query di test per verificare se i dati immessi sono corretti.<br>
	<strong>Query</strong>:<br>
	<%= sql %><br>
	<strong>Risultato:</strong><br>
	Se si vede questo messaggio la prova non &egrave; andata a buon fine: controllare i campi immessi.<br>
	<a href="javascript:history.go(-1)">INDIETRO</a><br>
	<%
	dim connContent
	set connContent = Server.CreateObject("ADODB.Connection")
	connContent.open Application("DATA_ConnectionString")
	rst.open sql, connContent, adOpenStatic, adLockOptimistic, adCmdText
	rst.close
	connContent.close
	set connContent = nothing
	
	'gestione immagini
	rs.open "SELECT * FROM tb_siti_tabelle WHERE tab_id = "& ID, conn, adOpenStatic, adLockOptimistic
	if CIntero(request("tfn_tab_thumb")) = 0 then
		rs("tab_thumb") = 0
		rs.update
	end if
	if CIntero(request("tfn_tab_zoom")) = 0 then
		rs("tab_zoom") = 0
		rs.update
	end if
	rs.close
	
	set rst = nothing
	'imposta parametri per salvare e chiudere la finestra corrente
	Classe.isReport = TRUE%>
	<script language="JavaScript" type="text/javascript">
		opener.location.reload(true);
		window.close();
	</script>
<%end Sub

	'salvataggio/modifica dati
	Classe.Salva()

	
function AddField(FormField)
	dim i, FieldName
	FieldName = ParseSQL(request(FormField), adChar)
	AddField = ""
	if FieldName <> "" then
		AddField = ", (" + FieldName + ") AS " + FormField
	end if
end function
%>