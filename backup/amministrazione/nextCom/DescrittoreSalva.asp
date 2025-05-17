<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="Tools_DocumentiFiles.asp" -->
<%
'controllo accesso
if Session("COM_ADMIN") = "" then
	response.redirect "Documenti.asp"
end if

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_descr_nome"
	Classe.Checkbox_Fields_List 	= "chk_descr_principale"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_descrittori"
	Classe.id_Field					= "descr_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, i, tipi
	
	if request("ID")<>"" then
		sql = "DELETE FROM rel_tipologie_descrittori WHERE rtd_descrittore_id=" & ID
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	end if
	tipi = split(request("tipi"), ",")
	for i = lbound(tipi) to ubound(tipi)
		sql = "INSERT INTO rel_tipologie_descrittori(rtd_tipologia_id, rtd_descrittore_id) VALUES (" & tipi(i) & ", " & ID & ")"
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	next
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "Descrittori.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>