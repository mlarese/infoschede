<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<%
'controllo accesso
if Session("COM_ADMIN") = "" then
	response.redirect "Documenti.asp"
end if

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_tipo_nome"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_tipologie"
	Classe.id_Field					= "tipo_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, i, descr
	
	if request("ID")<>"" then
		sql = "DELETE FROM rel_tipologie_descrittori WHERE rtd_tipologia_id=" & ID
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	end if
	descr = split(request("descr"), ",")
	for i = lbound(descr) to ubound(descr)
		sql = "INSERT INTO rel_tipologie_descrittori(rtd_tipologia_id, rtd_descrittore_id) VALUES (" & ID & ", " & descr(i) & ")"
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	next
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "Tipologie.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>