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
	Classe.Requested_Fields_List	 = "tft_nome_gruppo; utenti"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_gruppi"
	Classe.id_Field					= "id_gruppo"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, i, utenti
	
	if request("ID")<>"" then
		sql = "DELETE FROM tb_rel_dipgruppi WHERE id_gruppo=" & ID
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	end if
	
	utenti = split(request("utenti"), ",")
	for i = lbound(utenti) to ubound(utenti)
		sql = "INSERT INTO tb_rel_dipgruppi (id_gruppo, id_impiegato) VALUES (" & ID & ", " & utenti(i) & ")"
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	next
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "Gruppi.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>