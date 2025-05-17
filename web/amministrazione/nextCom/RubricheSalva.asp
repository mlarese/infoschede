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
	Classe.Requested_Fields_List	 = "tft_nome_rubrica; gruppi"
	Classe.Checkbox_Fields_List 	= "chk_locked_rubrica;chk_rubrica_esterna"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_rubriche"
	Classe.id_Field					= "id_rubrica"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, i, gruppi, contatti
	
	if request("ID")<>"" then
		sql = "DELETE FROM tb_rel_gruppiRubriche WHERE id_dellaRubrica=" & ID
		CALL conn.execute(sql, 0, adExecuteNoRecords)
		sql = "DELETE FROM rel_rub_ind WHERE id_rubrica=" & ID
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	end if
	
	gruppi = split(request("gruppi"), ",")
	for i = lbound(gruppi) to ubound(gruppi)
		sql = "INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_Gruppo_assegnato) VALUES (" & ID & ", " & gruppi(i) & ")"
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	next
	
	contatti = split(request("contatti"), ";")
	for i = lbound(contatti) to ubound(contatti)-1
		sql = "INSERT INTO rel_rub_ind (id_rubrica, id_indirizzo) VALUES (" & ID & ", " & contatti(i) & ")"
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	next
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "Rubriche.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>