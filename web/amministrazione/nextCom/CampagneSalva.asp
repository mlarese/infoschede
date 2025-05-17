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
	Classe.Requested_Fields_List	 = "tft_inc_nome"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_indirizzario_campagne"
	Classe.id_Field					= "inc_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, i, gruppi, contatti
	
	if request("ID")<>"" then
		sql = "DELETE FROM rel_cnt_campagne WHERE rcc_data_conclusione IS NULL AND rcc_campagna_id=" & ID
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	end if
	
	contatti = request("contatti")
	contatti = split(contatti, ";")
	for i = lbound(contatti) to ubound(contatti)-1
		if cIntero(contatti(i))>0 then
			sql = "INSERT INTO rel_cnt_campagne (rcc_campagna_id, rcc_cnt_id) VALUES (" & ID & ", " & contatti(i) & ")"
			CALL conn.execute(sql, 0, adExecuteNoRecords)
		end if
	next
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "Campagne.asp"
end Sub

'salvataggio/modifica dati
Classe.Salva()
	
%>