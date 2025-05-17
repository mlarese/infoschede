<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="Tools_Categorie.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_pro_nome_it;"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "mtb_profili"
	Classe.id_Field					= "pro_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE


'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim val, sql, ut_id
	
	'inserimento relazioni tra admin e profilo
	sql = "DELETE FROM mrel_profili_admin WHERE rpa_profilo_id = " & ID
	conn.Execute(sql)
	for each val in Split(request.form("admin_associati"), ";")
		if CIntero(val) > 0 then
			sql = " INSERT INTO mrel_profili_admin(rpa_profilo_id, rpa_admin_id)"& _
				  " VALUES (" & ID & ", " & val & ")"
			conn.Execute(sql)
		end if
	next
	val = ""
		
	'inserimento relazioni tra utenti e profilo
	sql = "DELETE FROM mrel_profili_utenti WHERE rpu_profilo_id = " & ID
	conn.Execute(sql)
	for each val in Split(request.form("utenti_associati"), ";")
		if CIntero(val) > 0 then
			ut_id = GetValueList(conn, null, "SELECT ut_ID FROM tb_utenti WHERE ut_NextCom_id = " & val)
			sql = " INSERT INTO mrel_profili_utenti(rpu_profilo_id, rpu_utenti_id)"& _
				  " VALUES (" & ID & ", " & ut_id & ")"
			response.write sql
			conn.Execute(sql)
		end if
	next
	val = ""
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "Profili.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>