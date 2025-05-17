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
	Classe.Requested_Fields_List	= "tft_mag_nome"
	Classe.Checkbox_Fields_List 	= "chk_mag_vendita_pubblico;chk_mag_disponibilita"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_magazzini"
	Classe.id_Field					= "mag_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql
	
	if request("chk_mag_vendita_pubblico")<>"" then
		sql = "UPDATE gtb_magazzini SET mag_vendita_pubblico=0 WHERE mag_id<>" & ID
		CALL conn.execute(sql, , adExecuteNoRecords)
	end if
	
	'inserisce esploso dei codici
	if request("ID") = "" then
		sql = " INSERT INTO grel_giacenze ( gia_magazzino_id, gia_art_var_id, gia_qta, gia_impegnato, gia_ordinato ) " + _
			  " SELECT " & ID & ", rel_id,0,0,0 FROM grel_art_valori "
		CALL conn.execute(sql, , adExecuteNoRecords)
	end if
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "Magazzini.asp"
	
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
	
	
%>
