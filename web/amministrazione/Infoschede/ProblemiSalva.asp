<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_prb_nome_it;"
	if request("prb_modalita_easy")="" then
		Classe.Requested_Fields_List = Classe.Requested_Fields_List + "tft_prb_avviso_per_conferma_it;" 
	end if
	Classe.Checkbox_Fields_List 	= "prb_modalita_easy;prb_visibile;prb_riscontrato"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "sgtb_problemi"
	Classe.id_Field					= "prb_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
    Classe.SetUpdateParams("prb_")
	

	
	
'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	'..............................................................................
	'sincronizzazione con i contenuti e l'indice
	'CALL Index_UpdateItem(conn, Classe.Table_Name, ID, false)
	'..............................................................................
	
	dim val, sql, ut_id
	
	'inserimento relazioni tra profili e problema
	sql = "DELETE FROM srel_problemi_profili WHERE rpp_problema_id = " & ID
	conn.Execute(sql)
	for each val in Split(request.form("profili_associati"), ",")
		if CIntero(val) > 0 then
			sql = " INSERT INTO srel_problemi_profili(rpp_problema_id, rpp_profilo_id)"& _
				  " VALUES (" & ID & ", " & val & ")"
			conn.Execute(sql)
		end if
	next
	val = ""
		
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	if request("salva_indietro")<>"" then
		Classe.Next_Page = "Problemi.asp"
	else
		Classe.Next_Page = "ProblemiMod.asp?ID=" & ID
	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>