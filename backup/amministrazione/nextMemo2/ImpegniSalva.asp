<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="Tools_Categorie.asp" -->
<!--#INCLUDE FILE="Tools_Memo2.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_imp_titolo_it;tfn_imp_tipo_id"
	Classe.Checkbox_Fields_List 	= "imp_protetto;chk_imp_invia_avviso"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "mtb_impegni"
	Classe.id_Field					= "imp_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
	Classe.SetUpdateParams("imp_")


dim dal, data_fine, al
dal = request("data_inizio") & " " & request("orario_inizio")
data_fine = cString(request("data_fine"))
if data_fine = "" then data_fine = DATA_SENZA_FINE
al = data_fine & " " & request("orario_fine")

if DateDiff("n",DateTimeIso(dal),DateTimeIso(al)) <= 0 then
	Session("ERRORE") = "ATTENZIONE! Orario di inizio maggiore o uguale all'orario di fine impegno."
end if


'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	'..............................................................................
	'sincronizzazione con i contenuti e l'indice
	CALL Index_UpdateItem(conn, Classe.Table_Name, ID, false)
	'..............................................................................
	
	dim val, sql, ut_id
	
	'inserimento relazioni tra profili e impegni
	sql = "DELETE FROM mrel_impegni_profili WHERE rip_impegno_id = " & ID
	conn.Execute(sql)
	for each val in Split(request.form("profili_associati"), ",")
		if CIntero(val) > 0 then
			sql = " INSERT INTO mrel_impegni_profili(rip_impegno_id, rip_profilo_id)"& _
				  " VALUES (" & ID & ", " & val & ")"
			conn.Execute(sql)
		end if
	next
	val = ""

	'inserimento relazioni utenti e documento
	sql = "DELETE FROM mrel_impegni_utenti WHERE riu_impegno_id = " & ID
	conn.Execute(sql)
	for each val in Split(request.form("utenti_associati"), ";")
		if CIntero(val) > 0 then
			ut_id = GetValueList(conn, null, "SELECT ut_ID FROM tb_utenti WHERE ut_NextCom_id = " & val)
			sql = " INSERT INTO mrel_impegni_utenti(riu_impegno_id, riu_utente_id)"& _
				  " VALUES (" & ID & ", " & ut_id & ")"
			'response.write sql
			conn.Execute(sql)
		end if
	next
	val = ""
	
	
	'salvo date e orari
	sql = " UPDATE mtb_impegni SET " & _
		  "		imp_data_ora_inizio = "&SQL_DateTime(conn, dal)&", imp_data_ora_fine = "&SQL_DateTime(conn, al) & _
		  " WHERE imp_id = " & ID
	conn.Execute(sql)
	
	'spedisco gli avvisi
	if cString(request("avviso_dopo_salvataggio")) <> "" then
		CALL SendAvvisoImpegno(conn,ID,cIntero(Session("ID_PAGINA_AVVISO")))
	end if
	
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	if request("salva_calendario")<>"" then
		Classe.Next_Page = "ImpegniCalendarioView.asp?FIRSTDATE=" & request("first_day_week")
	elseif request("salva_elenco")<>"" then
		Classe.Next_Page = "Impegni.asp"
	else
		Classe.Next_Page = "ImpegniMod.asp?ID=" & ID
	end if
end Sub

'salvataggio/modifica dati
Classe.Salva()
	
%>