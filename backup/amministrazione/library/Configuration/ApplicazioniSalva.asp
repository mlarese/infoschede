<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="../database/Tools4Database.asp" -->
<!--#INCLUDE FILE="../ClassSalva.asp" -->
<!--#INCLUDE FILE="../../nextPassport/ToolsApplicazioni.asp" -->
<%
'verifica dei permessi
CALL VerificaPermessiUtente(true)

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= GetConfigurationConnectionstring()
	Classe.Requested_Fields_List	= "tfn_id_sito; tft_sito_nome; tft_sito_p1"
	if request("sito_amministrazione")<>"" then
		Classe.Requested_Fields_List	= Classe.Requested_Fields_List & "; tft_sito_dir"
	end if
	Classe.Checkbox_Fields_List 	= "sito_amministrazione"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_siti"
	Classe.id_Field					= "id_sito"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim i, sql
	
	if request("ID")<>"" then
		'controlla livelli di permesso e cancella quelli non utilizzati
		for i=1 to 9
			if request("tft_sito_p" & i)="" and request("old_value_sito_p" & i)<>"" then
					'commentato perchè nel database configuration non c'è la tabella rel_admin_sito
					
					'livello di permesso cancellato: toglie le associazioni
					'sql = "DELETE FROM rel_admin_sito WHERE sito_id=" & ID & " AND rel_as_permesso=" & i
					'CALL conn.execute(sql, 0, adAsyncFetchNonBlocking)
				end if
		 next
	end if
	
	if (request("sito_amministrazione")="" AND cInteger(request("tfn_sito_rubrica_area_riservata"))=0) then 
		'in inserimento crea la rubrica collegata all'applicazione dell'area riservata
		rs.open "tb_rubriche", conn, adOpenKeyset, adLockOptimistic, adCmdTable
		rs.addNew
		rs("nome_rubrica") = "Utenti - " & request("tft_sito_nome")
		rs("locked_rubrica") = true
		rs("rubrica_esterna") = true
		rs.update
		'collega rubrica inserita all'appplicazione
		sql = "UPDATE tb_siti SET sito_rubrica_area_riservata=" & rs("id_rubrica") & _
			  " WHERE id_sito=" & ID
		CALL conn.execute(Sql, 0, adExecuteNoRecords)
		
		'collega rubrica inserita al gruppo di lavoro
		sql = "INSERT INTO tb_rel_gruppirubriche(id_dellaRubrica, id_gruppo_assegnato) " &_
			  " VALUES (" & rs("id_rubrica") & ", " & IIF(cInteger(Session("GruppoLavoroAreaRiservata"))>0, Session("GruppoLavoroAreaRiservata"), Application("NextPassport_GruppoLavoroAreaRiservata")) & ")"
		CALL conn.execute(Sql, 0, adExecuteNoRecords)
		rs.close
		
	elseif request("sito_amministrazione")<>"" AND cInteger(request("tfn_sito_rubrica_area_riservata"))>0 then
		'elimina collegamento con rubrica 
		sql = "UPDATE tb_siti SET sito_rubrica_area_riservata = NULL WHERE id_sito=" & ID
		CALL conn.execute(Sql, 0, adExecuteNoRecords)
		
		'cancellazione rubrica collegata all'applicazione perche' non piu' utilizzata come area riservata
		sql = "DELETE FROM tb_rubriche WHERE id_rubrica=" & cInteger(request("tfn_sito_rubrica_area_riservata"))
		CALL conn.execute(Sql, 0, adExecuteNoRecords)
	end if
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "Applicazioni.asp"
	
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>