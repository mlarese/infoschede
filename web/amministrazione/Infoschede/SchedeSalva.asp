<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="../library/Class_Mailer.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
dim conn, sql, rs
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tfn_sc_stato_id;tfn_sc_cliente_id;tfn_sc_modello_id"
	Classe.Checkbox_Fields_List 	= "chk_sc_in_garanzia;chk_sc_richiesta_garanzia"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "sgtb_schede"
	Classe.id_Field					= "sc_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
    Classe.SetUpdateParams("sc_")

	
'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	'..............................................................................
	'sincronizzazione con i contenuti e l'indice
	'CALL Index_UpdateItem(conn, Classe.Table_Name, ID, false)
	'..............................................................................
	
	'inserissco il numero della scheda, se è un nuovo inserimento
	if request("nuovo_inserimento") = "true" then
		sql = " UPDATE sgtb_schede SET sc_numero = (SELECT MAX(sc_numero) + 1 FROM sgtb_schede) WHERE sc_id = " & ID
		conn.execute(sql)
	end if
	
	'salva descrittori
	CALL DesSave(conn, ID, "srel_descrittori_schede", "rds_valore_", "rds_memo_", "rds_scheda_id", "rds_descrittore_id", "")
 
	'scrive sul log
	CALL WriteLogAdmin(conn,"sgtb_schede",ID,_
						IIF(request("ID")<>"","modifica","inserimento"),"Infoschede - schede")	
		
		
	'se è un nuovo inserimento, spedisco l'e-mail di avviso
	if request("nuovo_inserimento") = "true" then
		conn.commitTrans
		dim num, cliente_id, cntId_Dest, key, url
		set rs = server.createobject("adodb.recordset")
		sql = " SELECT sc_numero FROM sgtb_schede WHERE sc_id = " & ID
		num = CIntero(GetValueList(conn, NULL, sql))
		sql = " SELECT sc_cliente_id FROM sgtb_schede WHERE sc_id = "&ID
		cliente_id = CIntero(GetValueList(conn, NULL, sql))
		sql = " SELECT ut_NextCom_id FROM tb_utenti WHERE ut_id IN ("&cliente_id&")"
		cntId_Dest = CIntero(GetValueList(conn, NULL, sql))
		sql = " SELECT codiceInserimento FROM tb_indirizzario WHERE IDElencoIndirizzi = " & cntId_Dest
		key = cString(GetValueList(conn, NULL, sql))
		
		url = "&SCHEDAID="&ID&"&CLIENTEID="&cliente_id&"&IDCNT="&cntId_Dest&"&KEY="&key&"&CONFERMA=1&HTML_FOR_EMAIL=1"
		
		response.write GetPageSiteUrl(conn, ID_PAGINA_AVVISO_NUOVA_SCHEDA, "it") & url
		
		
		CALL SendPageFromAdminToContactExtended(conn, rs, "it", _
												"Inserimento richiesta di assistenza n. "&num, _
												GetPageSiteUrl(conn, ID_PAGINA_AVVISO_NUOVA_SCHEDA, "it")&url, _
												GetSiteBaseUrl(conn, ID_PAGINA_AVVISO_NUOVA_SCHEDA), _
												Session("ID_ADMIN"), _
												cntId_Dest, _
												false)
		set rs = nothing
		conn.BeginTrans
	end if						

	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	if request("salva_continua")<>"" then
		dim id_ass, sql
		sql = "SELECT sc_centro_assistenza_id FROM sgtb_schede WHERE sc_id = " & ID
		id_ass = cINtero(GetValueList(conn, NULL, sql))
		if id_ass > 0 then
			Classe.Next_Page = "Schede.asp?ASSEGNATA=true"
		else
			Classe.Next_Page = "Schede.asp?ASSEGNATA=false"
		end if
	else
		Classe.Next_Page = "SchedeMod.asp?ID=" & ID
	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>