<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
'controllo permessi
CALL CheckAutentication(request("ID") <> "" OR session("PASS_ADMIN") <> "")

dim Classe
Set Classe = New OBJ_Salva

'parametri obbligatori
if request.form("chk_rse_email_abilitato") <> "" then
	Classe.Requested_Fields_List	= "tfn_rse_email_paginaId"
end if
if request.form("chk_rse_fax_abilitato") <> "" then
	Classe.Requested_Fields_List	= Classe.Requested_Fields_List +";tfn_rse_fax_paginaId"
end if
Classe.Requested_Fields_List = TrimChar(Classe.Requested_Fields_List, ";")

'Impostazione parametri
Classe.ConnectionString 		= Application("DATA_ConnectionString")
Classe.Checkbox_Fields_List 	= "chk_rse_email_abilitato;chk_rse_email_admin_invio;chk_rse_email_utenti_invio;"+ _
								  "chk_rse_fax_abilitato;chk_rse_fax_admin_invio;chk_rse_fax_utenti_invio;"+ _
								  "chk_rse_sms_abilitato;chk_rse_sms_admin_invio;chk_rse_sms_utenti_invio"
Classe.Page_Ins_Form			= ""
Classe.Page_Mod_Form			= ""
Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
Classe.Next_Page_ID				= FALSE
Classe.Table_Name				= "rel_siti_eventi"
Classe.id_Field					= "rse_id"
Classe.Read_New_ID				= TRUE
Classe.isReport 				= TRUE
Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sezione, sezioni, sezioniId
	dim contatti, admin
	dim sql, i, val
	sezioni = Array("email", "fax", "sms")		'etichette degli input del form
	sezioniId = Array(MSG_EMAIL, MSG_FAX, MSG_SMS)
	
	'gestione relazioni destinatari aggiuntivi
	for i = 0 to UBound(sezioni)
		'contatti
		sql = "DELETE FROM rel_siti_eventi_contatti WHERE rec_tipo_messaggio_id = "& sezioniId(i) &" AND rec_sitoevento_id = "& ID
		conn.Execute(sql)
		for each val in Split(request.form(sezioni(i) &"_contatti"), ";")
			if CIntero(val) > 0 then
				sql = " INSERT INTO rel_siti_eventi_contatti(rec_sitoevento_id, rec_tipo_messaggio_id, rec_contatto_id)"& _
					  " VALUES ("& ID &", "& sezioniId(i) &", "& val &")"
				conn.Execute(sql)
			end if
		next
		
		'admin
		sql = "DELETE FROM rel_siti_eventi_admin WHERE rea_tipo_messaggio_id = "& sezioniId(i) &" AND rea_sitoevento_id = "& ID
		conn.Execute(sql)
		for each val in Split(request.form(sezioni(i) &"_admin"), ";")
			if CIntero(val) > 0 then
				sql = " INSERT INTO rel_siti_eventi_admin(rea_sitoevento_id, rea_tipo_messaggio_id, rea_admin_id)"& _
					  " VALUES ("& ID &", "& sezioniId(i) &", "& val &")"
				conn.Execute(sql)
			end if
		next
	next
	
	'imposta parametri per salvare e chiudere la finestra corrente
	Classe.isReport = TRUE%>
	<script language="JavaScript" type="text/javascript">
		opener.location.reload(true);
		window.close();
	</script>
<%end Sub

	'salvataggio/modifica dati
	Classe.Salva()
%>