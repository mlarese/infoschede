<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<%

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_att_oggetto;tft_att_testo"
	Classe.Checkbox_Fields_List 	= "chk_att_priorita;chk_att_conclusa"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_attivita"
	Classe.id_Field					= "att_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
	Classe.ID_Value = request("ID")

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim all, i, sql, utentiOrbi
	
	'flag bozza
	sql = "SELECT * FROM tb_attivita WHERE att_id="& ID
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	rs("att_inSospeso") = (request.form("salva") = "SALVA COME BOZZA" OR request.form("mod") = "SALVA COME BOZZA")
	rs("att_mittente_id") = Session("ID_ADMIN")
	
	'scrivo la data di chiusura se chiudo
	if request.form("chk_att_conclusa") <> "" then
		rs("att_dataChiusa") = Now()
		rs("att_utente_chiusura") = Session("ID_ADMIN")
		rs("att_conclusa") = true
	end if
	rs.update
	rs.close
	
	if request("ID") = "" then				
		'gesione inserimento access list
		CALL AL_ins(conn, AL_ATTIVITA, ID, (request.form("ere") <> ""))
		if cInteger(request("tfn_att_pratica_id"))>0 then
			sql = "UPDATE tb_pratiche SET pra_dataUM="& SQL_Now(conn) &" WHERE pra_id="& request("tfn_att_pratica_id")
			CALL conn.execute(sql, , adExecuteNoRecords)
		end if
		
		'gestione chiusura domanda
		if request.form("tfn_att_domanda_id") <> "" then		'se e' una risposta
			'chiudo la domanda
			sql = "UPDATE tb_attivita SET att_conclusa=1, att_utente_chiusura=" & Session("ID_ADMIN") & ", att_dataChiusa=" & SQL_Now(conn) & _
				  " WHERE att_id="& request.form("tfn_att_domanda_id")
			CALL conn.execute(sql, ,adExecuteNoRecords)
		end if
	end if
	
	if request("tfd_att_dataCrea") = "NOW" then			'se sono in inserimento o in modifica come creatore o sono amministratore
		'GESTIONE ALLEGATI
		sql = "DELETE FROM tb_allegati WHERE all_attivita_id="& ID
		CALL conn.execute(sql, , adExecuteNoRecords)
		if request.form("documenti") <> "" then
			all = split(left(request.form("documenti"), len(request.form("documenti"))-1), ";")
			for i = lbound(all) to ubound(all)
				sql = "INSERT INTO tb_allegati (all_attivita_id, all_documento_id) VALUES (" & ID & ", " & all(i) & ")"
				CALL conn.execute(sql, , adExecuteNoRecords)
			next
		end if
	end if
	
	'controllo la visibilita dei documenti da parte dei destinatati
	if request.form("documenti") <> "" then
		sql = "SELECT id_admin FROM tb_admin INNER JOIN tb_rel_dipGruppi ON " & _
				  "tb_admin.id_admin=tb_rel_dipGruppi.id_impiegato " & _
			  "WHERE (id_admin IN (SELECT al_utente_id FROM al_attivita_utenti " & _
					   		   	  "WHERE al_tipo_id="& ID &") " & _
			  "OR id_admin IN (SELECT id_impiegato FROM al_attivita_gruppi t INNER JOIN " & _
					   		  "tb_rel_dipGruppi r ON t.al_gruppo_id=r.id_gruppo " & _
							  "WHERE al_tipo_id="& ID &")) " & _
			  "AND (id_admin NOT IN (SELECT al_utente_id FROM al_documenti_utenti c INNER JOIN " & _
				  				 	"tb_allegati d ON c.al_tipo_id=d.all_documento_id " & _
				  		   		 	"WHERE all_attivita_id="& ID &") " & _
		 	  "AND id_admin NOT IN (SELECT id_impiegato FROM (al_documenti_gruppi a INNER JOIN " & _
				 				   "tb_rel_dipGruppi b ON a.al_gruppo_id=b.id_gruppo) INNER JOIN " & _
								   "tb_allegati e ON a.al_tipo_id=e.all_documento_id " & _
								   "WHERE all_attivita_id="& ID &")) " & _
			  "OR (id_admin NOT IN (SELECT al_utente_id FROM al_documenti_utenti e INNER JOIN " & _
				  				   "tb_allegati f ON e.al_tipo_id=f.all_documento_id " & _
				  		   		   "WHERE all_attivita_id="& ID &") " & _
		 	  "AND id_admin NOT IN (SELECT DISTINCT id_impiegato FROM ((al_documenti_gruppi g INNER JOIN " & _
				 				   "tb_rel_dipGruppi h ON g.al_gruppo_id=h.id_gruppo) INNER JOIN " & _
								   "tb_allegati i ON g.al_tipo_id=i.all_documento_id) " & _
								   "WHERE all_attivita_id="& ID &") "& _
			  "AND "& SQL_IsTrue(conn, "(SELECT att_pubblica FROM tb_attivita WHERE att_id="& ID &")") &")"
		utentiOrbi = GetValueList(conn, rs, sql)
		if utentiOrbi <> "" then
			sql = "SELECT (COUNT(*)) AS N_DOCS FROM tb_documenti d INNER JOIN tb_allegati a "& _
				  "ON d.doc_id=a.all_documento_id "& _
				  "WHERE doc_id NOT IN (SELECT al_tipo_id FROM al_documenti_utenti "& _
				  				       "WHERE al_utente_id IN ("& utentiOrbi &")) "& _
				  "AND doc_id NOT IN (SELECT al_tipo_id FROM al_documenti_gruppi aldg INNER JOIN "& _
				  					 "tb_rel_dipgruppi r ON aldg.al_gruppo_id=r.id_gruppo "& _
									 "WHERE id_impiegato IN ("& utentiOrbi &")) "& _
				  "AND all_attivita_id="& ID &" AND ("& AL_query(conn, AL_DOCUMENTI) & _
				  " OR doc_creatore_id="& Session("ID_ADMIN") &") AND NOT "& SQL_IsTrue(conn, "doc_pubblica")
			rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
			if rs("N_DOCS")<1 then
				utentiOrbi = ""
			end if
			rs.close
		end if
	else 
		utentiOrbi = ""
	end if
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	if request.form("mod") <> "" OR (request.form("salva") <> "" AND utentiOrbi = "") then
		Classe.Next_Page = "Attivita.asp"
	else
		Classe.Next_Page = "AttivitaMod.asp?ID="& ID
	end if
	
end Sub



'salvataggio/modifica dati
Classe.Salva()
%>