<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/ClassIndirizzarioLock.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<%
dim conn, sql, contatto, campo, rs, contatti, archiviata
set conn = server.createobject("adodb.connection")
set rs = server.createobject("adodb.recordset")
conn.open Application("DATA_ConnectionString")

'controllo accesso
if request("mod")<>"" OR request("AL")<>"" then
	if (NOT AL(conn, cIntero(request("ID")), AL_PRATICHE) OR Session("COM_POWER") = "" AND _
	   Session("ID_ADMIN") <> CInt(GetValueList(conn, rs, "SELECT pra_creatore_id FROM tb_pratiche WHERE pra_id="& cIntero(request("ID"))))) _
	   AND Session("COM_ADMIN") = "" then
		response.redirect "Pratiche.asp?ID="& Session("COM_PRA_CLIENTE")
	end if
end if


'controllo campi obbligatori
campo = request.form("tft_pra_nome")
if campo = "" then
	Session("errore") = "Campo nome pratica obbligatorio!"
end if

'attiva transazione
conn.beginTrans

if Session("ERRORE") = "" then
	if request.querystring("ID") = "" then
		'inserimento pratiche
		campo = request.form("contatti")
		if campo = "" then
			Session("errore") = "Scegli almeno un contatto!"
		else
			contatti = left(request.form("contatti"), len(request.form("contatti"))-1)
			sql = "SELECT * FROM tb_pratiche"
			rs.open sql, conn, adOpenKeyset, adLockOptimistic, adCmdText
			for each contatto in Split(contatti, ";")
				rs.AddNew
				rs("pra_nome") = request.form("tft_pra_nome")
				rs("pra_dataI") = NOW
				rs("pra_dataUM") = NOW
				rs("pra_archiviata") = 0
				rs("pra_note") = request.form("tft_pra_note")
				rs("pra_cliente_id") = contatto
				rs("pra_creatore_id") = Session("ID_ADMIN")
				rs("pra_mod_data") = NOW()
				rs("pra_mod_utente") = Session("ID_ADMIN")
				rs.Update
				CALL InserimentoAttivita(conn, cInt(rs("pra_id")), contatto)
			next
			rs.close
		end if
	else
		'modifica pratica
		if request.form("chk_pra_archiviata") = "" then
			archiviata = 0
		else
			'controllo che tutte le attivita siano chiuse
			sql = "SELECT (COUNT(*)) AS N_ATTIVITA FROM tb_attivita WHERE NOT "& SQL_IsTrue(conn, "att_conclusa") & _
				  " AND att_pratica_id="& cIntero(request("ID"))
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			if rs("N_ATTIVITA") > 0 then
				Session("ERRORE") = "Per archiviare la pratica chiudere tutte le attivit&agrave;!"
				response.redirect "PraticaMod.asp?ID="& request.querystring("ID")
			end if
			rs.close
			archiviata = 1
		end if
		sql = "SELECT * FROM tb_pratiche WHERE pra_id=" & cIntero(request("ID"))
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		rs("pra_nome") = request.form("tft_pra_nome")
		rs("pra_archiviata") = archiviata
		rs("pra_note") = request.form("tft_pra_note")
		rs("pra_mod_data") = NOW()
		rs("pra_mod_utente") = Session("ID_ADMIN")
		rs.update
		rs.close
	end if

	set rs = nothing
end if

if Session("ERRORE")<>"" then
	conn.RollbackTrans
	response.write Session("ERRORE")
	response.end
else
	conn.commitTrans
	response.redirect "Pratiche.asp"
end if


Sub InserimentoAttivita(conn, ID, contatto)
	dim rsa, fso, sql, prefisso, conta, att_id

	set rsa = server.createobject("adodb.recordset")
	
	'GESTIONE INSERIMENTO AL
	'inserisco AL di default
	CALL AL_ins(conn, AL_DEFAULT, ID, false)
	
	'inserisco l'attivita' principale con AL che eredita
	'l'inserimento modifica anche l'AL della pratica
	sql = "SELECT * FROM tb_attivita WHERE att_pratica_id=" & ID
	rsa.open sql, conn, AdOpenKeySet, adLockOptimistic, adCmdText
	rsa.AddNew
	rsa("att_oggetto") = "INIZIO PRATICA """& request("tft_pra_nome") &""""
	rsa("att_dataCrea") = Now
	rsa("att_conclusa") = true
	rsa("att_sistema") = true
	rsa("att_pratica_id") = ID
	rsa("att_mittente_id") = Session("ID_ADMIN")
	rsa("att_dataChiusa") = Now
	rsa.Update

	att_id = rsa("att_id")
	CALL AL_ins (conn, AL_ATTIVITA, att_id, true)
	
	'inserisco l'eventuale prima attivita'
	if request.form("tft_att_oggetto") <> "" AND request.form("tft_att_testo") <> "" then
		rsa.AddNew
		rsa("att_oggetto") = request.form("tft_att_oggetto")
		rsa("att_testo") = request.form("tft_att_testo")
		rsa("att_note") = request.form("tft_att_note")
		rsa("att_dataCrea") = Now
		if IsDate(request.form("tfd_att_dataS")) then
			rsa("att_dataS") = ConvertForSave_Date(request.form("tfd_att_dataS"))
		end if
		rsa("att_priorita") = (request.form("chk_att_priorita") <> "")
		rsa("att_conclusa") = (request.form("chk_att_conclusa") <> "")
		rsa("att_sistema") = false
		rsa("att_inSospeso") = false
		rsa("att_pratica_id") = ID
		rsa("att_mittente_id") = Session("ID_ADMIN")
		rsa.Update
		att_id = rsa("att_id")
		CALL AL_ins (conn, AL_ATTIVITA, att_id, true)
	end if
	rsa.close
				
	'GESTIONE CODICE AUTOCOMPOSTO della pratica
	if Application("NextCom_codice") <> "" then
		sql = "SELECT PraticaCount, PraticaPrefisso FROM tb_indirizzario WHERE IDElencoIndirizzi="& contatto
		rsa.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		rsa("PraticaCount") = cInteger(rsa("PraticaCount")) + 1
		rsa.update
		
		dim pra_codice
		
		'esegue composizione codice
		pra_codice = cString(replace(Application("NextCom_codice"), "<ANNO>", Year(Date)))
		pra_codice = replace(pra_codice, "<MESE>", Month(Date))
		pra_codice = replace(pra_codice, "<GIORNO>", Day(Date))
		pra_codice = replace(pra_codice, "<COUNTCLIENTE>", cString(rsa("PraticaCount")))
		pra_codice = replace(pra_codice, "<PREFISSOCLIENTE>", cString(rsa("PraticaPrefisso")))
		pra_codice = replace(pra_codice, "<ID>", ID)
		
		sql = "UPDATE tb_pratiche SET Pra_codice='" & ParseSQL(pra_codice, adChar) & "' WHERE pra_id=" & ID
		CALL conn.execute(sql, , adExecuteNoRecords)
		
		rsa.close
	end if
	
	set rsa = nothing
	
end Sub

%>