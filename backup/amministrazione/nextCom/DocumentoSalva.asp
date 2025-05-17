<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="Tools_DocumentiFiles.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<%
dim pratica, gloID
pratica = request.form("tfn_doc_pratica_id")

'controllo accesso
dim conn2
set conn2 = server.createobject("adodb.connection")
conn2.open Application("DATA_ConnectionString")
if request("mod")<>"" OR request("AL")<>"" then		'se sono in modifica
	if (NOT AL(conn2, request("ID"), AL_DOCUMENTI) OR Session("COM_POWER") = "" AND _
	   Session("ID_ADMIN") <> CInt(GetValueList(conn2, NULL, "SELECT doc_creatore_id FROM tb_documenti WHERE doc_id="& cIntero(request("ID"))))) _
	   AND Session("COM_ADMIN") = "" then
		response.redirect "documenti.asp"
	end if
else												'sono in inserimento
	if pratica <> "" AND pratica <> "0" then
		if NOT AL(conn2, pratica, AL_PRATICHE) then
			response.redirect "Pratiche.asp?ID="& Session("COM_PRA_CLIENTE")
		end if
	end if
end if
conn2.close
set conn2 = nothing

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_doc_nome;tfn_doc_tipologia_id"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_documenti"
	Classe.id_Field					= "doc_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql
	gloID = ID
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = (request.querystring("ATT")<>"")
	
	'GESTIONE INSERIMENTO AL
	if request.form("salva") <> "" then				'se sono in inserimento
		CALL AL_ins(conn, AL_DOCUMENTI, ID, (request.form("ere") <> ""))
	end if
	
	'GESTIONE DESCRITTORI
	CALL DesSalva(conn, ID, "rel_documenti_descrittori", "rdd_valore", "rdd_documento_id", "rdd_descrittore_id")

	'GESTIONE RELAZIONI CON FILE
	if request("old_documenti_id_list") <> request("documenti_id_list") then
		'inserimento documento o variazione files allegati al documento
		dim FileIdList, FileId
		if request("old_documenti_id_list")<>"" then
			'cancella relazioni con vecchi files
			FileIdList = left(replace(request("old_documenti_id_list"), ";", ","), len(request("old_documenti_id_list"))-1)
			sql = "DELETE FROM rel_documenti_files WHERE rel_documento_id=" & ID & " AND rel_files_id IN (" + FileIdList + ")"
			CALL conn.execute(Sql, , adExecuteNoRecords)
		end if
		if request("documenti_id_list")<>"" then
			'inserisce relazioni con nuovi files
			FileIdList = split(request("documenti_id_list"), ";")
			for each FileId in FileIdList
				if cInteger(FileId)>0 then
					sql = "INSERT INTO rel_documenti_files (rel_documento_id, rel_files_id) VALUES (" & ID & ", " & FileID & ")"
					CALL conn.execute(Sql, , adExecuteNoRecords)
				end if
			next
		end if
		'aggiorna stato dei files interessati dalle variazioni
		sql = "SELECT F_ID, F_Allegato, F_original_path FROM tb_files WHERE "
		if request("old_documenti_id_list")<>"" then
			sql = sql & " F_ID IN (" & left(replace(request("old_documenti_id_list"), ";", ","), len(request("old_documenti_id_list"))-1) & ") OR "
		end if
		if request("documenti_id_list")<>"" then
			sql = sql & " F_ID IN (" & left(replace(request("documenti_id_list"), ";", ","), len(request("documenti_id_list"))-1) & ") OR "
		end if
		sql = left(sql, len(sql) - 3)
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		'cicla su file interessati
		while not rs.eof
			sql = "SELECT (COUNT(*)) AS N_ASSOCIAZIONI FROM rel_documenti_files WHERE rel_files_id=" & rs("F_ID")
			rs("F_Allegato") = (conn.execute(sql, , adCmdText)("N_ASSOCIAZIONI") > 0 )
			if not rs("F_allegato") then
				'se il file non e' piu' allegato ad alcun documento lo riporta comune se originariamente era comune, 
				' se e' stato allegato da un utente diventa dell'utente che l'ha deselezionato
				if rs("F_original_path")<>"" then
					rs("F_original_path") = Session("LOGIN_4_LOG")
				end if
			end if
			rs.Update
			rs.movenext
		wend
		rs.close
	end if
	
	Classe.Next_Page = "Documenti.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()

if Session("errore")="" then
%>
<script language="javascript">
	if (opener != null) {
		opener.form1.documenti.value += '<%= gloID %>;'
		opener.form1.visDoc.value += '<%= JSEncode(request.form("tft_doc_nome"), "'") %>;'
		// opener.form1.submit(); 
		window.close(); 
	}
</script>
<% End If 
%>