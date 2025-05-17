<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/ClassContent.asp" -->
<%
dim content
set content = new ObjContent
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tfn_tab_sito_id;tft_tab_titolo;tft_tab_name;tft_tab_field_chiave;tft_tab_field_titolo_it;tft_tab_from_sql"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_siti_tabelle"
	Classe.id_Field					= "tab_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, rst, var
	set rst = server.CreateObject("ADODB.Recordset")
	
	sql = " SELECT COUNT(*) FROM tb_siti_tabelle WHERE tab_sito_id = " & request("tfn_tab_sito_id") & _
		  " AND tab_name LIKE '" & request("tft_tab_name") & "' AND "&SQL_IfIsNull(conn, "tab_priorita_base", "0")&" = " & cIntero(request("tfn_tab_priorita_base")) & _
		  " AND tab_id <> " & ID
	if cIntero(GetValueList(conn, NULL, sql)) > 0 then
		Session("ERRORE") = "Cambiare priorit&agrave; base. Esiste già una tabella con la stessa origine e la stessa priorit&agrave;."
	else
		'esegue una query di test per verificare se tutti i dati immessi sono corretti
		sql = " SELECT TOP 1 " + _
				request("tft_tab_field_chiave")
		
		for each var in request.form
			if instr(1, var, "tft_tab_field_", vbTextCompare)>0 AND _
			   instr(1, var, "_chiave", vbTextCompare)<1 AND _
			   instr(1, var, "tft_tab_field_return", vbTextCompare)<1 then
				sql = sql + AddField(var)
			end if
		next
		
		sql = sql + " FROM " & request("tft_tab_from_sql") %>
		Apertura della query di test per verificare se i dati immessi sono corretti.<br>
		<strong>Query</strong>:<br>
		<%= sql %><br>
		<strong>Risultato:</strong><br>
		Se si vede questo messaggio la prova non &egrave; andata a buon fine: controllare i campi immessi.<br>
		<a href="javascript:history.go(-1)">INDIETRO</a><br>
		<%rst.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		rst.close
		
		'gestione immagini
		rs.open "SELECT * FROM tb_siti_tabelle WHERE tab_id = "& ID, conn, adOpenStatic, adLockOptimistic
		if CIntero(request("tfn_tab_thumb")) = 0 then
			rs("tab_thumb") = 0
			rs.update
		end if
		if CIntero(request("tfn_tab_zoom")) = 0 then
			rs("tab_zoom") = 0
			rs.update
		end if
		
		if CIntero(request("tfn_tab_pagina_default_id")) = 0 then
			rs("tab_pagina_default_id") = 0
			rs.update
		end if
		rs.close
		dim immagini, i
		immagini = split(request("immagini"), ",")
		if request("ID") <> "" then
			sql = " DELETE FROM rel_immaginiFormati WHERE rif_tab_id = "& ID
			CALL conn.execute(sql, 0, adExecuteNoRecords)
		end if
		for i = lbound(immagini) to ubound(immagini)
			sql = "INSERT INTO rel_immaginiFormati(rif_imf_id, rif_tab_id) VALUES (" & immagini(i) & ", " & ID & ")"
			CALL conn.execute(sql, 0, adExecuteNoRecords)
		next
		
		set rst = nothing
		'imposta parametri per salvare e chiudere la finestra corrente
		Classe.isReport = TRUE%>
		
	<%	sql = "UPDATE tb_webs SET webs_modData_tabelle = " & SQL_Now(conn) 
		CALL conn.execute(sql, ,adExecuteNoRecords)
	%>

		<script language="JavaScript" type="text/javascript">
			opener.location.reload(true);
			window.close();
		</script>	
	<%
	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()

	
function AddField(FormField)
	dim i, FieldName
	FieldName = request(FormField)
	AddField = ""
	if FieldName <> "" then
		AddField = ", (" + FieldName + ") AS " + FormField
	end if
end function
%>