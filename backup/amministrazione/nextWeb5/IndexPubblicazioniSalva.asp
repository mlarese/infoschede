<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/tools.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_pub_titolo;tfn_pub_tabella_id;tfn_pub_padre_index_id"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_siti_tabelle_pubblicazioni"
	Classe.id_Field					= "pub_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
	
	if request("tipo_url")<>"" then
		CALL Classe.AddForcedValue("pub_pagina_id", NULL)
	end if

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	if (request("tfn_pub_categoria_tabella_id") = "" OR request("tft_pub_categoria_field") = "") _
	   AND CIntero(request("tfn_pub_padre_index_id")) = 0 then
	   	session("ERRORE") = "Scegliere la categoria o la posizione nell'indice!"
	elseif cInteger(request("tfn_pub_pagina_id"))=0 AND request("tipo_url")="" then	
		session("ERRORE") = "Scegliere la pagina per la composizione dell'url."
	else
		'codice di controllo per verificare congruita' della pagina con il ramo scelto.
		if cInteger(request("tfn_pub_pagina_id"))<>0 then
			sql = " SELECT COUNT(*) FROM tb_webs " + _
				  " WHERE id_webs IN ( SELECT id_web FROM tb_paginesito WHERE id_pagineSito=" & cInteger(request("tfn_pub_pagina_id")) & " ) " + _
				  "    OR id_webs IN ( SELECT idx_webs_id FROM tb_contents_index WHERE idx_id = " & CIntero(request("tfn_pub_padre_index_id")) & ") "
			if cIntero(GetValueList(conn, rs, sql)) > 1 then
				Session("ERRORE") = "La pagina scelta non appartiene allo stesso sito del ramo in cui verr&agrave; pubblicata."
				exit sub
			end if
		end if

		dim rst, sql, var
		set rst = server.CreateObject("ADODB.Recordset")
		
		'verifica della query immessa
		sql = Pubblicazione_GetQuery(conn, rst, ID, 0) 
		%>
		Apertura della query di test per verificare se i dati immessi sono corretti.<br>
		<strong>Query</strong>:<br>
		<%= sql %><br>
		<strong>Risultato:</strong><br>
		Se si vede questo messaggio la prova non &egrave; andata a buon fine: controllare i campi immessi.<br>
		<form action="<%= request.ServerVariables("HTTP_REFERRER") %>" method="post">
			<% for each var in request.form
				if instr(1, var, "salva", vbTextCompare)<1 then %>
					<input type="hidden" name="<%= var %>" value="<%= request.form(var) %>">
				<%end if
			next %>
			<input type="Submit" name="indietro" value="indietro"
		</form>
		
		<%rst.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		rst.close
		set rst = nothing

		'imposta parametri per passare alla pagina successiva
		Classe.isReport = FALSE
        if request("salva")<>"" then
            Classe.Next_Page = "IndexPubblicazioniMod.asp?ID=" & ID
        else
		    Classe.Next_Page = "IndexPubblicazioni.asp"
        end if
	end if
end Sub

'salvataggio/modifica dati
Classe.Salva()


%>