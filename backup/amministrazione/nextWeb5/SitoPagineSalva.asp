<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 10000000 %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
dim IDP
if request("ID") <> "" then
	IDP = cIntero(request("ID"))
elseif request("INDICE") <> "" then
	IDP = GetValueList(index.conn, NULL, "SELECT idx_link_pagina_id FROM tb_contents_index WHERE idx_id="& cIntero(request("INDICE")))
end if

'check dei permessi dell'utente
if CIntero(IDP) > 0 then
	if NOT index.content.ChkPrmF("tb_pagineSito", IDP) then
		response.redirect "SitoPagine.asp"
	end if
elseif NOT index.ChkPrm(prm_pagine_altera, 0) then
	response.redirect "SitoPagine.asp"
end if

dim Classe
	Set Classe = New OBJ_Salva
	
	'imposta id della pagina in corso di modifica
	Classe.ID_value = CIntero(IDP)
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	if NOT index.ChkPrm(prm_indice_trasparente, 0) then
		Classe.Requested_Fields_List = "idx"
	end if
	dim i, lingua
	for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
		lingua = Application("LINGUE")(i)
		if Session("LINGUA_" & lingua) then 
			Classe.Requested_Fields_List	= "tft_nome_ps_" & lingua & "; " & Classe.Requested_Fields_List
		end if
	next
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_pagineSito"
	Classe.id_Field					= "id_pagineSito"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
	
	'update campi gestione modifiche
	classe.SetUpdateParams("ps_")

'definizione eventuali operazioni su relazioni
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql,i, lingua, ContentId
	
	sql = "SELECT * FROM tb_pagineSito WHERE id_paginesito=" & ID
	rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
	
	
	if cInteger(request("ID"))=0 then
		'inserimento nuova pagina
	
		if cIntero(request("pagina_da_copiare"))<>0 then
			'copia da altra pagina sito	
			rs.close				
			CALL CopiaPaginaSito(conn, request("pagina_da_copiare"), ID, false)
			sql = "SELECT * FROM tb_pagineSito WHERE id_paginesito=" & ID
			rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

		else
			'impostazione template
			'verifica e creazione delle pagine corrispondenti alle lingue (tb_pages)
			CALL Ceck_page_exists(conn, rs)
			
			'verifica associazione automatica template alle pagine
			select case request("selezione_template")
				case "unico"
					'impostazione template unico a tutte le pagine
					if cInteger(request("sel_template_unico"))>0 then
						sql = "UPDATE tb_pages SET id_template=" & request("sel_template_unico") & " WHERE id_PaginaSito=" & ID & " AND ( lingua LIKE '" & join(Session("LINGUE"), "' OR lingua LIKE '") & "')"
						CALL conn.execute(sql, 0, adExecuteNoRecords)
					end if
				case "lingue"
					'impostazione template per ogni lingua
					for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
						if Session("LINGUA_" & Application("LINGUE")(i)) then
							if cInteger(request("sel_template_" & Application("LINGUE")(i)))>0 then
								sql = "UPDATE tb_pages SET id_template=" & ParseSQL(request("sel_template_" & Application("LINGUE")(i)), adChar) & _
									  " WHERE id_PaginaSito=" & ID & " AND lingua='" & Application("LINGUE")(i) & "' "
								CALL conn.execute(sql, 0, adExecuteNoRecords)
							end if
						end if
					next
			end select
		end if
	end if
	
	CALL PaginaSitoUpdatePages(conn, IDP)
	
	'inserimento voce nell'indice
	if IDP = "" AND CIntero(request("idx")) > 0 then
		ContentId = Index_UpdateItem(conn, "tb_pagineSito", ID, true)
		index.dizionario("idx_content_id") = ContentId
		index.dizionario("idx_link_tipo") = lnk_interno
		index.dizionario("idx_link_pagina_id") = ID
		index.dizionario("idx_padre_id") = request("idx")
		index.dizionario("idx_principale") = true
		index.dizionario("idx_visibile") = false
		CALL index.Salva(0)
	else
		'in modifica: aggiorna solo contenuto
		ContentId = Index_UpdateItem(conn, "tb_pagineSito", ID, false)
	end if
	
	rs.close
	
	classe.isReport = false
	if request.form("elenco") <> "" then
		if request("FROM") = FROM_ALBERO then
			Classe.Next_Page = "SitoPagineAlbero.asp"
		else
			Classe.Next_Page = "SitoPagine.asp"
		end if
	else
		Classe.Next_Page = "SitoPagineMod.asp?FROM="& request("FROM") &"&ID=" & ID
	end if
end Sub

	'salvataggio/modifica dati
Classe.Salva()
	
	

%>