<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<!--#INCLUDE FILE="../library/ClassDelete.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/ClassIndirizzarioLock.asp" -->
<html>
<head>
	<title><%= Session("NOME_APPLICAZIONE") %></title>
	<link rel="stylesheet" type="text/css" href="../library/stili.css">
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0" onload="window.focus()">
<%dim Class_delete
set Class_delete = new OBJ_Delete

'parametri da impostare per sito
Class_delete.Section = request.Querystring("SEZIONE")
Class_delete.ID_Value = request.Querystring("ID")
Class_delete.PageName = "Delete.asp"

Class_delete.ReloadOpener = TRUE
Class_delete.LinkStyle = "class=""button"""
Class_delete.MessageStyle = ""
Class_delete.CaptionStyle = "style=""font-weight:bold;"""
Class_delete.CaptionColor = "#E6E6E6"
Class_delete.BorderDarkColor = "#919191"
Class_delete.BorderLightColor = "#FFFFFF"
Class_delete.BackgroundColor = "#F4F4F4"
Class_delete.DeleteRelations = FALSE
Class_delete.AfterDelete = FALSE

'impostazione dei dati dell'indice
Class_delete.Index = Index

dim rs, sql, aux
set rs = Server.CreateObject("ADODB.Recordset")

'parametri da impostare per ogni sezione
Select case request.Querystring("SEZIONE")
	case "SITI"
		'check dei permessi dell'utente
		CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))
		
		Class_delete.Message = "Cancellare il sito <RECORD> e tutti i template e le pagine in esso presenti?"
		'permette lo storno della disponibilit&agrave; della camera
		Class_delete.AddOption "delete_files", "cancella anche le cartelle ed i files associati", true, ""
		Class_delete.Name_Field = "nome_webs"
		Class_delete.ID_Field = "id_webs"
		Class_delete.Table = "tb_Webs"
		Class_delete.Caption = "Gestione siti"
		Class_delete.DeleteRelations = TRUE
		Class_delete.AfterDelete = TRUE
	case "PAGINE"
		'check dei permessi dell'utente
		CALL CheckAutentication(index.content.ChkPrmF("tb_pagineSito", Class_delete.ID_Value))
		
		Class_delete.Message = "Cancellare la pagina <RECORD>?"
		Class_delete.Name_Field = SQL_PaginaSitoNome(Class_delete.conn, "nome_ps_IT")
		Class_delete.ID_Field = "id_paginesito"
		Class_delete.Table = "tb_paginesito"
		Class_delete.Caption = "Gestione pagine sito"
		Class_delete.DeleteRelations = TRUE
	case "TEMPLATE"
		'check dei permessi dell'utente
		CALL CheckAutentication(index.ChkPrm(prm_template_accesso, 0))
		
		Class_delete.Message = "Cancellare il template <RECORD>? <br> Tutte le pagine associate non verranno cancellate, ma perderanno l'associazione."
		Class_delete.Name_Field = "nomepage"
		Class_delete.ID_Field = "id_page"
		Class_delete.Table = "tb_pages"
		Class_delete.Caption = "Gestione template"
	Case "MENU"
		'check dei permessi dell'utente
		CALL CheckAutentication(index.ChkPrm(prm_menu_accesso, 0))
		
		Class_delete.Message = "Cancellare il menu <RECORD><br> e tutti i links associati?"
		Class_delete.Name_Field = "m_nome_it"
		Class_delete.ID_Field = "m_id"
		Class_delete.Table = "tb_menu"
		Class_delete.Caption = "Gestione siti - menu"
	Case "MENUITEM"
		'check dei permessi dell'utente
		CALL CheckAutentication(index.ChkPrm(prm_menu_accesso, 0))
		
		Class_delete.Message = "Cancellare il link <RECORD>?"
		Class_delete.Name_Field = "mi_titolo_IT"
		Class_delete.ID_Field = "mi_id"
		Class_delete.Table = "tb_menuitem"
		Class_delete.Caption = "Gestione siti - menu - links"
	case "PLUGIN"
		'check dei permessi dell'utente
		CALL CheckAutentication(index.ChkPrm(prm_plugin_accesso, 0))
		
		Class_delete.Message = "Cancellare il plugin <RECORD><br> e tutti i layers in cui viene richiamato?"
		Class_delete.Name_Field = "identif_objects"
		Class_delete.ID_Field = "id_objects"
		Class_delete.Table = "tb_objects"
		Class_delete.Caption = "Gestione plugin"
		Class_delete.DeleteRelations = TRUE
	case "PUBBLICAZIONI"
		'check dei permessi dell'utente
		CALL CheckAutentication(index.ChkPrm(prm_Pubblicazioni_accesso, 0))
		
        sql = " SELECT idx_id FROM tb_contents_index INNER JOIN rel_index_pubblicazioni ON tb_contents_index.idx_id = rel_index_pubblicazioni.rip_idx_id " + _
              " WHERE rip_pub_id = " & cIntero(request("ID"))
        aux = GetValueList(index.conn, rs, sql)
        if aux <> ""  then 'AND DB_Type(index.conn) = DB_SQL
            set rs = index.Discendenti(aux)
            if not rs.eof then
                Class_delete.AddOption "delete_indexes", "cancella anche le voci pubblicate in modo automatico", false, _
                                       "Verranno cancellate n&ordm;" & rs.recordcount & " voci dell'indice comprensive di voci della pubblicazione automatica e loro sotto-voci."
            end if
            rs.close
        end if
        
		Class_delete.Message = "Cancellare la pubblicazione <RECORD>?"
		Class_delete.Name_Field = "pub_titolo"
		Class_delete.ID_Field = "pub_id"
		Class_delete.Table = "tb_siti_tabelle_pubblicazioni"
		Class_delete.Caption = "Pubblicazioni automatiche dei dati  "
		Class_delete.AfterDelete = FALSE
		Class_delete.DeleteRelations = true
	case "IMMAGINIFORMATI"
		'check dei permessi dell'utente
		CALL CheckAutentication(index.ChkPrm(prm_immaginiFormati_accesso, 0))
		
		Class_delete.Message = "Cancellare il formato <RECORD>?"
		Class_delete.Name_Field = "imf_nome"
		Class_delete.ID_Field = "imf_id"
		Class_delete.Table = "tb_immaginiFormati"
		Class_delete.Caption = "Gestione formati immagini"
		Class_delete.DeleteRelations = TRUE
	case "URL"
		'check dei permessi dell'utente
		CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))
		
		Class_delete.Message = "Cancellare l'url <RECORD>?"
		Class_delete.Name_Field = "dir_url"
		Class_delete.ID_Field = "dir_id"
		Class_delete.Table = "tb_webs_directories"
		Class_delete.Caption = "Gestione siti - url alternativi"
		Class_delete.DeleteRelations = FALSE
	
	case "DOMINIO"
		'check dei permessi dell'utente
		CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))
		
		Class_delete.Message = "Cancellare il dominio aggiuntivo <RECORD>?"
		Class_delete.Name_Field = "dom_url"
		Class_delete.ID_Field = "dom_id"
		Class_delete.Table = "tb_webs_domini"
		Class_delete.Caption = "Gestione siti - domini aggiuntivi"
		Class_delete.DeleteRelations = FALSE
		
	case "FILTRI"
		Class_delete.Message = "Cancellare il filtro <RECORD>?"
		Class_delete.Name_Field = "fil_valore"
		Class_delete.ID_Field = "fil_id"
		Class_delete.Table = "tb_contents_log_filtri"
		Class_delete.Caption = "Filtri di esclusione log"
		
	case "META"
		'check dei permessi dell'utente
		CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))
		
		Class_delete.Message = "Cancellare il metatag <RECORD>?"
		Class_delete.Name_Field = "meta_name"
		Class_delete.ID_Field = "meta_id"
		Class_delete.Table = "tb_webs_metatag"
		Class_delete.Caption = "Gestione siti - metatag aggiuntivi"
		Class_delete.DeleteRelations = FALSE
	
	case "RSS"
		'check dei permessi dell'utente
		CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))
		
		Class_delete.Message = "Cancellare l'RSS <RECORD>?"
		Class_delete.Name_Field = "rss_titolo"
		Class_delete.ID_Field = "rss_id"
		Class_delete.Table = "tb_rss"
		Class_delete.Caption = "Gestione siti - rss"
		Class_delete.DeleteRelations = FALSE
	
end select

'definizione eventuali operazioni su relazioni
Sub Delete_Relazioni(conn, ID)
    Dim comID, aux, ObjContatto
    
	Select case request.Querystring("SEZIONE")
	
		case "SITI"
		    dim tabName
            sql = " SELECT co_id, co_F_table_id FROM tb_contents " + _
                  " WHERE co_F_table_id IN (SELECT tab_id FROM tb_siti_tabelle WHERE tab_name = 'tb_pagineSito') " + _
                    " AND co_F_key_id IN (SELECT id_pagineSito FROM tb_pagineSito WHERE id_web = "& ID &") "
            rs.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
            while not rs.eof
				tabName = GetValueList(index.conn, NULL, "SELECT tab_name FROM tb_siti_tabelle WHERE tab_id = "& CIntero(rs("co_F_table_id")))
                index.content.DeleteAll tabName, rs("co_id")
                rs.movenext
            wend
            rs.close
	
		case "TEMPLATE"
		
			sql = "UPDATE tb_pages SET id_template=0 WHERE id_template=" & ID
			conn.execute(Sql)
			
		case "PAGINE"
		    
            dim lingua, ListaPagine
            sql = "SELECT * FROM tb_PagineSito WHERE id_pagineSito=" & ID
            rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
            ListaPagine = ""
            for each lingua in Application("LINGUE")
                ListaPagine = ListaPagine & cIntero(rs("id_pagDyn_" & lingua)) & ", " & _
                                            cIntero(rs("id_pagStage_" & lingua)) & ", "
            next
            
            sql = " DELETE FROM tb_pages WHERE id_page IN (" & left(ListaPagine, len(ListaPagine)-2) & ")"
            CALL conn.execute(sql, 0, adExecuteNoRecords)
            
            'aggiorna data di ultima modifica della struttura delle pagine del sito
            CALL UpdateSitoDataModificaPagine(conn, rs("id_web"))
            rs.close
            
		case "PLUGIN"
		
			sql = "DELETE FROM tb_layers WHERE id_objects=" & ID
			CALL conn.execute(sql, 0, adExecuteNoRecords)
		
        case "PUBBLICAZIONI"
            'cancella gli indici bloccati solo da quella pubblicazione solo se richiesto, altrimenti li sblocca e basta
            if request("delete_indexes")<>"" then
                sql = " SELECT idx_id FROM tb_contents_index " + _
                      " WHERE idx_id IN (SELECT rip_idx_id FROM rel_index_pubblicazioni WHERE rip_pub_id= " & ID & ") " + _
                            " AND idx_id NOT IN (SELECT rip_idx_id FROM rel_index_pubblicazioni WHERE rip_pub_id<>" & ID & ") " + _
                      " ORDER BY idx_livello DESC"
                rs.open sql, conn, adOpenDynamic, adLockOptimistic
                
                while not rs.eof 
                    CALL index.Delete(cInteger(rs("idx_id")))
                    rs.movenext
                wend
                rs.close
            end if
            
            'sblocca gli indici o gli eventuali indici bloccati da altre pubblicazioni
            sql = "DELETE FROM rel_index_pubblicazioni WHERE rip_pub_id=" & ID
            CALL conn.execute(sql)
            
            sql = " UPDATE tb_contents_index SET idx_autopubblicato=0 " + _
                  " WHERE idx_id NOT IN (SELECT rip_idx_id FROM rel_index_pubblicazioni) "
            CALL conn.execute(sql)

    End Select
end Sub

Sub Operations_AfterDelete(conn, ID)
	dim fso
	
	Select case request.Querystring("SEZIONE")
	
		case "SITI"
		
			if request("delete_files")<>"" then
				set fso = Server.CreateObject("Scripting.FileSystemObject")
				CALL FolderRemove(fso, Application("IMAGE_PATH") & ID, false)
				set fso = nothing
			end if
		
	end select
	
end sub

Class_delete.Delete_Manager()
%>

</body>
</html>