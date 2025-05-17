<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1073741824 %>
<% Titolo_sezione = "Import documenti Memo 2"
Action = "INDIETRO"
href = "default.asp"%>
<!--#include file="Intestazione.asp"-->
<% 
dim conn, rs, sql, categoria
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>Import documenti e categorie a partire da una cartella</caption>
        <tr><th colspan="3">PARAMETRI DI IMPORT</th></tr>
		<% if request("importa")="" AND request("file_import")="" then %>
            <form action="" method="post" id="form1" name="form1">
			<tr>
				<td class="label" style="width:18%;">directory padre:</td>
				<td class="content" colspan="2">
					<% CALL WriteFileSystemPicker_Input(Application("AZ_ID"), FILE_SYSTEM_DIRECTORY, "images", "", "form1", "file_import", request("file_import"), "width:400px", true, true) %>
					<% 
					'CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "file_import", request("file_import") , "width:400px;", true) 
					%>
                    <span class="note">Selezionare la directory che contiene le cartelle e/o i file da importare nel memo 2.</span>
				</td>
			</tr>
        </table>
        <table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
			<tr>
				<td class="footer" colspan="3">
					(*) Campi obbligatori.
					<input style="width:20%;" type="submit" class="button" name="importa" value="AVANTI &gt;&gt;">
				</td>
			</tr>
			</form>
        <% else %>
            <tr>
				<td class="label" style="width:18%;">directory di partenza:</td>
				<td class="content">
					<%= request("file_import") %>
				</td>
			</tr>
			<tr><th colspan="2">ESECUZIONE IMPORT</th></tr>
			<% 
			
			'apertura transazione di import
			conn.begintrans
			
			dim FilePath, basePath, ConnectionString, RubricaId, CntId, ListaRecapiti, Recapito, Recapiti, rec, note, Field, Value
            dim objFSO, objFolder, objFile, colFiles

			basePath = Application("IMAGE_PATH") & Application("AZ_ID") & "\images\"
			FilePath = basePath & request("file_import")
			FilePath = Replace(FilePath, "/", "\")
            FilePath = Replace(FilePath, "\\", "\")

			
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			Set objFolder = objFSO.GetFolder(FilePath)
			'------
			CALL CreateRecordByPath(objFolder.Path, objFolder)
			'------
			Set colFiles = objFolder.Files
			For Each objFile In colFiles
				'scorre i file figli della directory principale
				'------
				CALL CreateRecordByPath(objFile.Path, objFile)
				'------
			Next
			ShowSubFolders(objFolder)

			Sub ShowSubFolders(objFolder)
				dim colFolders, objSubFolder, objFile, colFiles
				Set colFolders = objFolder.SubFolders
				For Each objSubFolder In colFolders
					'------
					CALL CreateRecordByPath(objSubFolder.Path, objSubFolder)
					'------
					Set colFiles = objSubFolder.Files
					For Each objFile In colFiles
						'------
						CALL CreateRecordByPath(objFile.Path, objFile)
						'------
					Next
					ShowSubFolders(objSubFolder)
				Next
			End Sub
			
			Sub CreateRecordByPath(path, obj)
				dim lastPart, secondLastPart, partList, id_padre, i, id_cat

				path = Replace(path, basePath, "")
				
				partList = Split(path, "\")
				i = uBound(partList)
				lastPart = partList(i)
				secondLastPart = ""
				if i > 0 then
					secondLastPart = partList(i - 1)
				else
					'entra in questo caso l'inserimento della prima categoria base (ovvero la cartella scelta con il file manager)
					secondLastPart = lastPart
					lastPart = ""
				end if
				
				
				if secondLastPart <> "" then
					id_cat = CreateCategoria(secondLastPart, 0)
				else
					id_cat = 0
				end if
				
				if lastPart <> "" then
					if inStr(lastPart, ".") > 0 then
						'è un file
						CALL CreateDocumento(lastPart, id_cat, path, obj)
					else
						'è una cartella
						id_cat = CreateCategoria(lastPart, id_cat)
					end if
				end if
				
				%>
				<tr>
					<td class="content"><%= secondLastPart %></td>
					<td class="content"><%= lastPart %></td>
				</tr>
				<%
				
			End Sub
			
			
			Function CreateCategoria(nome, id_cat_padre)
				sql = " SELECT * FROM mtb_documenti_categorie WHERE catC_nome_it LIKE '"&ParseSql(nome, adChar)&"'"
				if cIntero(id_cat_padre) > 0 then
					sql = sql & " AND catC_padre_id = " & id_cat_padre
				end if
				rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
				if rs.eof then
					rs.addNew
					rs("catC_nome_it") = nome
					rs("catC_ordine") = cIntero(GetValueList(conn, NULL, "SELECT MAX(catC_ordine) FROM mtb_documenti_categorie")) + 5
					'rs("catC_foglia") = 
					'rs("catC_livello") = 
					if cIntero(id_cat_padre) > 0 then
						rs("catC_padre_id") = cIntero(id_cat_padre)
					end if
					rs("catC_visibile") = 1
					rs("catC_albero_visibile") = 1
					rs.Update
				end if
				CreateCategoria = rs("catC_id")
				rs.close
			End Function
			
			Sub CreateDocumento(nome, id_cat, path, oFile)
				sql = " SELECT * FROM mtb_documenti WHERE doc_titolo_it LIKE '"&ParseSql(nome, adChar)&"' AND doc_categoria_id = " & id_cat
				rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
				if rs.eof then
					rs.addNew
					rs("doc_titolo_it") = nome
					rs("doc_visibile") = 1
					rs("doc_protetto") = 0
					rs("doc_pubblicazione") = DateIso(Now())
					rs("doc_categoria_id") = id_cat
					rs("doc_catalogo_sfogliabile") = 0
					rs("doc_file_it") = "/" & Replace(path, "\", "/")
					CALL SetUpdateParamsRS(rs, "doc_", true)
					rs("doc_insData") = oFile.DateCreated
					rs("doc_modData") = oFile.DateLastModified
					rs.Update
				end if
				rs.close
			End Sub
			
			Sub RebuildTree()
				dim rs_cat, categorie
				set rs_cat = Server.CreateObject("ADODB.RecordSet")
				set categorie = New objCategorie
				with categorie
					set .conn = conn
					.tabella = "mtb_documenti_categorie"
					.prefisso = "catC"
				end with
				
				'per ogni categoria
				rs_cat.open "SELECT catC_id FROM mtb_documenti_categorie", conn, adOpenForwardOnly, adLockOptimistic
				while not rs_cat.eof
					categorie.operazioni_ricorsive_tipologia(rs_cat("catC_id"))
					rs_cat.movenext
				wend
				rs_cat.close
				set rs_cat = nothing
				set categorie = nothing
			End Sub
			
			
			CALL RebuildTree()

			'chiusura transazione di import
			conn.committrans 
			%>
			<tr>
				<td class="content_b" colspan="3">IMPORT DATI COMPLETATO</td>
			</tr>
			<tr>
				<td class="footer" colspan="6">
					<a class="button" href="default.asp">FINE</a>
				</td>
			</tr>
			<% 
		end if

        %>
	</table>
</div>

<%
conn.close
set rs = nothing
set conn = nothing
%>