<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<%
'imposta costanti per i descrittori
tipiDisable = adBinary & ", "
UseSingleLanguage = true

'.................................................................................................
'.................................................................................................
'FUNZIONI E PROCEDURE
'.................................................................................................
'.................................................................................................

'.................................................................................................
'..			scrive automaticamente la parte di form per la gestione dei files allegati
'..			conn:			connessione al database aperta
'..			rs:				oggetto recordset chiuso
'..			DOC_ID:			eventuale id del documento di cui recuperare l'elenco di files
'.................................................................................................
sub GestioneFilesAssociati(conn, rs, DOC_ID)
	dim FilesID, FilesName, FilesOld, sql
	if DOC_ID = "" then
		'inserimento nuovo documento
		FilesID = request("documenti_id_list")
		FilesName = request("documenti_view_list")
		FilesOld = ""
	else
		FilesID = ""
		FilesName = ""
		sql = "SELECT F_ID, F_original_name FROM tb_files INNER JOIN rel_documenti_files " + _
			  " ON tb_files.F_ID = rel_documenti_files.rel_files_id " + _
			  " WHERE rel_documento_id=" & cIntero(DOC_ID)
		rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		while not rs.eof
			FilesID = FilesID & rs("F_ID") & ";"
			FilesName = FilesName & JSFileName(rs("F_original_name")) & "; "
			rs.movenext
		wend
		FilesOld = FilesID
		rs.close
	end if
	%>
	<table cellspacing="0" cellpadding="0">
		<tr>
			<td>
				<input type="hidden" name="old_documenti_id_list" value="<%= FilesOld %>">
				<input type="hidden" name="documenti_id_list" value="<%= FilesID %>">
				<input READONLY type="text" name="documenti_view_list" value="<%= FilesName %>" size="94"
					   onclick="OpenAutoPositionedScrollWindow('DocumentiFiles.asp?Open=1&DOC_ID=<%= DOC_ID %>', 'gestione_files', 640, 400, true)">
			</td>
			<td>
				<a class="button_input" href="javascript:void(0)" onclick="form1.documenti_view_list.onclick();" title="Apre la gestione files per associarli al documento caricandone di nuovi o selezionandoli dagli esistenti" <%= ACTIVE_STATUS %>>
					SELEZIONE FILES
				</a>
			</td>
		</tr>
	</table>
<%end sub

'.................................................................................................
'..			funzione che codifica il nome del file per la selezione via javascript
'..			text:		testo da codificare come nome del file
'.................................................................................................
function JSFileName(text)
	text = replace(text, "^", "")
	text = replace(text, "'", "")
	text = replace(text, """", "")
	text = replace(text, "(", "")
	text = replace(text, ")", "")
	JSFileName = text
end function


'.................................................................................................
'..			genera in visualizzazione l'elenco dei files associati ad un documento
'..			conn:		connessione aperta sul database
'..			rs:			recordset creato ma chiuso
'..			doc_id:		id del documento del quale elencare i files
'.................................................................................................
sub ElencoFileAssociati(conn, rs, DOC_ID)
	dim sql, Extension
	sql = "SELECT tb_files.* FROM tb_files INNER JOIN rel_documenti_files ON tb_files.F_id = rel_documenti_files.rel_files_id " & _
		  " WHERE rel_documenti_files.rel_documento_id=" & cIntero(DOC_ID)
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText%>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
		<% if rs.eof then %>
			<tr><td class="content">Nessun file allegato</td></tr>
		<% else %>
			<tr>
				<th class="L2" colspan="2">NOME</th>
				<th class="L2" width="5%">TIPO</th>
				<th class="l2_center" style="width:50px;">DIM.</th>
				<th class="l2_center" style="width:90px;">DATA MODIFICA</th>
			</tr>
			<% while not rs.eof 
				Extension = File_Extension(rs("F_original_name"))%>
				<tr>
					<td class="content_center" style="width:18px;"><img src="../grafica/filemanager/<%= File_Icon( Extension ) %>" title="<%= File_Type(Extension) %>"></td>
					<td class="content">
						<a href="DocumentiFilesView.asp?ID=<%= rs("F_ID") %>" target="_blank" 
						   title="Apre il file &quot;<%= rs("F_original_name") %>&quot; in una nuova finestra." <%= ACTIVE_STATUS %>>
							<%= rs("F_original_name") %>
						</a>
					</td>
					<td class="content" nowrap><%= File_Type(Extension) %></td>
					<td class="content_right"><%= File_Dimension(rs("F_size"))  %></td>
					<td class="content_center"><%= DateTimeITA(rs("F_Data")) %></td>
				</tr>
				<% rs.movenext
			wend %>
		<% end if %>
	</table>
	<% rs.close
end sub
%>