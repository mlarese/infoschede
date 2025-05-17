<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_DocumentiFiles.asp" -->
<%
dim conn, sql, rs, Extension, ShowPrivateMsg
set conn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.Recordset")
conn.open Application("DATA_ConnectionString")

'scansione cartelle per inserimento files.
'controlla cartella comune
CALL CheckNewFiles(conn, rs, "")
'controlla cartella utente
CALL CheckNewFiles(conn, rs, Session("LOGIN_4_LOG"))

'salvataggio delle modifiche
if request.ServerVariables("REQUEST_METHOD") = "POST" then
	if request("rinomina")<>"" AND cInteger(request("REN_F_ID")) > 0 then
		if request("tft_F_original_name")<>"" _
		   AND CheckChar(request("tft_F_original_name"), DOCUMENTS_FILES_CHARSET & " _-.") _
		   AND Count(request("tft_F_original_name"), ".")=1 then
			sql = "SELECT F_original_name FROM tb_files WHERE F_ID=" & cIntero(request("REN_F_ID"))
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			rs("F_original_name") = request("tft_F_original_name")
			rs.Update
			rs.close
			conn.close
			set rs = nothing
			set conn = nothing
			response.redirect "DocumentiFiles.asp"
		else
			if request("tft_F_original_name")<>"" then
				if Count(request("tft_F_original_name"), ".")=0 then
					Session("ERRORE") = "Nome file non valido: Estensione del file mancante"
				elseif Count(request("tft_F_original_name"), ".")>1 then
					Session("ERRORE") = "Nome file non valido: Impossibile riconoscere l'estrensione perch&egrave; sono presenti pi&ugrave; punti nel nome del file."
				else
					Session("ERRORE") = "Nome file non valido: Utilizzare solo numeri, lettere o i caratteri &quot;_&quot; o &quot;-&quot;"
				end if
			else
				Session("ERRORE") = "Nome del file mancante!"
			end if
		end if
	end if
end if
%>
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<%'--------------------------------------------------------
sezione_testata = "Gestione ed associazione files al documento" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

if request("OPEN") <>"" then
	'prima apertura della finestra
	Session("DocFiles_ass_no") =  "1"
	Session("DocFiles_ass_doc") = request("DOC_ID")
	Session("DocFiles_doc_id") = request("DOC_ID")
end if

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	if request("tutti")<>"" then
		Session("DocFiles_nome") = ""
		Session("DocFiles_ass_doc") = ""
		Session("DocFiles_ass_altri") = ""
		Session("DocFiles_ass_no") = ""
	elseif request("cerca")<>"" then
		Session("DocFiles_nome") = request("search_nome")
		Session("DocFiles_ass_doc") = request("search_ass_doc")
		Session("DocFiles_ass_altri") = request("search_ass_altri")
		Session("DocFiles_ass_no") = request("search_ass_no")
	end if
end if

sql = "SELECT tb_files.*, " 
if Session("DocFiles_doc_id")<>""  then
	sql = sql & " (SELECT COUNT(*) FROM rel_documenti_files WHERE rel_files_id=tb_files.F_ID " & _
				" AND rel_documento_id=" & Session("DocFiles_doc_id") & ") AS N_ASS_DOC, " & _
				" (SELECT COUNT(*) FROM rel_documenti_files WHERE rel_files_id=tb_files.F_ID " & _
				" AND rel_documento_id <> " & Session("DocFiles_doc_id") & ") AS N_ASS_ALTRI "
else
	sql = sql & " (0) AS N_ASS_DOC, " & _
				" (SELECT COUNT(*) FROM rel_documenti_files WHERE rel_files_id=tb_files.F_ID) AS N_ASS_ALTRI "
end if
sql = sql & " FROM tb_files WHERE ("

'aggiunge filtro su files non associati ad alcun documento
if Session("DocFiles_ass_no")<>"" OR _
   (Session("DocFiles_ass_no")="" AND Session("DocFiles_ass_doc")="" AND Session("DocFiles_ass_altri")="") then
   sql = sql & "( not " & SQL_isTrue(conn, "F_Allegato") & " AND (F_original_path='' OR F_original_path='" & Session("LOGIN_4_LOG") & "')) OR "
end if

'aggiunge filtro su files associati solo al documento corrente
if session("DocFiles_doc_id")<>"" AND (Session("DocFiles_ass_doc")<>"" OR _
   (Session("DocFiles_ass_no")="" AND Session("DocFiles_ass_doc")="" AND Session("DocFiles_ass_altri")="")) then
   sql = sql & "( " & SQL_isTrue(conn, "F_Allegato") & " AND " + _
   			   "F_ID IN (SELECT rel_files_id FROM rel_documenti_files WHERE rel_documento_id=" & session("DocFiles_doc_id") & ")) OR "
end if
'aggiunge filtro su files associati solo ad altri documenti
if Session("DocFiles_ass_altri")<>"" OR _ 
   (Session("DocFiles_ass_no")="" AND Session("DocFiles_ass_doc")="" AND Session("DocFiles_ass_altri")="") then
   sql = sql & "( " & SQL_isTrue(conn, "F_Allegato") & " AND " + _
   			   "F_ID IN (SELECT rel_files_id FROM rel_documenti_files WHERE "
	if session("DocFiles_doc_id")<>"" then
		sql = sql & " rel_documento_id <> " & session("DocFiles_doc_id") & " AND "
	end if
	sql = sql & "rel_documento_id IN (SELECT doc_id FROM tb_documenti WHERE " & AL_query(conn, AL_DOCUMENTI) & "))) OR "
end if
sql = left(sql, len(sql)-3) & ")"

if Session("DocFiles_nome")<>"" then
	sql = sql & " AND " + SQL_FullTextSearch(Session("DocFiles_nome"), "f_original_name")
end if

sql = sql & " ORDER BY F_original_name"

rs.Open sql, conn, AdOpenStatic, adLockReadOnly, adCmdText
%>
<script language="JavaScript" type="text/javascript">
	//imposta input da pagina padre
	var IDList, NameList;
	IDList = opener.form1.documenti_id_list;
	NameList = opener.form1.documenti_view_list;
	
	function FileSelection(selection, f_id){
		var objF_name;
		objF_name = document.getElementById('NAME_' + f_id);
		var re = eval('/' + objF_name.value + '; /g');		//espressione regolare per cercare il file tramite nome
		
		if (selection.checked){
			//selezione del file
			//controlla se non e' gia' presente un file con lo stesso nome
			if (!(NameList.value.match(re))){
				//nome file univoco in lista: lo seleziona
				IDList.value += f_id + ";"
				NameList.value += objF_name.value + "; "
			} 
			else {
				//nome file non univoco: non permette la selezione
				selection.checked = false;
				alert('Il file non puo\' essere selezionato perche\' e\' gia\' presente un file con lo stesso nome.');
			}
		}else
		{	//deseleziona il file
			NameList.value = NameList.value.replace(re, '')
			re = eval('/' + f_id + ';/g');
			IDList.value = IDList.value.replace(re, '')
		}
		FileState(f_id, selection.checked);
	}
	
	function FileState(f_id, state){
		var objF_ren, objF_name, objF_sel;
		objF_name = document.getElementById('NAME_' + f_id);
		objF_ren = document.getElementById('REN_' + f_id);
		objF_sel = document.getElementById('SEL_' + f_id);
		objF_sel.checked = state;
		if (state){
			//file selezionato
			objF_ren.className = 'button_L2_disabled';
			objF_ren.href = 'javascript:void(0);';
			objF_ren.title = 'Impossibile rinominare il file perche\' gia\' selezionato.';
		} 
		else {	
			//file deselezionato
			objF_ren.className = 'button_L2';
			objF_ren.href = 'DocumentiFiles.asp?REN_F_ID=' + f_id;
			objF_ren.title = 'Rinomina il file "' + objF_name.value + '"';
		}
	}
	
	function FileDelete(f_id){
		var objF_sel
		objF_sel = document.getElementById('SEL_' + f_id);
		objF_sel.checked = false;
		//deseleziona il file
		FileSelection(objF_sel, f_id);		
		//cancella il file
		OpenDeleteWindow('FILES',f_id);
	}
</script>
<div id="content_ridotto">
<form action="DocumentiFiles.asp" method="post" id="ricerca" name="ricerca">
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption>
		<table border="0" cellspacing="0" cellpadding="1" align="right">
			<tr>
				<td style="font-size: 1px; padding-right:1px;" nowrap>
					<input type="submit" name="cerca" value="CERCA" class="button">
					&nbsp;
					<input type="submit" name="tutti" value="VEDI TUTTI" class="button">
				</td>
			</tr>
		</table>
		Opzioni di ricerca
	</caption>
	<tr>
		<th>NOME FILE</th>
		<th colspan="3">ASSOCIAZIONE CON I DOCUMENTI</th>
	</tr>
	<tr>
		<td class="content" width="32%">
			<input type="text" name="search_nome" value="<%= replace(session("DocFiles_nome"), """", "&quot;") %>" style="width:100%;">
		</td>
		<td class="content warning">
			<input type="Checkbox" name="search_ass_no" class="checkbox" value="1" <%= IIF(session("DocFiles_ass_no")<>"", " checked", "") %>>
			non associati
		</td>
		<% if session("DocFiles_doc_id")<>"" then %>
			<td class="content">
				<input type="Checkbox" name="search_ass_doc" class="checkbox" value="1" <%= IIF(session("DocFiles_ass_doc")<>"", " checked", "") %>>
				<strong>associati al documento</strong>
			</td>
		<% end if %>
		<td class="content">
			<input type="Checkbox" name="search_ass_altri" class="checkbox" value="1" <%= IIF(session("DocFiles_ass_altri")<>"", " checked", "") %>>
			associati ad altri documenti
		</td>
	</tr>
</table>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption class="border">Files disponibili</caption>
	<tr>
		<td>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
				<tr>
					<td class="label_no_width" colspan="8">
						<% if rs.eof then %>
							Nessun file trovato.
						<% else %>
							Trovati n&ordm; <%= rs.recordcount %> files
						<% end if %>
					</td>
				</tr>
				<% if not rs.eof then %>
					<tr>
						<th class="L2">&nbsp;</th>
						<th class="L2" colspan="2">NOME</th>
						<th class="L2" width="5%">TIPO</th>
						<th class="l2_center" style="width:45px;">DIM.</th>
						<th class="l2_center" style="width:83px;">DATA</th>
						<th class="l2_center" colspan="2">OPERAZIONI</th>
					</tr>
					<%ShowPrivateMsg = false
					while not rs.eof 
						Extension = File_Extension(rs("F_original_name"))%>
						<tr>
							<td class="content_center" style="width:16px;">
								<% if cInteger(request("REN_F_ID")) = rs("F_ID") then %>
									&nbsp;
								<% else %>
									<input type="checkbox" class="checkbox" onClick="FileSelection(this, '<%= rs("F_ID") %>')" name="SEL_<%= rs("F_ID") %>" id="SEL_<%= rs("F_ID") %>" value="<%= rs("F_ID") %>">
									<input type="hidden" name="NAME_<%= rs("F_ID") %>" value="<%= JSFileName(rs("F_original_name")) %>">
								<% end if %>
							</td>
							<td class="content_center" style="width:18px;"><img src="../grafica/filemanager/<%= File_Icon( Extension ) %>" title="<%= File_Type(Extension) %>"></td>
							<% if cInteger(request("REN_F_ID")) = rs("F_ID") then %>
								<td class="content">
									<input type="hidden" name="REN_F_ID" value="<%= request("REN_F_ID") %>">
									<input type="text" name="tft_F_original_name" value="<%= rs("F_original_name") %>" class="text" style="width:100%">
								</td>
							<% else
								if rs("N_ASS_DOC")>0 then		'file associato al documento%>
									<td class="content_b">
								<% elseif rs("N_ASS_ALTRI")>0 then %>
									<td class="content">
								<% else %>	
									<td class="content warning">
									<% if rs("F_original_path")=Session("LOGIN_4_LOG") then
										ShowPrivateMsg = true%>
										<table cellpadding="0" cellspacing="0" align="right">
											<tr><td><img src="../grafica/padlock.gif" alt="File privato, non visibile agli altri utenti." title="File privato, non visibile agli altri utenti." <%= ACTIVE_STATUS %>></td></tr>
										</table>
									<% end if
								end if %>
									<a href="DocumentiFilesView.asp?ID=<%= rs("F_ID") %>" target="_blank" 
									   title="Apre il file &quot;<%= rs("F_original_name") %>&quot; in una nuova finestra." <%= ACTIVE_STATUS %>>
										<%= rs("F_original_name") %>
									</a>
								</td>
							<% end if %>
							</td>
							<td class="content" nowrap><%= File_Type(Extension) %></td>
							<td class="content_right"><%= File_Dimension(rs("F_size"))  %></td>
							<td class="content_center"><%= DateTimeITA(rs("F_Data")) %></td>
							<% if cInteger(request("REN_F_ID")) = rs("F_ID") then %>
								<td class="content_center" style="width:61px;">
									<input type="submit" name="rinomina" value="SALVA" class="button_L2" style="width:92%;">
								</td>
								<td class="content_center" style="width:62px;">
									<input type="button" name="annulla" value="ANNULLA" class="button_L2" style="width:95%;" onclick="document.location='DocumentiFiles.asp';">
								</td>
							<% else %>
								<td class="content_center" style="width:61px;">
									<a id="ren_<%= rs("F_ID") %>" class="button_L2" href="DocumentiFiles.asp?REN_F_ID=<%= rs("F_ID") %>" title="Rinomina il file &quot;<%= rs("F_original_name") %>&quot;" <%= ACTIVE_STATUS %>>
										RINOMINA
									</a>
								</td>
								<td class="content_center" style="width:62px;">
									<% if rs("N_ASS_ALTRI")>0 then %>
										<a class="button_L2_disabled" title="Impossibile cancellare il file perch&egrave; associato anche ad altri documenti." <%= ACTIVE_STATUS %>>
									<% else %>
										<a class="button_L2" href="javascript:void(0);"  title="Cancella il file &quot;<%= rs("F_original_name") %>&quot;" <%= ACTIVE_STATUS %> 
											onclick="FileDelete('<%= rs("F_ID") %>');">
									<% end if %>
										CANCELLA
									</a>
								</td>
							<% end if %>
						</tr>
						<% rs.moveNext
					wend
				end if %>
				<tr>
					<th class="L2" colspan="8">UPLOAD DI UN NUOVO FILE</th>
				</tr>
				<tr>
					<td colspan="8">
						<iframe src="DocumentiFilesUpload.asp" frameborder="0" scrolling="No" style="width:100%; height:30px;">
						</iframe>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td class="footer">
			<% if ShowPrivateMsg then %>
				<table cellpadding="0" cellspacing="0" align="left">
					<tr>
						<td><img src="../grafica/padlock.gif" alt="File privati non visibili agli altri utenti." title="File privati non visibili agli altri utenti." <%= ACTIVE_STATUS %>></td>
						<td>&nbsp;&nbsp;File privati non visibili agli altri utenti.</td>
					</tr>
				</table>
			<% end if %>
			<a class="button" href="javascript:window.close();" title="chiudi la finestra" <%= ACTIVE_STATUS %>>
				CHIUDI</a>
		</td>
	</tr>
</table>
</form>
</div>
<script language="JavaScript" type="text/javascript">
	/*
	impostazione della selezione corrente al caricamento della pagina sulla base dei dati 
	contenuti nel campo del form sottostante
	*/
	var aIDList;
	var element;
	aIDList = IDList.value.split(';');		//array di ID dei files selezionati
	for (var i=0; i<(aIDList.length-1); i++){
		element = document.getElementById('SEL_' + aIDList[i]);
		if (element)
			FileState(aIDList[i], true);
	}
</script>
</body>
</html>
<% 
conn.close 
set rs = nothing
set conn = nothing

sub CheckNewFiles(conn, rs, dir)
	dim path, fso, Folder, FolderName, File, sql
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	'calcola path completo
	path = Application("IMAGE_PATH") & "temp\docs\" & dir
	'verifica esistenza cartella da controllare e contenuto 
	if fso.FolderExists(path) then
		set Folder = fso.GetFolder(path)
		if Folder.Files.Count > 0 then
			Conn.BeginTrans
			
			'trovati nuovi files da indicizzare
			sql = "SELECT * FROM tb_files WHERE NOT(" & SQL_isTrue(conn, "F_allegato") & ") AND F_original_path='" & dir & "'"
			rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
			
			'ciclo su files presenti nella cartella per indicizzazione nel sistema
			for each File in Folder.Files
				rs.AddNew
				rs("F_original_path") = dir
				rs("F_original_name") = file.name
				rs("F_encoded_name") = ""
				rs("F_encoded_path") = ""
				rs("F_size") = file.size
				rs("F_data") = NOW()
				rs("F_allegato") = false
				rs("F_LastUpdate") = file.DateLastModified
				rs.Update
				response.flush
				
				'calcola nome univoco del file e lo imposta rinominando anche il file
				file.name = GetUniqueFileName(conn, cString(rs("F_ID")), file.name)
				
				'sposta il file nella cartella di destinazione
				'calcola nome della cartella di destinazione
				FolderName = Year(Date) & "_" & DatePart("ww", Date, vbUseSystem, vbUseSystem)
				
				path = Application("IMAGE_PATH") & "\docs\" & FolderName & "\"
				
				'verifica esistenza directory di destinazione, ed eventualmente la crea
				if not fso.FolderExists(path) then
					fso.CreateFolder(path)
				end if
				
				'sposta il file nella cartella di destinazione 
				file.move(path)
				
				'aggiorna il nome del file e la cartella di destinazione del file
				sql = "UPDATE tb_files SET F_encoded_name='" & file.name & "', F_encoded_path='" & FolderName & "' WHERE F_id=" & rs("F_ID")
				CALL conn.execute(Sql, , adExecuteNoRecords)
				
				rs.Update
			next
			rs.close
			Conn.CommitTrans
		end if
	else
		'non esiste la cartella: errore di configurazione dell'applicazione
		response.write "ERRORE DI CONFIGURAZIONE DEL SISTEMA: cartella " & path & " non trovata <br>Contattare l'amministratore."
		response.end
	end if
	set fso = nothing
end sub

function GetUniqueFileName(conn, FixedPart, FileName)
	dim count, sql, FileExtension
	
	FileExtension = right(FileName, len(FileName) - instrrev(FileName, ".", vbTrue, vbTextCompare))
	
	do 
		'calcola parte random e nome file codificato
		GetUniqueFileName = FixedPart & "_" & GetRandomString(uCase(DOCUMENTS_FILES_CHARSET), 8) & "." & FileExtension
		
		'verifica se sono gia' presenti file codificati con lo stesso nome
		sql = "SELECT (COUNT(*)) AS N_FILES FROM tb_files WHERE F_encoded_name LIKE '" & GetUniqueFileName & "'"
		count = cInteger(conn.execute(sql, 1, adCmdText).fields("N_FILES").value)
	loop until count = 0		'se il conteggio e' a zero il nome del file e' univoco
end function

%>