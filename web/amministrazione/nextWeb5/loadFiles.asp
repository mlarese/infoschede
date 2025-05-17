<%@ Language=VBScript CODEPAGE=65001%>
<% option explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/TOOLS.ASP" -->
<% response.buffer = true %>
<%
dim page_id, web_id
page_id = cInteger(request("PAGINA"))


dim conn, rs, sql
set conn = server.createObject("ADODB.Connection")
set rs = server.createObject("ADODB.RecordSet")
conn.Open Application("DATA_ConnectionString"),"",""

sql = "SELECT id_webs FROM tb_pages WHERE id_page=" & page_id
web_id = cInteger(GetValueList(conn, rs, sql))

dim path, fso, dir, FileType
Set fso = CreateObject("Scripting.FileSystemObject")

dir = replace(request("DIR"), "*", "\")
path = Application("IMAGE_PATH") & web_id & dir

'deternina tipo di file
if instr(1, Left(dir, 10), "testi", vbTextCompare)>0 then
	FileType = "TEXT"
elseif instr(1, Left(dir, 10), "flash", vbTextCompare)>0 then
	FileType = "FLASH"
else
	FileType = "IMAGES"
end if

if not fso.FolderExists(path) then
	'directory richiesta non trovata: restituisce uno schema vuoto
	response.clear
	response.ContentType = "text/plain"
	response.write "vuoto"
else
	'directory richiesta non trovata: restituisce la lista
	dim folder, f, FilesList
	Set folder = fso.GetFolder(path)
	
	'genera elenco files
	FilesList = ""
	for each f in folder.Files
		if CheckFile(f.name) then
			FilesList = FilesList & f.name & VbCrLf
		end if
	next
	response.ContentType = "text/xml"
	response.clear
	%><?xml version="1.0" encoding="UTF-8"?>
	<xml>
		<dir nome="<%= replace(dir, "\", "/") %>">
			<% for each f in folder.SubFolders 
				if instr(1, Trim(f.name), " ", VbTextCompare)<1 AND instr(1, Trim(f.name), "&", VbTextCompare)<1 then%>
					<directory><%= f.name %></directory>
				<% end if
			next %>
			<% if FilesList<>"" then %>
				<files><%= Server.HtmlEncode(left(FilesList, len(FilesList)-2)) %></files>
			<% end if %>
		</dir>
	</xml>
	<%set folder = nothing
	set fso = nothing
end if

conn.close
set rs = nothing
set conn = nothing


'verifica se il file e' corretto (non contiene spazi)
function CheckFile(fName)
	CheckFile = (instr(1, Trim(fName), " ", VbTextCompare)<1)
	if CheckFile then
		dim Extension
		Extension = File_Extension( fName )
		if FileType = "TEXT" then
			'verifica che il file sia effettivamente di tipo testo
			CheckFile = instr(1, EXTENSION_TEXT, " " & Extension & " " , vbTextCompare)>0
		elseif FileType = "FLASH" then
			'verifica che il file sia effettivamente di tipo flash
			CheckFile = instr(1, EXTENSION_FLASH, " " & Extension & " " , vbTextCompare)>0
		else
			'verifica che il file sia di tipo immagine
			CheckFile = instr(1, EXTENSION_IMAGES, " " & Extension & " " , vbTextCompare)>0
		end if
	end if
end function
%>