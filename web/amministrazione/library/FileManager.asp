<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% 'response.buffer = False
 %>
<!--#INCLUDE FILE="Tools.asp" -->
<!--#INCLUDE FILE="Tools4Admin.asp" -->
<!--#INCLUDE FILE="filemanager/class_directory.asp" -->

<%
if request("STANDALONE")<>"" AND request("FILEMAN_AZ_ID")<>"" then 
	'imposta i parametri di apertura in modalita' singola (selezione di un file).
	Session("SELECTFILE") = TRUE
    Session("SELECT_OBJECT_TYPE") = request("OBJECT_TYPE")
	Session("FILE_TYPE_FILTER") = request("file_type_filter")
	Session("FORM_NAME") = request("form_name")
	Session("FIELD_NAME") = request("field_name")
	Session("FIELD_ID") = request("field_id")
	Session("ABS_PATH") = request("abs_path")
	Session("SELECTED") = request("selected")
	Session("LOCK") = request("lock")
	Session("FILTER") = request("filter")
	'variabili gestione inserimento immagini e relative relazioni con record
	session("RS_URL") = request("RS_URL")
	session("RS_TAB") = request("RS_TAB")
	session("RS_ID") = request("RS_ID")
elseif request("FILEMAN_AZ_ID")<>"" then
	'parametri per apertura in next-Web
	Session("SELECTFILE") = FALSE
    Session("SELECT_OBJECT_TYPE") = ""
	Session("FORM_NAME") = ""
	Session("FIELD_NAME") = ""
	Session("FIELD_ID") = ""
	Session("ABS_PATH") = ""
	Session("SELECTED") = ""
	Session("FILTER") = request("filter")
	session("RS_URL") = ""
	session("RS_TAB") = ""
	session("RS_ID") = ""
end if

dim path, pathAdmin
path = request("F")

'path dell'amministratore
dim conn
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
pathAdmin = GetValueList(conn, NULL, "SELECT admin_dir FROM tb_admin WHERE id_admin = "& session("ID_ADMIN"))
conn.close
set conn = nothing

if CString(pathAdmin) <> "" then
	pathAdmin = "\images"& pathAdmin
	Session("LOCK") = pathAdmin
	if InStr(1, path, pathAdmin, vbTextCompare) = 0 then
		path = pathAdmin
	end if
elseif instr(1,path,"..")>0 or path="" then
	if Session("LOCK")<>"" then
		path = Session("LOCK")
	else
		path="\images"
	end if
end if

if Request("STANDALONE")<>"" AND Session("FMPATH")<>"" AND request("F")<>"" then
	'in caso di apertura piu' volte della stessa finestra ricorda il percorso dell'apertura precedente se il tipo lo permette
	if instr(1, Session("FMPATH"), request("F"), vbTextCompare)>0 then
		path = Session("FMPATH")
	end if
end if

if instr(1,path,"images", vbTextCompare)>0 then
	Session("file_type") = "images"
elseif instr(1,path,"testi", vbTextCompare)>0 then
	Session("file_type") = "testi"
elseif instr(1,path,"objects", vbTextCompare)>0 then
	Session("file_type") = "objects"
elseif instr(1,path,"flash", vbTextCompare)>0 then
	Session("file_type") = "flash"
end if
Session("FMPATH") = replace(Trim(path), "\\", "\")

dim d, NotValidFilesCount
set d = new directory
d.RelativeDirPath = Trim(replace(Session("FMPATH"), "/", "\"))
path = d.DIRPath

%>
<html>
<head>
	<title><%= ChooseValueByAllLanguages(Session("LINGUA"), "Gestione file", "File management", "", "", "", "", "", "")%></title>
	<link rel="stylesheet" type="text/css" href="stili.css">
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<script src="utils.js" language="JavaScript" type="text/javascript"></script>
<script src="filemanager/ua.js" language="JavaScript" type="text/javascript"></script>
<script src="filemanager/ftiens4.js" language="JavaScript" type="text/javascript"></script>
<script language="JavaScript" type="text/javascript">
	<!--#include file="filemanager/nodes.asp"-->
</script>
<script language="JavaScript" type="text/javascript">
	function Open_Delete(type,file){
		OpenAutoPositionedWindow('filemanager/FileRemove.asp?FILE=' + file, 'FileRemove', 405, 195);
	}
	
	function Open_Upload(type){
	<%	if GetNextWebCurrentVersion(null, null) >= 5 then %>
	    OpenAutoPositionedWindow('../../amministrazione2/filemanager/FileUpload.aspx?PATH=<%= Server.UrlEncode(Session("FMPATH")) %>', 'FileUpload', 405, 260);
	<%	else %>
        OpenAutoPositionedWindow('filemanager/FileUpload.asp?TYPE=' + type + '&FILE=', 'FileUpload', 405, 240);
	<%	end if %>
	}
	
	function Open_Multi(type) {
	    OpenAutoPositionedWindow('../../amministrazione2/filemanager/FileMultiUpload.aspx?PATH=<%= Server.UrlEncode(Session("FMPATH")) %>', 'FileMultiUpload', 600, 600);
	}
	
	function Open_UploadImage(){
		OpenAutoPositionedWindow('../../amministrazione2/filemanager/ImageUpload.aspx?PATH=<%= Server.UrlEncode(Session("FMPATH")) %>', 'FileUpload', 405, 380);
	}
	
	function Open_File(f){
		OpenAutoPositionedScrollWindow('http://'+f, "open_file", 405, 210, true);
	}
	
	function Open_NewDir(){
		OpenAutoPositionedWindow('filemanager/FolderNew.asp', 'FolderNew', 405, 210);
	}
	
	function Open_DeleteDirectory(name, dirpath){
		OpenAutoPositionedWindow('filemanager/FolderRemove.asp?FOLDER=' + name + '&FMPATH=' + dirpath, 'FolderRemove', 405, 195);
	}
	
	<% if Session("SelectFile") AND Session("SELECT_OBJECT_TYPE")<>"" AND session("RS_URL") = "" then %>
		function SelectFile(FileRelativeURL){
			if (FileRelativeURL.indexOf(" ")<0){
				<% if Session("FIELD_NAME") <> "" then %>
					var form = opener.document.getElementById('<%= Session("FORM_NAME") %>')
					form.<%= Session("FIELD_NAME") %>.value= FileRelativeURL;
				<% else %>
					<% if Session("ABS_PATH") = "true" then %>
						FileRelativeURL = '<%=GetSiteUrlImages()%>' + FileRelativeURL.substring(1, FileRelativeURL.length);
					<% end if %>
					window.opener.document.getElementById('<%= Session("FIELD_ID") %>').value= FileRelativeURL;
					// scateno l'evento onChange automaticamente
					var element
					element = window.opener.document.getElementById('<%= Session("FIELD_ID") %>');
					if ("fireEvent" in element){
						element.fireEvent("onchange");
					}
					else if(document.createEventObject){
						element.target.fireEvent("onchange");
					}
					else
					{
						var evt = window.opener.document.createEvent("HTMLEvents");
						evt.initEvent("change", false, true);
						element.dispatchEvent(evt);
					}
				<% end if %>
				window.close();
			}
			else{
				alert("File non valido: il nome contiene spazi.");
			}
		}
	<% end if %>
</script>
<% if Session("SELECTFILE") AND Session("SELECT_OBJECT_TYPE")<>"" then %>
	<body leftmargin="5" topmargin="4" rightmargin="4" onload="window.focus();">
<% else %>
	<body leftmargin="0" topmargin="0" rightmargin="0">
<% end if
Dim fso, f, f1, fc, s
dim FoldersCount, FileCount, TotalSize, extension, SelectionURL, Index
TotalSize = 0
Set fso = CreateObject("Scripting.FileSystemObject")
%>
<table cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td valign="top" width="22%">
			<table cellspacing="0" cellpadding="0" class="tabella_madre">
				<caption class="border"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Cartelle: ", "Folders: ", "", "", "", "", "", "")%></caption>
				<tr>
					<td>
						<div style="height:398px; width:160px; overflow:auto; padding-left:2px;">
							<script language="JavaScript" type="text/javascript">
								initializeDocument()
							</script>
							<noscript>
							A tree for site navigation will open here if you enable JavaScript in your browser.
							</noscript>
						</div>
					</td>
				</tr>
			</table>
		</td>
		<td valign="top" style="padding-left:4px;">
			<table cellspacing="1" cellpadding="0" class="tabella_madre">
				<caption><%= ChooseValueByAllLanguages(Session("LINGUA"), "Oggetti nella cartella ", "Object in folder ", "", "", "", "", "", "")%>"<%= Session("FMPATH") %>"</caption>
				<tr>
					<td class="header">
						<table border="0" cellspacing="0" cellpadding="0" align="right">
							<tr>
								<td style="text-align:right; padding-bottom:2px;">
									<a href="javascript:void(0);" class="button" onclick="Open_NewDir()" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apre la finestra per la creazione di una nuova cartella", "open the window to create a new folder", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
										<%= ChooseValueByAllLanguages(Session("LINGUA"), "NUOVA CARTELLA", "NEW FOLDER", "", "", "", "", "", "")%></a>
									<% 	if session("RS_URL") = "" then %>
								<% response.write session("RS_URL") %>	
									<%		if GetNextWebCurrentVersion(null, null) >= 5 then %>
										<a href="javascript:void(0);" class="button" onclick="Open_UploadImage()" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apre la finestra per caricare una nuova immagine", "open the window to upload a new image", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
											<%= ChooseValueByAllLanguages(Session("LINGUA"), "NUOVA IMMAGINE", "NEW IMAGE", "", "", "", "", "", "")%></a>
										<a href="javascript:void(0);" class="button" onclick="Open_Multi('<%=Session("file_type")  %>')" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apre la finestra per caricare un nuovo file", "open the window to upload a new file", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
											<%= ChooseValueByAllLanguages(Session("LINGUA"), "CARICA NUOVI FILES", "UPLOAD NEW FILES", "", "", "", "", "", "")%></a>
									<%		else %>
										<a href="javascript:void(0);" class="button" onclick="Open_Upload('<%=Session("file_type")  %>')" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apre la finestra per caricare un nuovo file", "open the window to upload a new file", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
											<%= ChooseValueByAllLanguages(Session("LINGUA"), "NUOVO FILE", "NEW FILE", "", "", "", "", "", "")%></a>
									<%		end if %>
									<%	end if %>
									<% if Session("SELECTFILE") then %>
										<a href="javascript:void(0);" class="button" onclick="window.close();" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "chiudi la finestra", "close the window", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
											<%= ChooseValueByAllLanguages(Session("LINGUA"), "CHIUDI", "CLOSE", "", "", "", "", "", "")%></a>
									<% end if %>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<div style="height:378px; width:100%; overflow:auto;">
							<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
								<tr>
									<th colspan="2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "NOME", "NAME", "", "", "", "", "", "")%></th>
									<th><%= ChooseValueByAllLanguages(Session("LINGUA"), "TIPO", "TYPE", "", "", "", "", "", "")%></th>
									<th><%= ChooseValueByAllLanguages(Session("LINGUA"), "DIMENSIONE FILE", "FILE SIZE", "", "", "", "", "", "")%></th>
									<th><%= ChooseValueByAllLanguages(Session("LINGUA"), "MODIFICA", "MODIFY", "", "", "", "", "", "")%></th>
									<th class="center">&nbsp;</th>
								</tr>
								<% 
								'listing delle directory
								' on error resume next
								
								'response.write path
								Set f = fso.GetFolder(path)
								' if Err.Number>0 then
								'	Set f = fso.GetFolder(d.AbsoluteBasePath &"\"& Session("file_type"))
								' end if
								'on error goto 0
								Set fc = f.SubFolders
								FoldersCount = fc.Count
								
								
								if ClearPath(Session("FMPATH"))<>"images" AND ClearPath(Session("FMPATH"))<>"objects" AND  ClearPath(Session("FMPATH"))<>"testi" then	
									if ClearPath(Session("FMPATH")) <> ClearPath(Session("LOCK")) then %>
									<tr>
										<td class="content" style="width:4%;">
											<a href="?F=<%= d.parentpath %>" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apri cartella superiore", "open containing folder", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
												<img src="../grafica/filemanager/FileIcon_folderUp.gif" alt="" border="0">
											</a>
										</td>
										<td class="content" colspan="5">
											<a href="?F=<%= d.parentpath %>" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apri cartella superiore", "open containing folder", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
												<%= ChooseValueByAllLanguages(Session("LINGUA"), "livello superiore...", "containing folder...", "", "", "", "", "", "")%>
											</a>
										</td>
									</tr>
								<% end if %>
									<tr>
										<td class="header"  style="width:4%;" title="directory corrente">
											<img src="../grafica/filemanager/FileIcon_openfolder.gif" alt="<%= ChooseValueByAllLanguages(Session("LINGUA"), "directory corrente", "current directory", "", "", "", "", "", "")%>" border="0">
										</td>
										<td class="header" colspan="4" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "directory corrente", "current directory", "", "", "", "", "", "")%>">
											<%= replace(Session("FMPATH"), "\", " \ ") %>
										</td>
										<td class="header_right" nowrap title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "directory corrente", "current directory", "", "", "", "", "", "")%>">
                                            <% CALL Write_DirectoryOperations(d.foldername(Session("FMPATH")), d.parentpath,"/" + d.RelativeURL(""),"button", false) %>
										</td>
									</tr>
								<%end if
								Index = 0
								
								For Each f1 in fc 
									Index = Index + 1
									if Index MOD 200 = 0 then
										response.flush
									end if %>
									<tr>
										<td class="content_center" width="2%">
											<a href="?F=<%= d.RelativePath %>\<%= f1.name  %>" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apri cartella", "open folder", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
												<img src="../grafica/filemanager/FileIcon_folder.gif" alt="" border="0">
											</a>
										</td>
										<td class="content">
											<a href="?F=<%= d.RelativePath %>\<%= f1.name  %>" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apri cartella", "open folder", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
												<%= f1.name  %>
											</a>
										</td>
										<td class="content"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Cartella di file", "File folder", "", "", "", "", "", "")%></td>
										<td class="content">
											&nbsp;
											<% 
											'.............................................................................................
											'Nicola, 23/02/2010
											'commentato per velocizzare filemanager: per recuperare la dimensione rallenta molto il caricamento.
											'.............................................................................................
											'= File_Dimension(f1.size)  
											%>
										</td>
										<td class="content"><%= f1.DateCreated %></td>
										<td class="content_right" nowrap>
                                            <%CALL Write_DirectoryOperations(f1.name, Session("FMPATH"), d.RelativeURL(f1.name), "button", true) %>
           							    </td>
									</tr>
									<% 
									'.............................................................................................
									'Nicola, 23/02/2010
									'commentato per velocizzare filemanager: per recuperare la dimensione rallenta molto il caricamento.
									'.............................................................................................
									'TotalSize = TotalSize + f1.size
								next 
								'listing dei files
								Set fc = f.Files
								FileCount = fc.Count
								NotValidFilesCount = 0
						
										
								For Each f1 in fc 
									Index = Index + 1
									if index mod 500 = 0 then
										Response.flush()
									end if
									extension = File_Extension( f1.name )%>
									<tr>
										<td class="content_center">
											<a href="http://<%= d.URLPath(f1.name) %>" target="open_file" onclick="Open_File('<%= JsEncode(d.URLPath(f1.name),"'")  %>');" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "click per aprire il file", "click to open the file", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
												<% if instr(1, f1.name, " ", vbTextCompare)>0 then 
													NotValidFilesCount = NotValidFilesCount + 1%>
													<img src="../grafica/filemanager/FileIcon_NotValid.gif" alt="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Documento non valido: presente uno spazio nel nome.", "Invalid document: there's a blank space in the name.", "", "", "", "", "", "")%>" border="0">
												<% else %>
													<img src="../grafica/filemanager/<%= File_Icon( Extension ) %>" alt="<%= ChooseValueByAllLanguages(Session("LINGUA"), "doppio click per aprire il file", "double click to open the file", "", "", "", "", "", "")%>" border="0">
												<% end if %>
											</a>
										</td>
										<td class="content">
											<a href="http://<%= d.URLPath(f1.name) %>" target="open_file" onclick="Open_File('<%= JsEncode(d.URLPath(f1.name),"'")  %>');" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "click per aprire il file", "click to open the file", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
												<%= f1.name  %>
											</a>
										</td>
										<td class="content"><%= File_Type( Extension ) %></td>
										<td class="content"><%= File_Dimension(f1.size)  %></td>
										<td class="content"><%= f1.DateCreated %></td>
										
										<td class="content_right" nowrap>
											<% if Session("SelectFile") AND Session("SELECT_OBJECT_TYPE") = FILE_SYSTEM_FILE then
												SelectionURL = d.RelativeURL(f1.name)												
												if Ucase(SelectionURL) = Ucase(Session("SELECTED")) then%>
													<a onclick="SelectFile('<%= JsEncode(SelectionURL,"'") %>')" class="button_disabled" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "file attualmente selezionato", "file now selected", "", "", "", "", "", "")%>"
												   	   href="javascript:void(0);" style="margin-right:2px;" <%= ACTIVE_STATUS %>>
														<%= ChooseValueByAllLanguages(Session("LINGUA"), "SELEZIONA", "SELECT", "", "", "", "", "", "")%></a>
													<a class="button_disabled" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "impossible cancellare il file perch&egrave; gi&agrave; selezionato", "unable to delete file because it's already selected", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
														<%= ChooseValueByAllLanguages(Session("LINGUA"), "CANC", "DEL", "", "", "", "", "", "")%></a>
												<% else 
													if Session("FILE_TYPE_FILTER") = "" OR instr(1, Session("FILE_TYPE_FILTER"), Extension, vbTextCompare)>0 then%>
														<a onclick="SelectFile('<%= JsEncode(SelectionURL,"'") %>')" class="button" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "click per selezionare il file", "click to select the file", "", "", "", "", "", "")%>"
													   	   href="javascript:void(0);" style="margin-right:2px;" <%= ACTIVE_STATUS %>>
															<%= ChooseValueByAllLanguages(Session("LINGUA"), "SELEZIONA", "SELECT", "", "", "", "", "", "")%></a>
													<% end if %>
													<a onclick="Open_Delete('<%= Session("file_type") %>','<%= JsEncode(f1.name,"'") %>');" class="button" href="javascript:void(0);" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apri la finestra per la cancellazione del file", "open new window to delete the file", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
													<%= ChooseValueByAllLanguages(Session("LINGUA"), "CANC", "DEL", "", "", "", "", "", "")%></a>
												<% end if
											else %>
												<a onclick="Open_Delete('<%= Session("file_type") %>','<%= JsEncode(f1.name,"'") %>');" class="button" href="javascript:void(0);" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apri la finestra per la cancellazione del file", "open new window to delete the file", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
												<%= ChooseValueByAllLanguages(Session("LINGUA"), "CANC", "DEL", "", "", "", "", "", "")%></a>
											<% end if %>
										</td>
									</tr>
									<% TotalSize = TotalSize + f1.size
								next %>
								<tr>
									<th style="border-bottom:0px;" colspan="3">
										<% if FoldersCount > 0 OR FileCount > 0 then 
											if FoldersCount > 0 then %>
												<%= ChooseValueByAllLanguages(Session("LINGUA"), "N&ordm; " & FoldersCount & " cartelle", "N&ordm; " & FoldersCount & " folders", "", "", "", "", "", "")%> 
											<% end if
											if FoldersCount > 0 AND FileCount > 0 then %>,&nbsp;
											<% end if
											if FileCount > 0 then %>
												<%= ChooseValueByAllLanguages(Session("LINGUA"), "N&ordm; " & FileCount & " files", "N&ordm; " & FileCount & " files", "", "", "", "", "", "")%> 
											<% end if %>.
										<% else %>
											<%= ChooseValueByAllLanguages(Session("LINGUA"), "Nessun oggetto trovato.", "No object found.", "", "", "", "", "", "")%> 
										<% end if %>
									</th>
									<th style="border-bottom:0px;"><%= File_Dimension(TotalSize) %></th>
									<th style="border-bottom:0px;">&nbsp;</th>
									<th style="border-bottom:0px;">&nbsp;</th>
								</tr>
								<tr><td colspan="6">&nbsp;</td></tr>
							</table>
						</div>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<% if NotValidFilesCount>0 then %>
		<tr>
			<td colspan="2" style="padding-top:4px;">
				<table cellspacing="0" cellpadding="1" class="tabella_madre" style="border-bottom:0px;">
					<tr>
						<th>
							<img src="../grafica/filemanager/FileIcon_NotValid.gif" alt="">
						</td>
						<th>
							<%= ChooseValueByAllLanguages(Session("LINGUA"), "Files non validi: il nome del file contiene almeno uno spazio. Tali files non potranno essere usati per la costruzione delle pagine.", "Invalid files: the name of the file contains at least one blank space. These files cannot be used to build the pages.", "", "", "", "", "", "")%>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	<% end if %>
</table>
</body>
</html>
<% 
'funzione che ripulisce il path passato come paramentro da eventuali barre per confronti interni
function ClearPath(PathToClear)
	ClearPath = replace(PathToClear, "\", "")
	ClearPath = replace(ClearPath, "/", "")
end function

function Write_DirectoryOperations(dirName, dirPath, selectionUrl, buttonClass, IsDirEmpty)
	if Session("SelectFile") AND Session("SELECT_OBJECT_TYPE") = FILE_SYSTEM_DIRECTORY then
		if Ucase(SelectionURL) = Ucase(Session("SELECTED")) then%>
			<a onclick="SelectFile('<%= JsEncode(SelectionURL,"'") %>')" class="<%= buttonClass %>_disabled" title="la directory &quot;<%= SelectionURL %>&quot; &egrave; attualmente selezionata"
			   href="javascript:void(0);" style="margin-right:2px;" <%= ACTIVE_STATUS %>>
				<%= ChooseValueByAllLanguages(Session("LINGUA"), "SELEZIONA", "SELECT", "", "", "", "", "", "")%>
			</a>
			<a class="<%= buttonClass %>_disabled" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "impossibile cancellare la directory &quot;" & SelectionURL & "&quot; perch&egrave; gi&agrave; selezionata", "impossible to delete directory " & SelectionURL & " because it's already selected", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
				<%= ChooseValueByAllLanguages(Session("LINGUA"), "CANC", "DEL", "", "", "", "", "", "")%>
			</a>
	<% 	else %>
			<a
	<% 		'se non scelta directory per file upload delle immagini
			if session("RS_URL") = "" then %>
			   onclick="SelectFile('<%= JsEncode(SelectionURL,"'") %>')"
			   href="javascript:void(0);"
	<% 		else %>
			   href="../../amministrazione2/filemanager/FileMultiUpload.aspx?PATH=<%= "\" + TrimChar(dirPath, "\") + "\" + dirName %>&RS_URL=<%= session("RS_URL") %>&RS_TAB=<%= session("RS_TAB") %>&RS_ID=<%= session("RS_ID") %>"
	<% 		end if %>
			   class="<%= buttonClass %>" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "click per selezionare la directory &quot;" & SelectionURL & "&quot;", "click to select directory &quot;" & SelectionURL & "&quot;", "", "", "", "", "", "")%>"
			   style="margin-right:2px;" <%= ACTIVE_STATUS %>>
				<%= ChooseValueByAllLanguages(Session("LINGUA"), "SELEZIONA", "SELECT", "", "", "", "", "", "")%>
			</a>
	<% 		if IsDirEmpty then %>
				<a onclick="Open_DeleteDirectory('<%= JsEncode(dirName, "'") %>', '<%= JsEncode(dirPath, "'") %>');" class="<%= buttonClass %>" href="javascript:void(0);" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apre la finestra per la cancellazione della cartella &quot;" & SelectionURL & "&quot;", "open the window to delete folder &quot;" & SelectionURL & "&quot;", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
				    <%= ChooseValueByAllLanguages(Session("LINGUA"), "CANC", "DEL", "", "", "", "", "", "")%>
				</a>
	<% 		else %>
				<a class="<%= buttonClass %>_disabled" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "impossible cancellare la directory &quot;" & SelectionURL & "&quot; perch&egrave; contiene files o directory", "unable to delete folder &quot;" & SelectionURL & "&quot; because it contains files or directories", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
					<%= ChooseValueByAllLanguages(Session("LINGUA"), "CANC", "DEL", "", "", "", "", "", "")%>
				</a>
	<% 		end if
		end if
	else
		if IsDirEmpty then %>
			<a onclick="Open_DeleteDirectory('<%= JsEncode(dirName, "'") %>', '<%= JsEncode(dirPath, "'") %>');" class="<%= buttonClass %>" href="javascript:void(0);" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "apri la finestra per la cancellazione della cartella &quot;" & SelectionURL & "&quot;", "open new window to delete folder &quot;" & SelectionURL & "&quot;", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
				<%= ChooseValueByAllLanguages(Session("LINGUA"), "CANC", "DEL", "", "", "", "", "", "")%>
			</a>
		<% else %>
			<a class="<%= buttonClass %>_disabled" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "impossible cancellare la directory &quot;" & SelectionURL & "&quot; perch&egrave; contiene files o directory", "unable to delete folder &quot;" & SelectionURL & "&quot; because it contains files or directories", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
				<%= ChooseValueByAllLanguages(Session("LINGUA"), "CANC", "DEL", "", "", "", "", "", "")%>
			</a>
		<% end if
	end if
end function

%>