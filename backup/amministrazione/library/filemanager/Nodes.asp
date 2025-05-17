// You can find instructions for this file here:
// http://www.treeview.net

// Decide if the names are links or just the icons
USETEXTLINKS = 1  //replace 0 with 1 for hyperlinks


// Decide if the tree is to start all open or just showing the root folders
STARTALLOPEN = 1 //replace 0 with 1 to show the whole tree

ICONPATH = '../grafica/filemanager/' //change if the gif's folder is a subfolder, for example: 'images/'
USEFRAMES = 0


foldersTree = gFld("<i>Risorse sito </i>", "javascript:void(0);")
foldersTree.iconSrc = ICONPATH + "FileIcon_webfolder.gif"  
foldersTree.iconSrcClosed = ICONPATH + "FileIcon_webfolder.gif"
<%
dim EditorVersion
'recupera versione NEXT-web installata
EditorVersion = GetNextWebCurrentVersion(NULL, NULL)
const MaxLevelDirectory = 2	'livello massimo di default delle cartelle "aperte" sull'albero a sx
const MaxSubdirectories = 10 'numero massimo di subdirectory visualizzate

if (Session("FILTER")="" OR instr(1, Session("FILTER"), "images", vbtextCompare)>0) AND _
	(Session("LOCK")="" OR instr(1, Session("LOCK"), "images", vbtextCompare)>0) then
	response.write ShowFolderList("images",Application("IMAGE_PATH") & Session("FILEMAN_AZ_ID") & "\",1)
end if
if EditorVersion > 4 then
	'utilizzato dalla versione nextweb5 e successive
	if (Session("FILTER")="" OR instr(1, Session("FILTER"), "flash", vbtextCompare)>0) AND _
		(Session("LOCK")="" OR instr(1, Session("LOCK"), "flash", vbtextCompare)>0) then
		response.write ShowFolderList("flash",Application("IMAGE_PATH") & Session("FILEMAN_AZ_ID") & "\",1)
	end if
end if
if EditorVersion < 4 then
	'non piu' presente per nextweb4 e versioni successive
	if (Session("FILTER")="" OR instr(1, Session("FILTER"), "objects", vbtextCompare)>0) AND _
		(Session("LOCK")="" OR instr(1, Session("LOCK"), "objects", vbtextCompare)>0) then
		response.write ShowFolderList("objects",Application("IMAGE_PATH") & Session("FILEMAN_AZ_ID") & "\",1)
	end if
end if
if (Session("FILTER")="" OR instr(1, Session("FILTER"), "testi", vbtextCompare)>0) AND _
	(Session("LOCK")="" OR instr(1, Session("LOCK"), "testi", vbtextCompare)>0) then
	response.write ShowFolderList("testi",Application("IMAGE_PATH") & Session("FILEMAN_AZ_ID") & "\",1)
end if

Function ShowFolderList(foldername, basepath, level)
	response.write vbCrLf + "/*" + vbCrLF + _
				   vbCrLF + "foldername:" + foldername + vbCrLF + _
				   vbCrLF + "basepath:" + basepath + vbCrLF + _
				   vbCrLF + "level:" & level & vbCrLF + _
				   vbCrLf + "*/"
	
	Dim folderspec,fso, f, f1, fc, s, path
	'corregge eventuali problemi con sottocartelle
	foldername = replace(foldername, "/", "\")
	if left(foldername, 1) = "\" then 
		foldername = right(foldername, len(foldername)-1)
	end if
	folderspec = basepath & foldername
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	if fso.folderExists(folderspec) then
		if level < 2 then
			Set f = fso.GetFolder(folderspec)
			Set fc = f.SubFolders
			if level < 2 or fc.count <= MaxSubdirectories then
				if level=1 then
					s = s & "aux"&level&" = insFld(foldersTree, gFld("""&foldername&""", """ & IIF(CanBrowseFolder(folderspec, level), "filemanager.asp?F=" & relPath( folderspec,Application("IMAGE_PATH") & Session("FILEMAN_AZ_ID") & "\" ), "") & """))" & vbCRLF
					s = s & "aux"&level&".iconSrc = ICONPATH + ""FileIcon_openfolder.gif""" & vbCRLF
					s = s & "aux"&level&".iconSrcClosed = ICONPATH + ""FileTree_folder.gif""" & vbCRLF
				end if
				For Each f1 in fc
					if CanBrowseFolder(folderspec & "\" & f1.name, level) then
						s = s & "aux"&level+1&" = insFld(aux"&level&", gFld("""&f1.name&""", ""filemanager.asp?F="&relPath( folderspec & "\" & f1.name,Application("IMAGE_PATH") & Session("FILEMAN_AZ_ID") & "\" )&"""))" & vbCRLF
						s = s & "aux"&level+1&".iconSrc = ICONPATH + ""FileIcon_openfolder.gif""" & vbCRLF
						s = s & "aux"&level+1&".iconSrcClosed = ICONPATH + ""FileTree_folder.gif""" & vbCRLF
						s = s & ShowFolderList(f1.name,folderspec & "\",level+1)
					end if
				Next
			end if
		end if
		ShowFolderList = s
	else
		ShowFolderList = ""
	end if
End Function

function relPath( path,basepath )
	if instr(1,path,basepath, vbTextCompare) > 0 then
		relPath = server.urlencode(right(path,len(path)-len(basepath)+1))
	else
		relPath = ""
	end if
end function


function CanBrowseFolder(path, level)

	if Session("LOCK") <> "" then
		dim lockedpath, relativePath
		relativePath = Replace(path, Application("IMAGE_PATH") & Session("FILEMAN_AZ_ID"), "")
		LockedPath = replace(Session("LOCK"), "/", "\")
		Path = replace(path, "/","\")
		CanBrowseFolder = instr(1, Path, LockedPath, vbTextCompare) > 0 OR instr(1, lockedPath, relativePath, vbTextCompare) > 0
	else
		if level < MaxLevelDirectory + 1 then
			CanBrowseFolder = true
		else
			CanBrowseFolder = false
		end if
	end if
	
end function


%>