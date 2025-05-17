<%@LANGUAGE="VBSCRIPT"%>
<% option explicit %>
<!--#INCLUDE FILE="../TOOLS.ASP"-->
<!--#INCLUDE FILE="../ClassConfiguration.asp"-->
<% 
dim Config
set Config = new Configuration
'impostazione delle proprieta' di default
Config.AddDefault "file", ""

'caricamento proprieta' specifiche
Config.SetConfigurationString(Session("CONFSTR"))

dim fso, Path, File
set fso = Server.CreateObject("scripting.filesystemobject")
Path = Application("IMAGE_PATH") & "\" & Session("AZ_ID") & "\testi\" & Config("File")

if fso.FileExists(Path) then
	set File = fso.OpenTextFile(Path, 1, false)%>
	<%= File.ReadAll %>
	<%File.close
	set File = nothing
else
	'file non trovato%>
	<div class="noRecords">Fle <%= Config("file") %> non trovato.</div>
<%end if
set fso = nothing

set Config = nothing%>