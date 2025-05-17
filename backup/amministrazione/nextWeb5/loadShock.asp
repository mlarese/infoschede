<%@ Language=VBScript CODEPAGE=65001%>
<% option explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/TOOLS.ASP" -->
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<% Response.Buffer = false 

'check dei permessi dell'utente
dim conn
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
if NOT index.ChkPrm(prm_pagine_altera, 0) then
	conn.close
	set conn = nothing %>
	<script language="JavaScript">
		window.close()
	</script>
	<%
end if
conn.close
set conn = nothing


'......................................................
const editor_dcr_name = "Layers5_3.dcr"
'......................................................


dim http
if instr(1,Request.ServerVariables("HTTPS"),"on",vbTextCompare) then
	http = "https://"
else
	http = "http://"
end if
%>
<!DOCTYPE HTML>
<html>
<head>
	<title>Modifica layout della pagina <%= cInteger(request("PAGINA")) %></title>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
	<meta http-equiv=Content-Type content="text/html; charset=utf-8">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <% if request("NEXTEMAIL")="" then %>onunload="window.opener.document.location.reload(true);"<% end if %>>
<object classid="clsid:233C1507-6A77-46A4-9443-F871F945D258"
 codebase="http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=11,0,0,0"
 ID=layers5sviluppo11 width=1600 height=3000 VIEWASTEXT>
	<param name=src value="<%= editor_dcr_name %>">								<% 'nome del file compilato dell'editor %>
	<param name=swRemote value="swSaveEnabled='true' swVolume='true' swRestart='true' swPausePlay='true' swFastForward='true' swContextMenu='false' ">
	<param name=swStretchStyle value=stage>
	<param name=swUrl value="<%= GetCurrentBaseUrl() %>">						<% 'percorso dello script corrente %>
	<param name=swText value="<%= cInteger(request("PAGINA")) %>">				<% 'id della pagina corrente %>
	<PARAM NAME=bgColor VALUE=#FFFFFF> 
	<PARAM NAME=swStretchHAlign VALUE=Left> 
	<PARAM NAME=swStretchVAlign VALUE=Top>
	<param name=PlayerVersion value=11>
	<embed src="<%= editor_dcr_name %>" 										<% 'nome del file compilato dell'editor %>
	 bgColor=#FFFFFF  
	 width=1600 
	 height=3000 
	 swRemote="swSaveEnabled='true' swVolume='true' swRestart='true' swPausePlay='true' swFastForward='true' swContextMenu='false' " 
	 swStretchStyle=none
	 type="application/x-director"
	 PlayerVersion=11
	 pluginspage="<%= http %>www.macromedia.com/shockwave/download/"
	 swUrl="<%= GetCurrentBaseUrl() %>"											<% 'percorso dello script corrente %>
	 swText="<%= cInteger(request("PAGINA")) %>">								<% 'id della pagina corrente %>
	</embed>
</object>
<p>&nbsp;</p>
</body>
</html>