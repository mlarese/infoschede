<%
dim ParentFrameName, bodyClass

if Session("LOGIN_4_LOG")="" then
	'utente non loggato
	%>
	<script language="JavaScript" type="text/javascript">
		//esegue reload della finestra padre
		try { parent.location.reload(true);}
		catch(e){/*istruzione messa solo per sintassi*/}
	</script>
	<%response.end
end if%>

<!--#INCLUDE FILE="Tools.asp" -->
<!--#INCLUDE FILE="Tools4Admin.asp" -->
<!--#INCLUDE FILE="class_testata.asp" -->
<html>
<head>
	<title><%= Session("NOME_APPLICAZIONE") %></title>
	<link rel="stylesheet" type="text/css" href="<%= GetLibraryPath() %>stili.css">
	<SCRIPT LANGUAGE="javascript" src="<%= GetLibraryPath() %>utils.js" type="text/javascript"></SCRIPT>
	<% if not IsHttpsActive() then %>
		<script src="http://script.aculo.us/prototype.js" type="text/javascript"></script>
		<script src="http://script.aculo.us/scriptaculous.js" type="text/javascript"></script>
		<!--<script src="<%= GetLibraryPath() %>lightbox/js/lightbox.js" type="text/javascript"></script>-->
	<% end if %>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
	<meta http-equiv=Content-Type content="text/html; charset=utf-8">
</head>
<body <%=IIF(bodyClass <> "", " class=""" & bodyClass & """ ", "")%> leftmargin="0" topmargin="0" onload="SetParentFrameHeight('<%=ParentFrameName%>');" onresize="SetParentFrameHeight(<%=ParentFrameName%>);" style="margin-right:0px; margin-bottom:0px;">
<!-- barra alta -->
<script language="JavaScript" type="text/javascript">
	window.onresize += function(event) {
		SetParentFrameHeight('<%=ParentFrameName%>');
	}
	
</script>
<div class="content_iframeform" id="iframeform">