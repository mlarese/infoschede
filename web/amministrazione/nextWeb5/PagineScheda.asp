<%@ Language=VBScript CODEPAGE=65001%>
<% response.charset = "UTF-8" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Chiusura modifica editor</title>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
	<meta http-equiv=Content-Type content="text/html; charset=utf-8">
</head>

<body>

<script language="JavaScript" type="text/javascript">
	var e;
	try{
		<% if Session("is_NextCom_Page") then 
			'blocco di codice eseguito se la pagina modificata e' una email
			'generata nel next-web
			%>
			opener.SetPreview(<%= request("pagina") %>);
		<% else %>
			window.opener.document.location.reload(true);
		<% end if %>
	} 
	catch(e){
	}
	window.close();
</script>

</body>
</html>
