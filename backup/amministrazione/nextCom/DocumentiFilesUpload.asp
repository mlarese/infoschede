<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4admin.asp" -->
<!--#INCLUDE FILE="../library/ClassUpload.asp" -->
<html>
<head>
	<meta http-equiv=Content-Type content="text/html; charset=utf-8">
	<title>Upload Nuovo file</title>
	<link rel="stylesheet" type="text/css" href="../library/stili.css">
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0">
	<table cellspadding="0" cellspacing="0" width="100%">
		<tr>
			<td class="content">
				<%dim UPL_OBJ
				SET UPL_OBJ = New UploadFile
				UPL_OBJ.File_Path = Application("IMAGE_PATH") & "temp\docs\" & Session("LOGIN_4_LOG")
				UPL_OBJ.ShowConsigli = false
				UPL_OBJ.OnlyExtensionAllowed = false
				UPL_OBJ.OverWrite = true
				UPL_OBJ.Stile_Submit = "class=""button_L2"""
				UPL_OBJ.Stile_Titoli = "class=""content_b"""
				UPL_OBJ.Stile_testo	= "class=""content"""
				UPL_OBJ.Upload()
				if UPL_OBJ.OperationOK then%>
					<script language="JavaScript">
						parent.location.reload(true);
					</script>
				<%end if%>
			</td>
		</tr>
	</table>
</body>
</html>