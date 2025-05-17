<%@ Language=VBScript CODEPAGE=65001%>
<% option explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="class_directory.asp" -->
<!--#INCLUDE FILE="../tools.asp" -->
<!--#INCLUDE FILE="../tools4Admin.asp" -->
<!--#INCLUDE FILE="../class_testata.asp" -->
<!DOCTYPE HTML>

<html>
<head>
	<title><%= ChooseValueByAllLanguages(Session("LINGUA"), "Rimozione cartella", "Removing folder", "", "", "", "", "", "")%></title>
	<link rel="stylesheet" type="text/css" href="../stili.css">
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>

<body leftmargin="0" topmargin="0" onload="window.focus()">
<% 
	dim intestazione
	'disegna intestazione
	set intestazione= New testata
	intestazione.sezione = ChooseValueByAllLanguages(Session("LINGUA"), "Rimozione cartella", "Removing folder", "", "", "", "", "", "")
	intestazione.scrivi_ridotta()
	
dim d, path, errore
set d = new directory
d.RelativePath = request("fmpath")
path = request("FOLDER")

	
	response.write "<!-- " + vbCrLf + _
				   "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>" + vbCrLf + _
				   path + vbCrLf + _
				   " -->"
				   
If Request.ServerVariables("REQUEST_METHOD") = "POST" then
	
	if path<>"" then
		if d.ExistFolder(path) then
			if not d.RemoveFolder(path) then
				errore = ChooseValueByAllLanguages(Session("LINGUA"), "La cartella non &egrave; vuota", "The folder is not empty", "", "", "", "", "", "")
			end if
		else
			errore = ChooseValueByAllLanguages(Session("LINGUA"), "La cartella non esiste", "The folder does not exist", "", "", "", "", "", "")
		end if
	else
		errore = ChooseValueByAllLanguages(Session("LINGUA"), "Nome cartella non valido", "Invalid folder name", "", "", "", "", "", "")
	end if
end if
%>

<div id="content_ridotto">
	<form method="POST" action="" id="form1" name="form1">
	<input type="hidden" name="fmpath" id="fmpath" value="<%= request("FMPATH") %>">
	<table cellspadding="0" cellspacing="1" class="tabella_madre">
		<caption class="border"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Rimozione cartella", "Removing folder", "", "", "", "", "", "")%></caption>
			<% If Request.ServerVariables("REQUEST_METHOD") = "POST" and errore = "" then
				'cancellazione cartella OK
				%>
				<tr>
					<td class="content_b">
						
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "Cartella " & request("FOLDER") & " eliminata correttamente.", "Folder " & request("FOLDER") & " correctly removed.", "", "", "", "", "", "")%>
					</td>
				</tr>
				<tr>
					<td class="note">
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "Questa finestra si chiuder&agrave; automaticamente tra 5 secondi.", "This window is going to close automatically in 5 seconds.", "", "", "", "", "", "")%>
						<script language="JavaScript">
							opener.location.reload(true);
							window.setTimeout("close();", 5000);
						</script>
					</td>
				</tr>
				<tr>
					<td class="footer" colspan="2">
						<input type="button" class="button" name="annulla" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "CHIUDI", "CLOSE", "", "", "", "", "", "")%>" onclick="window.close();">
					</td>
				</tr>
			<% else 
				if errore<>"" then
					'creazione non andata a buon fine
					%>
					<tr>
						<td class="errore" colspan="2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Cancellazione non eseguita: ", "Cancelation not complete: ", "", "", "", "", "", "")%><%= errore %></td>
					</tr>
				<% end if %>
				<tr>
					<td class="label" rowspan="2" style="width:20%;"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Nome cartella", "Folder name", "", "", "", "", "", "")%></td>
					<td class="content_b">
						<%= request("FOLDER") %>
					</td>
				</tr>
				<tr>
					<td class="note">
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "Cartella contenuta in ", "Containing folder: ", "", "", "", "", "", "")%>"<%= request("FMPATH") %>".
					</td>
				</tr>
				<% if d.ExistFolder(path) then
					if d.IsEmptyFolder(path) then %>
						<tr>
							<td class="footer" colspan="2">
								<input type="submit" class="button" name="crea" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "CANCELLA CARTELLA", "DELETE FOLDER", "", "", "", "", "", "")%>">
								<input type="button" class="button" name="annulla" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "ANNULLA", "CANCEL", "", "", "", "", "", "")%>" onclick="window.close();">
							</td>
						</tr>
					<% else %>
						<tr>
							<td class="errore" colspan="2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Cancellazione non posibile: la directory non &egrave; vuota.", "Cancelation not allowed: The folder is not empty.", "", "", "", "", "", "")%></td>
						</tr>
						<tr>
							<td class="footer" colspan="2">
								<input type="button" class="button" name="annulla" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "CHIUDI", "CLOSE", "", "", "", "", "", "")%>" onclick="window.close();">
							</td>
						</tr>
					<% end if
				else %>
					<tr>
						<td class="errore" colspan="2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Cancellazione non posibile: la directory non esiste.", "Cancelation not allowed: The folder does not exists.", "", "", "", "", "", "")%></td>
					</tr>
					<tr>
						<td class="footer" colspan="2">
							<input type="button" class="button" name="annulla" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "CHIUDI", "CLOSE", "", "", "", "", "", "")%>" onclick="window.close();">
						</td>
					</tr>
				<% end if
			end if %>
		</table>
	</form>
</div>
</body>
</html>