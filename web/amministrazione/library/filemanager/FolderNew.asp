<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="class_directory.asp" -->
<!--#INCLUDE FILE="../classUpload.asp" -->
<%'--------------------------------------------------------
sezione_testata = ChooseValueByAllLanguages(Session("LINGUA"), "Creazione nuova cartella", "Creating new folder", "", "", "", "", "", "") %>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim d,path ,errore, nome_iniziale

If Request.ServerVariables("REQUEST_METHOD") = "POST" then
	path = request("FOLDER")
	if CheckChar(path, FOLDER_VALID_CHARSET) then
		set d = new directory
		d.RelativePath = request("fmpath")
		
		if path <> "" and instr(1, path, " ", vbTextCompare)<1 then
			if not d.ExistFolder(path) then
				d.CreateFolder(path)
			else
				errore = ChooseValueByAllLanguages(Session("LINGUA"), "La cartella esiste gi&agrave;", "Folder already exist", "", "", "", "", "", "")
			end if
		end if
	else
		errore = ChooseValueByAllLanguages(Session("LINGUA"), "Nome cartella non valido.", "Invalid folder name", "", "", "", "", "", "")
	end if
end if
%>

<div id="content_ridotto">
	<form method="POST" action="" id="form1" name="form1">
		<input type="hidden" name="fmpath" id="fmpath" value="<%= Session("FMPATH") %>">
		<table cellspadding="0" cellspacing="1" class="tabella_madre">
			<caption class="border"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Creazione nuova cartella", "Creating new folder", "", "", "", "", "", "")%></caption>
			<% If Request.ServerVariables("REQUEST_METHOD") = "POST" and errore = "" then
				'creazione cartella OK
				%>
				<tr>
					<td class="content_b">
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "Cartella " & request("FOLDER") & " creata correttamente.", "Folder " & request("FOLDER") & " corretly created.", "", "", "", "", "", "")%>
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
						<td class="errore" colspan="2"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Creazione non riuscita: " & errore , "Creation failed: " & errore, "", "", "", "", "", "")%></td>
					</tr>
				<% end if %>
				<%
				if Request.ServerVariables("REQUEST_METHOD") = "POST" and request("FOLDER") <> "" then
					nome_iniziale = request("FOLDER")
				else
					nome_iniziale = ChooseValueByAllLanguages(Session("LINGUA"), "Nuova cartella", "New folder", "", "", "", "", "", "")
				end if
				%>
				<tr>
					<td class="label" style="width:20%;"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Nome cartella", "Folder name", "", "", "", "", "", "")%></td>
					<td class="content">
						<input type="text" class="text" size="30" name="FOLDER" value="<%= nome_iniziale%>">
					</td>
				</tr>
				<tr>
					<td class="note" colspan="2">
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "La cartella sar&agrave; creata in ", "The folder will be created in the path " & Session("FMPATH") & ".", "", "", "", "", "", "")%>
						<br>
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "Il nome non pu&ograve; contenere spazi al suo interno.", "Folder name cannot contain blank spaces inside.", "", "", "", "", "", "")%>
					</td>
				</tr>
				<tr>
					<td class="footer" colspan="2">
						<input type="submit" class="button" name="crea" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "CREA CARTELLA", "CREATE FOLDER", "", "", "", "", "", "")%>">
						<input type="button" class="button" name="annulla" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "ANNULLA", "CANCEL", "", "", "", "", "", "")%>" onclick="window.close();">
					</td>
				</tr>
			<% end if %>
		</table>
	</form>
</div>
</body>
</html>
