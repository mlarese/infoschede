<%@ Language=VBScript CODEPAGE=65001%>
<% option explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 100000 %>
<!--#INCLUDE FILE="class_directory.asp" -->
<!--#INCLUDE FILE="../classUpload.asp" -->
<%'--------------------------------------------------------
sezione_testata = "Caricamento nuovo file" %>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

%>

<div id="content_ridotto">
	<table cellspadding="0" cellspacing="1" class="tabella_madre">
		<caption class="border">Caricamento nuovo file in " <%= Session("FMPATH") %> "</caption>
		<tr>
			<td class="content">
				<%dim UPL_OBJ,d
				set d = new directory
				d.RelativePath = Session("FMPATH")
				
				SET UPL_OBJ = New UploadFile
				UPL_OBJ.File_Path = d.FILEPath("")
				UPL_OBJ.ShowConsigli = false
				UPL_OBJ.OnlyExtensionAllowed = true
				UPL_OBJ.OverWrite = NULL
				UPL_OBJ.Stile_Submit = "class=""button"""
				UPL_OBJ.Stile_Titoli = "class=""content_b"""
				UPL_OBJ.Stile_testo	= "class=""content"""
				UPL_OBJ.Upload()%>
			</td>
		</tr>
		<%if UPL_OBJ.OperationOK then%>
			<tr>
				<td class="note">
					Il file &egrave; stato caricato in " <%= d.RelativePath %> ".
				</td>
			</tr>
			<tr>
				<td class="note">
					Questa finestra si chiuder&agrave; automaticamente tra 5 secondi.
					<script language="JavaScript">
						opener.location.reload(true);
						window.setTimeout("close();", 5000);
					</script>
				</td>
			</tr>
		<% else 
			if request.ServerVariables("REQUEST_METHOD")<>"POST" then%>
			<tr>
				<td class="note">
					Il file sar&agrave; caricato " <%= d.RelativePath %> ".<br>
					Il nome del file non pu&ograve; contenere spazi al suo interno, se si il file verr&agrave; comunque caricato, ma non verr&agrave; utilizzato dalle applicazioni.
				</td>
			</tr>
			<%end if
		end if%>
		<tr>
			<td class="footer" colspan="2">
				<% if request.ServerVariables("REQUEST_METHOD")="POST" AND not UPL_OBJ.OperationOK then %>
					<input type="button" class="button" name="indietro" value="INDIETRO" onclick="history.back();">
				<% end if %>
				<input type="button" class="button" name="annulla" value="CHIUDI" onclick="window.close();">
			</td>
		</tr>
	</table>
</div>
</body>
</html>
