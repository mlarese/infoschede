<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="class_directory.asp" -->
<!--#INCLUDE FILE="../classUpload.asp" -->
<%'--------------------------------------------------------
sezione_testata = "Cancellazione file" %>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim d, fileType
set d = new directory
d.RelativePath = Session("fmpath")


'detaermina tipo di file
if instr(1, Session("FMPATH"), FILE_TYPE_IMAGE, vbtextCompare)>0 then
	fileType = FILE_TYPE_IMAGE
elseif  instr(1, Session("FMPATH"), FILE_TYPE_FLASH, vbtextCompare)>0 then
	fileType = FILE_TYPE_FLASH
elseif  instr(1, Session("FMPATH"), FILE_TYPE_OBJECTS, vbtextCompare)>0 then
	fileType = FILE_TYPE_OBJECTS
elseif  instr(1, Session("FMPATH"), FILE_TYPE_TEXT, vbtextCompare)>0 then
	fileType = FILE_TYPE_TEXT
elseif  instr(1, Session("FMPATH"), FILE_TYPE_XML, vbtextCompare)>0 then
	fileType = FILE_TYPE_XML
elseif  instr(1, Session("FMPATH"), FILE_TYPE_CSS, vbtextCompare)>0 then
	fileType = FILE_TYPE_CSS
end if



%>

<div id="content_ridotto">
	<table cellspadding="0" cellspacing="1" class="tabella_madre">
		<caption class="border">Cancellazione file</caption>
		<tr>
			<td class="note" colspan="2">
				File contenuto in "<%= Session("FMPATH") %>".
			</td>
		</tr>
		<% 
		
		if FileCanBeRemoved(NULL, NULL, NULL, fileType, Session("FILEMAN_AZ_ID"), d.RelativeURL(request("FILE"))) then %>
			<tr>
				<td class="content">
					<%dim DEL_OBJ
					SET DEL_OBJ = New DeleteFile
					DEL_OBJ.File_Path = d.FILEPath("")
					
					DEL_OBJ.File_Name = request.Querystring("FILE")	
					DEL_OBJ.Stile_Input = ""	
					DEL_OBJ.Stile_Submit = "class=""button"""
					DEL_OBJ.Stile_Titoli = "class=""content_b"""
					DEL_OBJ.Stile_testo	= "class=""content"""
					DEL_OBJ.Update_Record = ""
					
					DEL_OBJ.Delete()%>
				</td>
			</tr>
			<%if DEL_OBJ.OperationOK then%>
				<tr>
					<td class="note">
						Questa finestra si chiuder&agrave; automaticamente tra 5 secondi.
						<script language="JavaScript">
							opener.location.reload(true);
							window.setTimeout("close();", 5000);
						</script>
					</td>
				</tr>
			<%end if
		else %>
			<tr>
				<td class="errore">
					File non cancellabile perch&egrave; utilizzato nella costruzione delle pagine o dal sistema.
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="footer" colspan="2">
				<input type="button" class="button" name="annulla" value="CHIUDI" onclick="window.close();">
			</td>
		</tr>
	</table>
</div>
</body>
</html>