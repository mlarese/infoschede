<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/ClassUpload.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<%'--------------------------------------------------------
sezione_testata = "Gestione allegati" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>


<script language="JavaScript" type="text/javascript">
	//aggiunge file appena caricato a lista visibile
	function AllegaFile(field, file){
		var file_list = field.value;
		file_list = file_list.toUpperCase( );
		
		var Ufile = file.toUpperCase( );
		
		if (file_list.indexOf(Ufile)<0){
			field.value += file;
			}
	}
	
</script>
<%
if Request.querystring("MODE")="ALLEGA" then
	'caricamento di un nuovo file
	dim UPL_OBJ
	SET UPL_OBJ = New UploadFile
	UPL_OBJ.stile_titoli = " style=""color:#FFF;"""							
	UPL_OBJ.File_Path = Application("IMAGE_PATH") & "\temp\"
	UPL_OBJ.FileUpload()
	if UPL_OBJ.OperationOK then%>
		<script language="JavaScript" type="text/javascript">
			AllegaFile(opener.document.form1.tft_email_docs, "  <%= UPL_OBJ.File_Name %>;");
		</script>
	<%else
		Session("ERRORE") = "Errore nel caricamento del file."
	end if	
elseif Request.querystring("MODE")="RIMUOVI" then
	'cancellazione file
	dim DEL_OBJ
	SET DEL_OBJ = New DeleteFile
	DEL_OBJ.stile_titoli = " style=""color:#FFF;"""		
	DEL_OBJ.File_Path = Application("IMAGE_PATH") & "\temp\"
	DEL_OBJ.File_Name = request.Querystring("FILE")	
	DEL_OBJ.FileDelete()
	if DEL_OBJ.OperationOK then%>
		<script language="JavaScript" type="text/javascript">
			//toglie file da lista visibile
			var file_list = opener.document.form1.tft_email_docs.value;
			var start = file_list.indexOf("  <%= DEL_OBJ.File_Name %>;");
			end = start + <%= len("  " & DEL_OBJ.File_Name & ";")%>;
			
			file_list = file_list.substr(0, start) + file_list.substr(end);
			opener.document.form1.tft_email_docs.value = file_list;
		</script>
	<%else
		Session("ERRORE") = "Errore nella cancellazione del file."
	end if
end if

if Request.querystring("MODE")<>"" then%>
	<script language="JavaScript" type="text/javascript">
		document.location="ComunicazioniNew_Wizard_2_Allegati.asp?docs=" + opener.document.form1.tft_email_docs.value;
	</script>
<%end if%>
<div id="content_ridotto">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td>
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						<table width="100%" border="0" cellspacing="0">
							<tr>
								<% if Session("ERRORE")<>"" then %>
									<td class="errore"><%= Session("ERRORE") %></td>
									<% Session("ERRORE") = ""
								else %>
									<td class="caption">Files allegati</td>
								<% end if %>
							</tr>
						</table>
					</caption>
					<% if request.QueryString("docs")<>"" then %>
						<tr>
							<th width="84%">FILE</th>
							<th class="center">RIMUOVI</th>
						</tr>
					<% dim FileList, FileName, i
					FileList = split(request.querystring("docs"), ";")
					for i = lbound(FileList) to ubound(FileList)
						FileName = Trim(FileList(i))
						if FileName<>"" then%>
							<tr>
								<td class="content"><%= FileName %></td>
								<td class="content_center">
									<a class="button" href="?file=<%= FileName %>&docs=<%= Trim(request.querystring("docs")) %>&MODE=RIMUOVI">
										RIMUOVI
									</a>
								</td>
							</tr>
						<%end if
					next
					else %>
						<tr><td style="border-bottom:0px;" class="noRecords">Nessun file allegato.</th></tr>
					<% end if %>
				</table>
			</td>
		</tr>
		<tr><td style="font-size:5px;">&nbsp;</td></tr>
		<form action="?docs=<%= Trim(request.querystring("docs")) %>&MODE=ALLEGA" method="post" enctype="multipart/form-data" name="form1" id="form1">
		<tr>
			<td>
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>Allega file</caption>
					<tr>
						<th colspan="2">SELEZIONE DEL FILE</th>
					</tr>
					<tr>
						<td class="label" style="width:22%;" rowspan="2">file da allegare:</td>
						<td class="content"><input type="File" name="FILE1" value="" style="width:100%;"></td>
					</tr>
					<tr>
						<td class="note">
							ATTENZIONE: Se il file da allegare &egrave; di grandi dimensioni l'operazione potrebbe durare alcuni minuti.
						</td>
					</tr>
					<tr>
						<td class="footer" colspan="2">
							<input style="width:15%;" type="submit" class="button" name="allega" value="ALLEGA">
						</td>
					</tr>
				</table>
			</td>
		</tr>
		</form>
	</table>
	<script language="JavaScript" type="text/javascript">
		FitWindowSize(this);
	</script>
</body>
</html>


