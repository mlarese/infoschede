<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="intestazione.asp" -->
<% 
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_strumenti_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Analisi utilizzo immagini"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoAnalisi.asp"
dicitura.scrivi_con_sottosez()  

dim ImagePath, ImageURL
dim conn, sql, rs, List, i, FileName
ImagePath = Application("IMAGE_PATH") & Session("AZ_ID") & "\images"  '"
ImageURL = "http://" & Application("IMAGE_SERVER") & "/" & Session("AZ_ID") & "/images/"

'inizializza lista files
List = array()
CALL GetFullFileList(List, ImagePath, "")

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Analisi immagini</caption>
		<tr>
			<th>NOME</th>
			<th class="center">UTILIZZATA</th>
			<th class="center" colspan="3" style="width:32%;">OPERAZIONI</th>
		</tr>
				<% for i=lbound(List) to UBound(List)
					if Ucase(List(i)) <> "VUOTO.JPG" then
						sql = "SELECT (COUNT(*)) AS DIPENDENZE FROM tb_layers INNER JOIN tb_pages ON tb_layers.id_pag=tb_pages.id_page "  &_
							  " WHERE tb_pages.id_webs=" & Session("AZ_ID") & " AND id_tipo=" & LAYER_IMAGE &_
							  " AND (tb_layers.nome LIKE '" & PArseSql(List(i), adChar) & "' OR tb_layers.nome LIKE '" & ParseSql(right(List(i), len(list(i))-1), adChar) & "')"
						rs.open sql, conn, adOpenstatic, adLockReadOnly, adAsyncFetch
						if Left(List(i), 1) = "/" then
							FileName = replace(List(i), "/", "", 1, 1)		'toglie prima barra se presente
						else
							FileName = List(i)
						end if %>
						<tr>
							<td class="content">
								<a HREF="javascript:void(0);" onclick="OpenAutoPositionedWindow('<%= JsEncode(ImageURL & FileName, "'") %>', 'image', 420, 210)"
									title="Visualizza l'immagine in una nuova finestra">
									<%= FileName %>
								</a>
							</td>
							<td class="content_center"><input type="checkbox" class="checkbox" disabled <% if rs("DIPENDENZE")>0 then %> checked <% end if %>></td>
							<td class="content_center">
								<a HREF="javascript:void(0);" onclick="OpenAutoPositionedWindow('<%= JsEncode(ImageURL & FileName, "'") %>', 'image', 420, 210)"
									title="Visualizza l'immagine in una nuova finestra" class="button">
									VISUALIZZA
								</a>
							</td>
							<% if rs("DIPENDENZE")>0 then 
								'immagine utilizzata%>
								<td class="content_center">
									<a href="javascript:void(0);" onclick="OpenAutoPositionedWindow('SitoAnalisiDipendenze.asp?TYPE=images&FILE=<%= JsEncode(List(i), "'") %>', 'dipendenze', 420, 410)"
										title="Visualizza tutte le pagine in cui l'immagine &egrave; utilizzata" class="button">
										DIPENDENZE
									</a>
								</td>
								<td class="content_center">
									<a title="Impossibile cancellare l'immagine perch&egrave; utilizzata" class="button_disabled">
										CANCELLA
									</a>
								</td>
							<% else 
								'immagine non utilizzata%>
								<td class="content_center">
									<a title="nessuna dipendenza: immagine non utilizzata" class="button_disabled">
										DIPENDENZE
									</a>
								</td>
								<td class="content_center">
									<a HREF="javascript:void(0);" onclick="OpenAutoPositionedWindow('../library/filemanager/FileRemove.asp?FILE=<%= server.urlEncode(FileName) %>', 'delete', 420, 210)"
										title="cancella l'immagine" class="button">
										CANCELLA
									</a>
								</td>
							<% end if %>
						</tr>
						<% rs.close
					end if
				next %>
	</table>
</div>
</body>
</html>
<%
conn.close
set rs = nothing
set conn = nothing 


'...................................................................................................
'funzione che restituisce un array contenente tutti i file, compresi quelli delle sottocartelle
'		List: 			Array contenente i file
'		BasePath		Path fisico di partenza dell'analisi
'		RelativePath	Path di partenza restituito nel nome del file
'...................................................................................................
sub GetFullFileList(Byref List, BasePath, RelativePath)
	dim fso, BaseDir, Dir, File, FileName, i
	Set fso = CreateObject("Scripting.FileSystemObject")

	'esegue ricorsione per sottodirectory
	Set BaseDir = fsO.GetFolder(BasePath)
	For each Dir in BaseDir.SubFolders
		CALL GetFullFileList(List, Dir.Path, RelativePath & "/" & Dir.Name)
	next

	'general lista file per directory corrente
	For each File in BaseDir.Files
		ReDim Preserve List(Ubound(list) + 1)
		List(ubound(list)) = IIF(RelativePath<>"", RelativePath & "/", "")  & File.name
	next
end sub
%>
