<%
'.................................................................................................
'COSTANTI USATE NELL'EXPORT DATI
'.................................................................................................
'costanti per la definizione delle famiglie di tipi di dato
const TYPE_LIST_INTEGER = 	" 2 3 16 17 18 19 20 21 "
const TYPE_LIST_REAL = 		" 4 5 6 14 131 139 "
const TYPE_LIST_TEXT = 		" 8 129 130 200 201 202 203 "
const TYPE_LIST_DATE = 		" 7 64 133 134 135 "
const TYPE_LIST_BOOLEAN = 	" 11 "

'costanti per la definizione dei formati di export dati
const FORMAT_HTML 			= "HTML"
const FORMAT_XML 			= "XML"
const FORMAT_EXCEL_FILE 	= "EXCEL_2000"
const FORMAT_EXCEL_XML 		= "EXCEL_XP"
const FORMAT_TXT 			= "TXT"
const FORMAT_ACCESS 		= "ACCESS"

'.................................................................................................
'FUNZIONI PER L'EXPORT DATI
'.................................................................................................

'...............................................................................................................
'funzione che ritorna la famiglia del tipo di dato richiesto
'  typ 			codice ado del tipo di dato
'...............................................................................................................
function TypeFamily(TypeCode)
	TypeCode = " " & TypeCode & " "
	'cerca il tipo nelle costanti preparate
	if instr(1, TYPE_LIST_INTEGER, TypeCode, vbTextCompare)>0 then
		'numerico intero
		TypeFamily = adInteger
	elseif instr(1, TYPE_LIST_REAL, TypeCode, vbTextCompare)>0 then
		'numerico con la virgola
		TypeFamily = adNumeric
	elseif instr(1, TYPE_LIST_TEXT, TypeCode, vbTextCompare)>0 then
		'testo 
		TypeFamily = adChar
	elseif instr(1, TYPE_LIST_DATE, TypeCode, vbTextCompare)>0 then
		'data o time
		TypeFamily = adDate
	elseif instr(1, TYPE_LIST_BOOLEAN, TypeCode, vbTextCompare)>0 then
		'booleano
		TypeFamily = adBoolean
	else
		'tipo non riconosciuto
		TypeFamily = O
	end if
end function


'...............................................................................................................
'funzione che scrive il link per l'export dati o per l'apertura della palette di selezione del formato di export
'		LinkLabel				testo del link
'		ConnString				nome della connessione al database da cui prendere i dati
'		SessionQueryName		nome della variabile di sessione che contiene la query da esportare
'		Format					formato di export dei dati o formati abilitati:
'									se vuoto apre la palette di selezione del formato
'									se contiene un solo tipo: esegue direttamente l'export nel formato richiesto
'									se contiene una lista separata da "," o ";" apre la palette di selezione solo per i formati elencati
'								i valori sono dichiarati secondo le costati in testa a questo file
'		library_path_offset		displacement della directory library dalla directory in cui viene richiamata la 
'								funzione. Es: "../" per tutte le directory dopo la radice
'...............................................................................................................
sub WRITE_EXPORT_LINK(LinkLabel, ConnString, SessionQueryName, Format, CloseOnClick)
	CALL WRITE_EXPORT_LINK_ADV(LinkLabel, ConnString, SessionQueryName, Format, CloseOnClick, GetLibraryPath())
end sub

'...............................................................................................................
'funzione che scrive il link per l'export dati o per l'apertura della palette di selezione del formato di export
'		LinkLabel				testo del link
'		ConnString				nome della connessione al database da cui prendere i dati
'		SessionQueryName		nome della variabile di sessione che contiene la query da esportare
'		Format					formato di export dei dati o formati abilitati:
'									se vuoto apre la palette di selezione del formato
'									se contiene un solo tipo: esegue direttamente l'export nel formato richiesto
'									se contiene una lista separata da "," o ";" apre la palette di selezione solo per i formati elencati
'								i valori sono dichiarati secondo le costati in testa a questo file
'		library_path_offset		displacement della directory library dalla directory in cui viene richiamata la 
'								funzione. Es: "../" per tutte le directory dopo la radice
'		baseFilePath			url di base da usare per richiamare gli script di export
'...............................................................................................................
sub WRITE_EXPORT_LINK_ADV(LinkLabel, ConnString, SessionQueryName, Format, CloseOnClick, baseFilePath)
	dim title, link
	
	link = baseFilePath		'GetLibraryPath() Modificato da Nicola il 22/01/2015
	
	if format = "" OR instr(format, ",") OR instr(format, ";") then
		'apre la palette di selezione del formato di export dati
		title = "Apre la palette di selezione del formato di export"
		link = link + "ExportFormatSelection.asp" + _
			   "?conn=" + Server.UrlEncode(ConnString) + _
			   "&query=" + Server.UrlEncode(SessionQueryName) + _
			   "&format=" + format%>
		<a style="width:94%; display:block; text-align:center; line-height:12px;" class="button"
   	   	   title="<%= title %>" <%= ACTIVE_STATUS %> <%=IIF(CloseOnClick, " onClick=""window.close()""", "")%>
   	   	   onclick="OpenAutoPositionedWindow('<%= link %>', 'export', 400, 150);" href="javascript:void(0);">
		   <%= LinkLabel %>
		</a>
	<%else 
		'apre direttamente la funzione di export dati
		title = "Apre una nuova finestra contenente i dati "
		Select Case format
			case FORMAT_HTML
				title = title + " formattati in una tabella HTML"
			case FORMAT_XML
				title = title + " in formato MS-XML "
			case FORMAT_EXCEL_FILE
				title = title + " in un file MS Excel 97/2000"
			case FORMAT_EXCEL_XML
				title = title + " in un file MS Excel XP/2003"
			case FORMAT_TXT
				title = title + " in un file di testo con campi divisi da &quot;;&quot;"
		end select
		link = link + "ExportQuery.asp" + _
			   "?conn=" + Server.UrlEncode(ConnString) + _
			   "&query=" + Server.UrlEncode(SessionQueryName) + _
			   "&format=" + Format	   
		%>
		<a style="width:94%; display:block; text-align:center; line-height:12px;" class="button"
   	   	   title="<%= title %>" <%= ACTIVE_STATUS %> <%=IIF(CloseOnClick, " onClick=""window.close()""", "")%>
   	   	   href="<%= link %>" target="export<%=Format%>">
		   <%= LinkLabel %>
		</a>
	<% end if
end sub


'...............................................................................................................
'funzione che scrive il link per l'export dati o per l'apertura della palette di selezione del formato di export
'		LinkLabel				testo del link
'		ConnString				nome della connessione al database da cui prendere i dati
'		SessionQueryName		nome della variabile di sessione che contiene la query da esportare
'		Format					formato di export dei dati o formati abilitati:
'									se vuoto apre la palette di selezione del formato
'									se contiene un solo tipo: esegue direttamente l'export nel formato richiesto
'									se contiene una lista separata da "," o ";" apre la palette di selezione solo per i formati elencati
'								i valori sono dichiarati secondo le costati in testa a questo file
'		library_path_offset		displacement della directory library dalla directory in cui viene richiamata la 
'								funzione. Es: "../" per tutte le directory dopo la radice
'...............................................................................................................
sub WRITE_CONTATTI_EXPORT_LINK(LinkLabel, SessionQueryName, Format, CloseOnClick)
	dim title, link
	link = GetAmministrazionePath()
	
	'apre la finestra per l'export dati
	title = "Apri export dati"
	link = link + "Nextcom/ContattiExport.asp" + _
		   "?sql=" + Server.UrlEncode(SessionQueryName)
	if cString(format)<>"" then
		link = link + "&export=" + format + "&esporta=1"
	end if%>
	<a style="width:94%; display:block; text-align:center; line-height:12px;" class="button"
	   title="<%= title %>" <%= ACTIVE_STATUS %> <%=IIF(CloseOnClick, " onClick=""window.close()""", "")%>
	   onclick="OpenAutoPositionedWindow('<%= link %>', 'export', 400, 150);" href="javascript:void(0);">
	   <%= LinkLabel %>
	</a>
	<%
end sub


'...............................................................................................................
'procedura che esporta in formato HTML i dati contenuti nel recordset
'	rs			recordset di cui esportare  i dati
'...............................................................................................................
sub ExportRecordset_HTML(rs)
	dim Field, i
	response.ContentType = "text/html"%>
	<html>
		<head>
		<title>Export dati in formato HTML</title>
		<meta name="robots" content="noindex,nofollow" />
		<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
		</head>
		<body>
		<table border="1">
			<tr>
				<%for each Field in rs.Fields%>
					<th><%= Field.name %></th>
				<%next%>
			</tr>
			<% 	response.Flush
				i = 0
				while not rs.eof
					i = i+1
					if i mod 500 = 0 then
						response.Flush
					end if %>
				<tr>
					<%for each Field in rs.Fields%>
						<td><%= IIF(cString(Field.value)="", "&nbsp;", Field.value) %></td>
					<%next%>
				</tr>
				<% rs.MoveNext
			wend %>
		</table>
		</body>	
	</html>
<%end sub


'...............................................................................................................
'procedura che esporta in formato XML i dati contenuti nel recordset
'	rs			recordset di cui esportare  i dati
'...............................................................................................................
sub ExportRecordset_XML(rs)
	CALL Export_XML(rs, false)
end sub


'...............................................................................................................
'trasforma i caratteri accentati in carattere piu accento o viceversa
'	txt			testo da ripulire
'	inAccento	se true converte da accentati a "char'" else il contrario
'...............................................................................................................
Function ConvertiAccentati(txt, inAccento)
	if inAccento then
		ConvertiAccentati = Server.HtmlEncode(CString(txt))
	else
		ConvertiAccentati = Server.HtmlDecode(CString(txt))
	end if
End Function


'...............................................................................................................
'procedura che esporta in formato EXCEL per excel XP o superiore i dati contenuti nel recordset
'	rs			recordset di cui esportare  i dati
'...............................................................................................................
sub ExportRecordset_EXCEL_XML(rs)
	dim field, ExportPath, FileName, FSO, xml, val
	
	ExportPath = Application("IMAGE_PATH") & "temp\"
	FileName = Session.SessionID & ".xls"
	Set FSO = CreateObject("Scripting.FileSystemObject")
	set xml = fso.CreateTextFile(ExportPath + FileName, True)
	with xml
	.WriteLine	"<?xml version=""1.0"" encoding=""UTF-8""?>"& _
				"<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"">"& vbCrLf & _
 				"	<DocumentProperties>"& vbCrLf & _
  				"		<Author>"& Session("USER_4_LOG") &"</Author>"& vbCrLf & _
	  			"		<Created>"& DateIso(Date) &"T"& Hour(date) &":"& Minute(date) &":00Z</Created>"& vbCrLf & _
  				"		<LastSaved>"& DateIso(Date) &"T"& Hour(date) &":"& Minute(date) &":00Z</LastSaved>"& vbCrLf & _
 				"	</DocumentProperties>"& vbCrLf & _
 				"	<Worksheet ss:Name=""Dati esportati"">"& vbCrLf & _
	  			"		<Table ss:DefaultColumnWidth=""100"">"& vbCrLf & _
		   		"			<Row>"
	for each field in rs.fields
		.WriteLine	"				<Cell><Data ss:Type=""String"">"& ConvertiAccentati(field.name, true) &"</Data></Cell>"& vbCrLf
	next
	.WriteLine	"			</Row>"& vbCrLf
	while not rs.eof
		.WriteLine	"			<Row>"& vbCrLf
		for each field in rs.fields
			val = ConvertiAccentati(field.value, true)
   			select case TypeFamily(field.type)
 				case adInteger
			response.write "integer<br>"
					.WriteLine	"				<Cell><Data ss:Type=""Number"">"& val &"</Data></Cell>"& vbCrLf
				case adNumeric
			response.write "numeric<br>"
					.WriteLine	"				<Cell><Data ss:Type=""Number"">"& val &"</Data></Cell>"& vbCrLf
  				case adBoolean
					if rs(field.name) then
						.WriteLine	"				<Cell><Data ss:Type=""String"">Si</Data></Cell>"& vbCrLf
					else
						.WriteLine	"				<Cell><Data ss:Type=""String""></Data></Cell>"& vbCrLf
					end if
				case adDate
					if isDate(rs(field.name)) then
						.WriteLine	"				<Cell><Data ss:Type=""DateTime"">"& DateIso(rs(field.name)) &"T00:00:00.000</Data></Cell>"& vbCrLf
					else
						.WriteLine	"				<Cell><Data ss:Type=""String""></Data></Cell>"& vbCrLf
					end if
				case adChar
					.WriteLine	"				<Cell><Data ss:Type=""String"">"& Trim("" & val) &"</Data></Cell>"& vbCrLf
				case else
					.WriteLine	"				<Cell><Data ss:Type=""String"">###</Data></Cell>"& vbCrLf
			end select
		next
		.WriteLine	"			</Row>"& vbCrLf
		rs.movenext
	wend
	.WriteLine	"		</Table>"& vbCrLf & _
 				"	</Worksheet>"& vbCrLf & _
				"</Workbook>"& vbCrLf
	end with
	
	xml.Close
	response.redirect "http://" & Application("IMAGE_SERVER") & "/temp/" & FileName
end sub


'...............................................................................................................
'procedura che esporta in un file EXCEL 97/2000 e poi esegue il redirect al file
'	rs			recordset di cui esportare  i dati
'...............................................................................................................
sub ExportRecordset_Access(rs)
	dim FileName, FSO, ExportPath, XLSConnString, XLSConn, XLSrs, Xsql, field, FileStream, fieldName, fields, i, aux
	set fields = server.CreateObject("Scripting.Dictionary")
	
	ExportPath = Application("IMAGE_PATH") & "temp\"
	FileName = Session.SessionID & ".mdb"
	
	'crea il file access
	Set FSO = CreateObject("Scripting.FileSystemObject")
	FSO.CopyFile Server.Mappath("../library") & "/FilesTemplate/Master.mdb", ExportPath + FileName , true
	
	'apre connessione a file ACCESS
	set XLSConn = Server.CreateObject("ADODB.Connection")
	XLSConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (ExportPath & FileName)
	XLSConn.open XLSConnString, "", ""
	
	'generazione sql per creazione tabella di export
	Xsql = "CREATE TABLE [Dati esportati] ( "
	for each field in rs.fields
		
		'eliminazione caratteri non validi
		fieldName = Replace(field.name, ".", "_")
		fieldName = Sintesi(fieldName, 57, "___")
		
		'controllo presenza di un campo con nome uguale
		i = 0
		aux = fieldName
		while InStr(1, Xsql, "["& aux &"]", vbTextCompare) > 0
			aux = fieldName &"_"& i
			i = i+1
		wend
		if i > 0 then
			fieldName = aux
		end if
		fields.Add fieldName, field.name
		
		if field.type = adBigInt OR _
		   field.type = adInteger OR _
		   field.type = adSmallInt OR _
		   field.type = adTinyInt OR _
		   field.type = adUnsignedBigInt OR _
		   field.type = adUnsignedInt OR _
		   field.type = adUnsignedSmallInt OR _
		   field.type = adUnsignedTinyInt then
		   	'numerico intero
			Xsql = Xsql & "[" & fieldName & "] int, "
		elseif field.type = adBoolean then
			'valore booleano
			Xsql = Xsql & "[" & fieldName & "] bit, "
		elseif field.type = adDecimal OR _
			   field.type = adDouble OR _
			   field.type = adNumeric OR _
			   field.type = adSingle OR _
			   field.type = adVarNumeric then
			'valore reale
			Xsql = Xsql & "[" & fieldName & "] real, "
		elseif field.type = adCurrency then
			'valore "moneta"
			Xsql = Xsql & "[" & fieldName & "] Currency, "
		elseif field.type = adDate OR _
		  	   field.type = adDBDate OR _
		  	   field.type = adDBTime OR _
		  	   field.type = adDBTimeStamp OR _
		  	   field.type = adFileTime then
			'valore data
			Xsql = Xsql & "[" & fieldName & "] Date, "
		elseif field.type = adChar OR _
		  	   field.type = adLongVarChar OR _
		  	   field.type = adLongVarWChar OR _
		  	   field.type = adVarChar OR _
		  	   field.type = adVarWChar OR _
		  	   field.type = adWChar OR _
		  	   field.type = adDBDate then
			'valore testuale
			if field.definedSize > 250 then
				Xsql = Xsql & "[" & fieldName & "] text, "
			else
				Xsql = Xsql & "[" & fieldName & "] varchar(" & Field.DefinedSize & "), "
			end if
		end if
	next
	Xsql = left(Xsql, len(Xsql)-2) + ")"
	
	'creazione tabella
	CALL XLSConn.execute(Xsql, 0, adExecuteNoRecords)
	
	'apre recordset su tabella
	set XLSrs = Server.CreateObject("ADODB.RecordSet")
	XLSrs.open "[Dati esportati]", XLSConn, adOpenStatic, adLockOptimistic, adCmdTable
	
	'riempie tabella
	while not rs.eof
		XLSrs.AddNew
		for each field in XLSrs.fields
			if NOT IsNull(rs(fields(field.name))) then
				XLSrs(field.name) = rs(fields(field.name))
			end if
		next
		rs.movenext
		XLSrs.Update
	wend
	
	XLSrs.close
	XLSConn.close
	set XLSrs = nothing
	set XLSConn = nothing
	
	'reindirizza al file
	response.redirect "http://" & Application("IMAGE_SERVER") & "/temp/" & FileName
end sub


'...............................................................................................................
'procedura che esporta in un file EXCEL 97/2000 e poi esegue il redirect al file
'	rs			recordset di cui esportare  i dati
'...............................................................................................................
sub Export_Excel2000(rs)
	dim FileName, FSO, ExportPath, XLSConnString, XLSConn, XLSrs, Xsql, field, FileStream, fieldName, fields, i, aux
	set fields = server.CreateObject("Scripting.Dictionary")
	
	ExportPath = Application("IMAGE_PATH") & "temp\"
	FileName = Session.SessionID & ".XLS"
	
	'crea il file excel
	Set FSO = CreateObject("Scripting.FileSystemObject")
	if FSO.FileExists(ExportPath + FileName) then
		FSO.DeleteFile (ExportPath + FileName), true
	end if
	FSO.CopyFile Server.Mappath("..\library") & "\FilesTemplate\Master.xls", ExportPath + FileName, true
	
	'apre connessione a file EXCEL
	set XLSConn = Server.CreateObject("ADODB.Connection")
	XLSConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (ExportPath & FileName) &_
					";Extended Properties=""Excel 8.0;HDR=YES;"";Jet OLEDB:Engine Type=35;"
	XLSConn.open XLSConnString, "", ""
	
	'generazione sql per creazione tabella di export
	Xsql = "CREATE TABLE [Dati esportati] ( "
	for each field in rs.fields
		
		'eliminazione caratteri non validi
		fieldName = Replace(field.name, ".", "_")
		fieldName = Sintesi(fieldName, 57, "___")
		
		'controllo presenza di un campo con nome uguale
		i = 0
		aux = fieldName
		while InStr(1, Xsql, "["& aux &"]", vbTextCompare) > 0
			aux = fieldName &"_"& i
			i = i+1
		wend
		if i > 0 then
			fieldName = aux
		end if
		fields.Add fieldName, field.name
		
		if field.type = adBigInt OR _
		   field.type = adInteger OR _
		   field.type = adSmallInt OR _
		   field.type = adTinyInt OR _
		   field.type = adUnsignedBigInt OR _
		   field.type = adUnsignedInt OR _
		   field.type = adUnsignedSmallInt OR _
		   field.type = adUnsignedTinyInt then
		   	'numerico intero
			Xsql = Xsql & "[" & fieldName & "] int, "
		elseif field.type = adBoolean then
			'valore booleano
			Xsql = Xsql & "[" & fieldName & "] bit, "
		elseif field.type = adDecimal OR _
			   field.type = adDouble OR _
			   field.type = adNumeric OR _
			   field.type = adSingle OR _
			   field.type = adVarNumeric then
			'valore reale
			Xsql = Xsql & "[" & fieldName & "] real, "
		elseif field.type = adCurrency then
			'valore "moneta"
			Xsql = Xsql & "[" & fieldName & "] Currency, "
		elseif field.type = adDate OR _
		  	   field.type = adDBDate OR _
		  	   field.type = adDBTime OR _
		  	   field.type = adDBTimeStamp OR _
		  	   field.type = adFileTime then
			'valore data
			Xsql = Xsql & "[" & fieldName & "] Date, "
		elseif field.type = adChar OR _
		  	   field.type = adLongVarChar OR _
		  	   field.type = adLongVarWChar OR _
		  	   field.type = adVarChar OR _
		  	   field.type = adVarWChar OR _
		  	   field.type = adWChar OR _
		  	   field.type = adDBDate then
			'valore testuale
			if field.definedSize > 250 then
				Xsql = Xsql & "[" & fieldName & "] text, "
			else
				Xsql = Xsql & "[" & fieldName & "] varchar(" & Field.DefinedSize & "), "
			end if
		end if
	next
	Xsql = left(Xsql, len(Xsql)-2) + ")"
	
	'creazione tabella
	CALL XLSConn.execute(Xsql, 0, adExecuteNoRecords)
	
	'apre recordset su tabella
	set XLSrs = Server.CreateObject("ADODB.RecordSet")
	XLSrs.open "[Dati esportati]", XLSConn, adOpenStatic, adLockOptimistic, adCmdTable
	
	'riempie tabella
	while not rs.eof
		XLSrs.AddNew
		for each field in XLSrs.fields
			if NOT IsNull(rs(fields(field.name))) then
				XLSrs(field.name) = rs(fields(field.name))
				'XLSrs(field.name) = replace(rs(fields(field.name)),chr(13),"")
			end if
		next
		rs.movenext
		XLSrs.Update
	wend
	
	XLSrs.close
	XLSConn.close
	set XLSrs = nothing
	set XLSConn = nothing
	
	'reindirizza al file
	response.redirect "http://" & Application("IMAGE_SERVER") & "/temp/" & FileName
	
	'carica file su stream e lo restituisce
	
'	Set FileStream = Server.CreateObject("ADODB.Stream")
'	FileStream.Open
'	FileStream.Type = 1
'	FileStream.LoadFromFile(ExportPath & FileName)

'	Response.Clear
'	response.contentType="application/vnd.ms-excel"
'	Response.BinaryWrite(FileStream.Read)

'	FileStream.Close
'	Set FileStream = Nothing

'	cancella file XLS creato
'	FSO.DeleteFile (ExportPath & FileName) , true
'	set FSO = nothing
	
end sub


'...............................................................................................................
'procedura che esporta in formato TXT separati da ; i dati contenuti nel recordset
'	rs			recordset di cui esportare  i dati
'...............................................................................................................
sub ExportRecordset_TXT(rs)
	dim Field
	response.ContentType = "text/plain"
	
	'scrive testata
	for each Field in rs.Fields
		response.write Field.name & ";"
	next 
	response.write vbCrLf
	response.write rs.GetString(adClipString, , ";", vbCrLf, "")
end sub



'.................................................................................................
'CLASSI PER LA GESTIONE DEI DATI
'.................................................................................................

'.................................................................................................
'classe per il calcolo dei totali numerici
'.................................................................................................
'un esempio di utilizzo della classe si trova in 
'D:\frameworks\turismovenezia.it.2004\web\APTDistibuzione\ProdottiStatistiche.asp
'.................................................................................................
class CalcolaTotali
	public values
	
	Private Sub Class_Initialize()
		'crea oggetto contenente i totali
		set values = Server.CreateObject("Scripting.Dictionary")
		values.CompareMode = vbTextCompare
		
	end sub
	
	Private Sub Class_Terminate()
		set values = nothing
	End Sub
	
	'verifica l'esistenza del totale richiesto
	Public function Exists(field)
		Exists = values.Exists(field)
	end function
	
	Public Sub Reset()
		values.RemoveAll
	end sub
	
	'restituisce il valore dell'elemento richiesto
	Public Default Property Get totale(ByVal field)
		totale = values(field)
	end Property
		
	'imposta il valore dell'elemento richiesto
	Public Property Let totale(ByVal Key, ByVal Value)
		if not values.Exists(Key) then
			values.Add Key, cReal(value)
		else
			values(Key) = values(Key) + cReal(value)
		end if	
	end Property
	
end class
%>