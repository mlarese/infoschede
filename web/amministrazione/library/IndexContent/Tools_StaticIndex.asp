<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<% Server.ScriptTimeout = 2147483647 %>
<%
dim StaticFilePath, StaticCssFilePath, BaseCssPath, OriginalCssPath, StaticBaseCssPath
StaticFilePath = replace(Application("IMAGE_PATH") + "\static\", "\\", "\")
StaticCssFilePath = replace(Application("IMAGE_PATH") + "\static_css\", "\\", "\")
BaseCssPath = "\app_themes\default\"
StaticBaseCssPath = "\upload\static_css\"
OriginalCssPath = replace(replace(Application("IMAGE_PATH"), "upload", "web") + BaseCssPath, "\\", "\")


Sub Index_StaticizzaCancellaElementi(sql)
	dim conn
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open Application("DATA_ConnectionString")
	
	dim rs, htmls
	set rs = Server.CreateObject("ADODB.recordset")
	rs.CursorLocation = adUseClient
	
	CALL Index_SalvaCSS()
	
	%><!--<%=sql%>--><%
	
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	Set rs.ActiveConnection = Nothing
	%>
	<script language="JavaScript" type="text/javascript">
		var t=setTimeout(function(){document.location.reload(true);},3600000)
	</script>
	<% if Session("LOGIN_4_LOG") = "" then
		Session("LOGIN_4_LOG") = "NEXTAIM"
	end if %>
	<!--#INCLUDE FILE="..\Intestazione_Ridotta_include.asp" -->
	<div id="content_ridotto">
		<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<caption>STATICIZZAZIONE VOCI INDICE</caption>
			<tr>
				<th class="L2" colspan="2"> Esecuzione della staticizzazione dell'indice.</th>
			</tr>
			<tr>
				<td class="label">query:</td>
				<td class="content"><%=sql%></td>
			</tr>
			<tr>
				<td class="label">record mancanti:</td>
				<td class="content"><%=rs.recordcount%></td>
			</tr>
			<% while not rs.eof and rs.absoluteposition < 11 %>
				
				<tr>
					<td class="header" colspan="2"><%=rs.absoluteposition%> di <%=rs.recordcount%></td>
				</tr>
				<tr>
					<td class="label">co_f_key_id:</td>
					<td class="content"><%=rs("co_f_key_id")%></td>
				</tr>
				<tr>
					<td class="label">co_f_table:</td>
					<td class="content"><%=rs("co_f_table")%></td>
				</tr>
				<tr><td class="content" colspan="2">genera html</td></tr>
				<% if Index_StaticizzaItemInHtml(conn, rs("co_f_table"), rs("co_f_key_id")) then %>
					<tr><td class="content ok" colspan="2">record eseguito</td></tr>
				<% else %>
					<tr><td class="content alert" colspan="2">record fallito</td></tr>
				<% end if %>
				<%rs.movenext
			wend%>
		</table>
		<%if not rs.eof then
			%>
			<script language="JavaScript" type="text/javascript">
				document.location.reload(true);
			</script>
			<%
		end if
	set rs = nothing
		
	conn.close
	set conn = nothing
end sub


'procedura che staticizza il singolo contenuto e ne salva l'html nella parte "statica".
Function Index_StaticizzaItemInHtml(conn, table, KeyValue)
	dim sql, lingua, i, url, id, file
	dim rst, htmlCollection, baseUrl, cssList
	
	set Index.conn = conn
	set Index.content.conn = conn
	set rst = Server.CreateObject("ADODB.recordset")
	set htmlCollection = server.createobject("Scripting.Dictionary")
	set cssList = Index_SalvaCSS()
	
	'verifico se ci sono voci dell'indice pubblicate per il contenuto in posizione non "FOGLIA"
	sql = "SELECT COUNT(*) FROM v_indice WHERE isNull(idx_foglia,0)=0 AND co_f_key_id = " & KeyValue & " AND "
	if CIntero(table) > 0 then
		sql = sql &" tab_id = "& cIntero(table)
	else
		sql = sql &" tab_name LIKE '" & ParseSql(table, adChar) & "'"
	end if
	if cIntero(GetValuelist(conn, rst, sql))=0 then
	
		'nessun contenuto "non foglia", procedo alla staticizzazione di tutte le voci.
		sql = " SELECT idx_webs_id, idx_id, co_id, tab_id, " & _
			  SQL_MultiLanguage("idx_link_url_<LINGUA>,idx_link_url_rw_<LINGUA>,co_link_url_<LINGUA>,co_link_url_rw_<LINGUA>", ",") & _
			  " FROM v_indice WHERE co_f_key_id = " & KeyValue & " AND "
		if CIntero(table) > 0 then
			sql = sql &" tab_id = "& cIntero(table)
		else
			sql = sql &" tab_name LIKE '" & ParseSql(table, adChar) & "'"
		end if
		%><!--<%=sql%>--><%
		rst.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		if not rst.eof then
			baseUrl = GetSiteUrl(null, rst("idx_webs_id"), 0)
			
			while not rst.eof
				
				'scorre le lingue di ogni singolo elemento e staticizza html
				for each lingua in Application("LINGUE")
					response.write vbCrLF + "<!-- " & lingua & "-->" & vbCrLF
					CALL Index_StaticizzaUrlInHtml(htmlCollection, lingua, rst("idx_link_url_rw_" & lingua), rst("idx_link_url_" & lingua), baseUrl)
					CALL Index_StaticizzaUrlInHtml(htmlCollection, lingua, rst("co_link_url_rw_" & lingua), rst("co_link_url_" & lingua), baseUrl)
				next
				
				rst.movenext
			wend
		end if
		rst.close
		
		'comincia lavoro di salvataggio html
		conn.begintrans
		
		if htmlCollection.count > 0 then
			
			dim keys, items
			keys = htmlCollection.Keys()
			items = htmlCollection.Items()
			
			response.write vbCrLF + "<!-- REGISTRAZIONE DATI E FILES -->" & vbCrLF
			'scorre html salvati per registrazione su database e disco
			for i=0 to htmlCollection.count-1
				response.write vbCrLF + "<!-- Registra: " & keys(i) & "-->" & vbCrLF
				lingua = split(keys(i),":")(0)
				url = split(keys(i),":")(1)
				
				'inserisce record url
				sql = " INSERT INTO rel_index_url_redirect (riu_url, riu_lingua, riu_html_data) " + _
					  " VALUES ('" & ParseSql(url, adChar) & "', '" & lingua & "', GETDATE() ) "
				response.write vbCrLF + "<!-- " & sql & "-->" & vbCrLF
				CALL conn.execute(sql)
				
				'recupera valore
				sql = " SELECT top 1 riu_id FROM rel_index_url_redirect WHERE riu_url = '" & ParseSql(url, adChar) & "' ORDER BY riu_id DESC "
				response.write vbCrLF + "<!-- " & sql & "-->" & vbCrLF
				id = cIntero(GetValueList(conn, rst, sql))
				
				if id > 0 then
					'salva file e registra nome su record
					file = Index_SalvaHTML(id, items(i))
					sql = " UPDATE rel_index_url_redirect SET " + _
						  SetUpdateParamsSQL(conn, "riu_", true) + _
						  " riu_html_file='" & ParseSql(file, adChar) & "'" + _
						  " WHERE riu_id = " & id
					response.write vbCrLF + "<!-- " & sql & "-->" & vbCrLF
					CALL conn.execute(sql)
					
					CALL WriteLogAdminHttp(conn, table, KeyValue, "staticizzazione", "staticizzazone completata url:" & url, false)
					
				else
					conn.rollbacktrans
					Index_StaticizzaItemInHtml = false
					exit function
				end if
			next
			
		end if
		
		'esegue singola funzione di chiusura del lavoro a staticizzazione ultimata correttamente.
		'viene dichiarata nel file chiamante per ogni singola tipologia di elementi.
		CALL Index_StaticizzazioneEseguitaCorrettamente(conn, table, KeyValue)
		conn.committrans
		
		Index_StaticizzaItemInHtml = true
	else
		'trovate pubblicazioni non foglia: salta la staticizzazione.
		Index_StaticizzaItemInHtml = false
	end if
	
end Function


'funzione che aggiunge l'url ed il relativo html nel dizionario di memorizzazione
sub Index_StaticizzaUrlInHtml(byref htmlDictionary, lingua, url, alternativeUrl, baseUrl)
	url = trim(cString(url))
	alternativeUrl = trim(cString(alternativeUrl))
	
	if url<>"" and not (htmlDictionary is Nothing) then
		
		if left(url, 1) = "?" then
			'normalizza url se necessario
			url = "default.aspx" + url
		end if
		if left(alternativeUrl, 1) = "?" then
			'normalizza url se necessario
			alternativeUrl = "default.aspx" + alternativeUrl
		end if
		
		'url esistente: verifica se non presente nel dizionario
		if not htmlDictionary.Exists(lingua + ":" + url) then
			'l'url non esiste nel dizionario: lo genera.
			dim html
			set html = ExecuteHttpUrlGetStream(baseUrl + "/" + url)
			if html.size = 0 then
				CALL SendEmailSupportEXAttach("Errore staticizzazione: "&Request.ServerVariables("SERVER_NAME"), "Errore timeout generazione pagina per staticizzazione url: " + url, "")
				response.write "<h1>Risultato esecuzione url: " + url + "</h1><hr>"
				response.write Server.HtmlEncode(html)
				response.end
			end if 
			CALL htmlDictionary.Add(lingua + ":" + url, html)
			
			if alternativeUrl <> "" AND not htmlDictionary.Exists(lingua + ":" + alternativeUrl) then
				CALL htmlDictionary.Add(lingua + ":" + alternativeUrl, html)
			end if
		end if
	end if
end sub


'funzione che salva l'html nella matrice statica e ne restituisce il file ed il percorso.
function Index_SalvaHTML(id, html)
	dim fso, i, c, path, file, idPath
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	idPath = FixLenght(right(cString(id),7), "0", 7)

	'verifica esistenza directory registrazione statica
	if not fso.FolderExists(StaticFilePath) then
		fso.CreateFolder(StaticFilePath)
	end if
	
	'scorre caratteri dell'id per identificare la matrice corretta di salvataggio
	path = StaticFilePath
	for i=1 to len(idPath)
		c = Mid(idPath, i, 1)
		path = path + "\" + c
		if not fso.FolderExists(path) then
			fso.CreateFolder(path)
		end if
	next
	
	'calcola path definitivo file
	path = path + "\" + cString(id) + ".html"
	
	response.write vbCrLF + "<!-- " & path & "-->" & vbCrLF
	
	'esegue replace degli stili
	dim htmlString, originalCssFiles, staticCssFiles
	originalCssFiles = Session("StaticizingIndex_CSS_dictionary").Items()
	staticCssFiles = Session("StaticizingIndex_CSS_dictionary").Keys()
	
	'recupera html da sostituire
	html.position = 0
	htmlString = html.readText
	
	'cambia puntamento degli stili
	for i=0 to Session("StaticizingIndex_CSS_dictionary").count-1
		htmlString = replace(htmlString, originalCssFiles(i), staticCssFiles(i), 1, -1, vbTextCompare)
	next
	
	html.position = 0
	html.setEOS
	'riscrive il contenuto html
	html.WriteText htmlString
	html.flush
	
	'registra file su disco
	CALL html.SaveToFile(path, adSaveCreateOverWrite)
	
	Index_SalvaHTML = replace(replace(path, StaticFilePath, ""), "\", "/")
	
end function


'staticizza i file css e li salva nel dizionario in sessione per non doverli ricalcolare.
function Index_SalvaCSS()
Session("StaticizingIndex_CSS") = false
	'se non esiste lo crea.
	if not cBoolean(Session("StaticizingIndex_CSS") , false) then
		dim fso, cssStaticFileList, DatePrefix
		set cssStaticFileList = server.createobject("Scripting.Dictionary")
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		if fso.FolderExists(OriginalCssPath) then
			'se trova la cartella degli stili li staticizza e ne salva i nomi
			dim cssFile, cssFolder
			set cssFolder = fso.GetFolder(OriginalCssPath)
			for each cssFile in cssFolder.Files
				if lCase(File_Extension( cssFile.name )) = "css" then
					CALL Index_SalvaCSS_File(cssStaticFileList, fso, cssFile)
				end if
			next
			set Session("StaticizingIndex_CSS_dictionary") = cssStaticFileList
			Session("StaticizingIndex_CSS") = true
		else
			response.write "<h1>Manca directory css: non trovata</h1>"
			response.end
		end if
	end if
	
	set Index_SalvaCSS = Session("StaticizingIndex_CSS_dictionary")
	
end function

'salva il singolo file di stili aggiungendo al dizionario (staticCssFiles) che contiene la lista di sostituzioni  da fare nell'html
sub Index_SalvaCSS_File(staticCssFiles, fso, file)
	dim cssStaticPath
	
	'verifica esistenza directory registrazione statica
	if not fso.FolderExists(StaticFilePath) then
		fso.CreateFolder(StaticFilePath)
	end if
	'verifica esistenza directory registrazione statica dei file css
	if not fso.FolderExists(StaticCssFilePath) then
		fso.CreateFolder(StaticCssFilePath)
	end if
	
	'calcola nome file da staticizzare
	dim staticFileName
	staticFileName = replace(file.name, ".css", "__" & DateTimeISOFile(file.DateLastModified) & ".css")
	cssStaticPath = StaticCssFilePath & "\" & staticFileName
	
	'se non c'è il file: lo crea
	if not fso.FileExists(cssStaticPath) then
		CALL file.Copy(cssStaticPath, true)
	end if
	
	'mette a posto path upload dentro gli stili.
	
	'registra dati per sostituzione su  html
	dim staticCssUrl, originalCssUrl
	staticCssUrl = replace(StaticBaseCssPath, "\", "/") + staticFileName
	originalCssUrl = replace(BaseCssPath, "\", "/") + file.name
	
	'aggiunge file in area da sostituire
	CALL staticCssFiles.Add(staticCssUrl, originalCssUrl)
	
end Sub


%>