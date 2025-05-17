<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit
response.charset = "UTF-8"
Response.Buffer = True%>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<%dim conn, sql, rs
set conn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.Recordset")
conn.open Application("DATA_ConnectionString")

sql = "SELECT F_allegato, F_encoded_name, F_encoded_path, F_original_name FROM tb_files WHERE F_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

if rs("F_Allegato") then
	sql = "SELECT (COUNT(*)) AS F_VALID FROM rel_documenti_files WHERE rel_files_id=" & cIntero(request("ID")) & _
	 	  " AND rel_documento_id IN (SELECT doc_id FROM tb_documenti " & _
		  " WHERE doc_creatore_id="& Session("ID_ADMIN") &" OR " & AL_query(conn, AL_DOCUMENTI) & ") " 
else
	sql = "SELECT (COUNT(*)) AS F_VALID FROM tb_files WHERE F_id = " & cIntero(request("ID")) & _
		  " AND (F_original_path='' OR F_original_path LIKE '" & Session("LOGIN_4_LOG") & "')"
end if
if cInt(conn.execute(sql)("F_VALID")) > 0 then
	'l'utente ha i permessi per vedere il file
	On error resume next
	dim objStream
	set objStream = Server.CreateObject("ADODB.Stream") 
	objStream.Open 
	objStream.Type = adTypeBinary 
	'carica il file in uno stream
	objStream.LoadFromFile Application("IMAGE_PATH") & "\docs\" & rs("F_encoded_path") & "\" & rs("F_encoded_name")
	
	response.clear
	'restituisce il file come risposta
	response.AddHeader "Content-Disposition", "attachment; filename=" & rs("F_original_name")
	response.AddHeader "Content-Length", objStream.Size 
	response.BinaryWrite objStream.Read

	'controlla che tutto sia andato a buon fine
	if err.number<>0 then
		response.ContentType = "text/html"
		response.write "Impossibile trovare il file!"
	else
		'invia risultato al browser
		response.Flush
	end if

	objStream.Close 
	set objStream = Nothing 
else
	'l'utente non ha i permessi di vedere il file%>
	<script language="JavaScript" type="text/javascript">
		alert('Non si posseggono i permessi per visualizzare il file');
		window.close();
	</script>
<%end if
rs.close
conn.close
set rs = nothing
set conn = nothing
%>