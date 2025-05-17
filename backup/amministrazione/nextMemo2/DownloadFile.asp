<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Response.Buffer = True  %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="Tools_Categorie.asp" -->
<% 
dim conn, rs, sql, File, DipID, UtID
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

'recupera id utente
if request("DIP")<>"" then
	sql = "SELECT 0 AS Ut_ID, Id_admin FROM tb_admin WHERE admin_login LIKE '" & ParseSQL(request("DIP"), adChar) & "'"	
else
	sql = "SELECT Ut_ID, 0 AS Id_admin FROM tb_utenti WHERE ut_login LIKE '" & ParseSQL(request("UT"), adChar) & "'"	
end if
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
if not rs.eof then
	DipID = rs("id_admin")
	UtID = rs("Ut_ID")
else
	dipID = 0
	utID = 0
end if
rs.close


'registrazione download
sql = "INSERT INTO log_documenti (log_ut_id, log_dip_id, log_doc_id, log_data) " &_
	  " VALUES (" & UtID & ", " & DipID & "," & cInteger(request("ID")) & ", " & SQL_now(conn) & ")"
CALL conn.execute(sql, 0, adExecuteNoRecords)

'recupera nome del file
sql = "SELECT doc_File_it FROM mtb_documenti WHERE doc_id =" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
File = rs("doc_File_it")
rs.close

conn.close
set rs = nothing
set conn = nothing

'On error resume next
dim objStream
set objStream = Server.CreateObject("ADODB.Stream") 
objStream.Open 
objStream.Type = adTypeBinary 
objStream.LoadFromFile Application("IMAGE_PATH") & "/" & Application("AZ_ID") & "/images/" & File

response.clear
response.AddHeader "Content-Disposition", "attachment; filename=" & File
response.AddHeader "Content-Length", objStream.Size 
response.BinaryWrite objStream.Read


if err.number<>0 then
	response.ContentType = "text/html"
	response.write "Impossibile trovare il file!"
else
	'invia risultato al browser
	response.Flush
end if

objStream.Close 
set objStream = Nothing 
%>
