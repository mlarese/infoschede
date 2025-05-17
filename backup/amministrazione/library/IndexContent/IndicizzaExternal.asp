<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="Tools_IndexContent.asp" -->
<%
'check dei permessi

On Error Resume Next

if isFromLocal() then
	if request("CANC") <> "" then
		dim tabName
		tabName = GetValueList(index.conn, NULL, "SELECT tab_name FROM tb_siti_tabelle WHERE tab_id = "& CIntero(request("co_F_table_id")))
		'cancella
		index.content.conn.BeginTrans()
		CALL index.content.DeleteAll(tabName, request("co_F_key_id"))
		index.content.conn.CommitTrans()
	else
		'indicizza
		index.conn.BeginTrans()
		CALL Index_UpdateItem(index.conn, request("co_F_table_id"), request("co_F_key_id"), true)
		index.conn.CommitTrans()
	end if
	
	response.write "OK"
else
	response.write "KO"
 end if
 
If Err.Number <> 0 Then
	CALL SendEmailSupportEX("Errore indice: "& Request.ServerVariables("SERVER_NAME"), _
							"url= http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME") & "?" & Request.ServerVariables("QUERY_STRING")) 
	response.write "KO"
end if
%>