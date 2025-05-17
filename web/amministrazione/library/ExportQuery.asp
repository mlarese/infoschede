<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit
response.charset = "UTF-8"
response.buffer = true%>
<!--#INCLUDE FILE="Tools.asp" -->
<!--#INCLUDE FILE="Tools4Admin.asp" -->
<!--#INCLUDE FILE="ExportTools.asp" -->
<%
Server.ScriptTimeout = 2147483647

dim ConnString, Query, Format
ConnString = request("conn")
Query = request("query")
Format = request("format")

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application(ConnString)
set rs = Server.CreateObject("ADODB.RecordSet")
sql = Session(Query)

'response.write Query & "=" & Session(Query) + "<br>"
'response.write Format + "<br>"
'response.write ConnString + "<br>"


rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

Select Case Format
	case FORMAT_HTML
		CALL ExportRecordset_HTML(rs)
	case FORMAT_XML
		CALL ExportRecordset_XML(rs)
	case FORMAT_EXCEL_FILE
		CALL Export_Excel2000(rs)
	case FORMAT_EXCEL_XML
		CALL ExportRecordset_EXCEL_XML(rs)
	case FORMAT_TXT
		CALL ExportRecordset_TXT(rs)
	case FORMAT_ACCESS
	 	CALL ExportRecordset_Access(rs)
end select

rs.close
set rs = nothing
conn.close
set conn =nothing
%>
