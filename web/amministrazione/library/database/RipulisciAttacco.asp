<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<% Server.ScriptTimeout = 1073741824 %>
<!--#INCLUDE FILE="Tools4DataBase.asp" -->
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<html>
<head>
	<title>Ripulitura attacco Provincia</title>
	<link rel="stylesheet" type="text/css" href="../stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0,5"  onload="window.focus();">
<%
dim conn, rst, rsc, rsf, sql, columns, field, update
Set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
conn.CommandTimeout = 90
Set rst = Server.CreateObject("ADODB.RecordSet")
Set rsc = Server.CreateObject("ADODB.RecordSet")
set rsf = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT * " + _
	  " FROM sysobjects " + _
	  " WHERE sysobjects.type='U' " + _
	  "		  AND sysobjects.id IN (SELECT syscolumns.id FROM syscolumns " + _
	  "								WHERE xtype IN (35, 99, " + _
	  "											 	175, 239, 231, 167) ) " + _
	  " AND NOT (name like 'itb_eventi') " + _
	  " AND NOT (name like 'itb_anagrafiche') " + _
	  " AND NOT (name like 'irel_luoghi') " + _
	  " AND NOT (name like 'irel_periodi') " + _
	  " AND NOT (name like 'irel_anagrafiche_descrTipi') " + _
	  " AND NOT (name like 'irel_eventi_descrCat') "
rst.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

while not rst.eof
	sql = "SELECT name, xtype FROM syscolumns WHERE id = " & rst("id") & " AND xtype IN (35, 99, 175, 239, 231, 167)"
	rsf.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	
	while not rsf.eof
		if rsf("xtype") = 99 OR rsf("xtype") = 35 then
			'campi memo
			sql = "SELECT " & rsf("name") & " FROM " & rst("name") & " WHERE " & rsf("name") & " LIKE '%</title%' "
			response.write sql & "<br>"
			rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			while not rsc.eof
				update = false
				for each field in rsc.fields
					if cString(field.value)<>"" then
						if instr(1, field.value, "</title>", vbTextCompare)>0 then
							'response.write "<h6>prima</h6><pre>" & Server.htmlencode(field.value) & "</pre>"
							rsc(field.name) = left(field.value, instr(1, field.value, "</title>", vbTextCompare)-1)
							'response.write "<h6>dopo</h6><pre>" & Server.htmlencode(field.value) & "</pre>"
							response.write rsc.Absoluteposition & " di " & rsc.recordcount & "<br>"
							update = true
						end if
					end if
				next
				if update then
					rsc.update
				end if
				
				rsc.movenext
			wend
			rsc.close
			
		else
			'campi normali
			sql = "UPDATE " & rst("name") & _
				  " SET  " & rsf("name") & " = left(" & rsf("name") & ", charindex('title', " & rsf("name") & ")-1) " + _
				  " where " & rsf("name") & " LIKE '%title%' "
			CALL conn.execute(sql)
		end if
		rsf.movenext
	wend
	
	rsf.close
	
	rst.movenext
wend
	
rst.close



%>
