<%@LANGUAGE="VBSCRIPT"%>
<% option explicit %>
<!--#INCLUDE VIRTUAL="amministrazione/library/Tools.asp"-->
<!--#INCLUDE VIRTUAL="amministrazione/library/Tools4Plugin.asp"-->
<!--#INCLUDE VIRTUAL="amministrazione/library/ClassConfiguration.asp"-->
<%
'configuarazione proprieta' oggetto
dim Config
dim conn, rs, sql,lingua,pagina,web,title,tag,stile
set Config = new Configuration
'impostazione delle proprieta' di default

Config.AddDefault "StileH",""
Config.AddDefault "TagH","1" 'valori numerici da 1 a 6
'impostazione dati di default: da sovrascrivere via propriet&agrave; del plugin o dagli stili
Config.AddDefault "testo_IT", ""
Config.AddDefault "testo_EN", ""
Config.AddDefault "testo_FR", ""
Config.AddDefault "testo_DE", ""
Config.AddDefault "testo_ES", ""

'caricamento proprieta' specifiche
Config.SetConfigurationString(Session("CONFSTR"))
' creo il tag h in base al parametro
tag = cIntero(Config("TagH"))
web = cInteger(Session("AZ_ID"))
if tag >=1 and tag<=6 then
	tag = "h" & tag
end if
lingua = Session("LINGUA")
if Config("testo_IT")<>"" then
	title = CBL(Config, "testo") 
else
	if request.querystring("pagina")<>"" AND IsNumeric(request.querystring("pagina")) then
		'recupera pagina
		pagina = cIntero(request.querystring("pagina")) 
		set conn = Server.CreateObject("ADODB.Connection")
		set rs = Server.CreateObject("ADODB.Recordset")
		conn.Open Application("l_conn_ConnectionString"),"",""
		'*************************************************************************************************************************
		'recupera dati della paginasito corrispondente per recuperare nomi pagina e metatag
		'*************************************************************************************************************************
		sql = " SELECT * FROM tb_paginesito WHERE id_web=" & web & _
			  " AND (id_pagDyn_" & lingua & "=" & pagina & " ) "	
		
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		if not rs.eof then
			title = CBL(rs, "nome_ps")
		else
			title = sql
		end if
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end if
end if
if Config("StileH")<>"" then
	stile = " class="""&Config("StileH")&""" "
else
	stile = ""
end if
response.write "<"+tag+stile+">"+title+"</"+tag+">"+vbCRLF 
%> 