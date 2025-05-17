<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<% Server.ScriptTimeout = 100000 %>
<!--#INCLUDE FILE="Tools4DataBase.asp" -->
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<% 
'*****************************************************************************************************************
'verifica dei permessi
CALL VerificaPermessiUtente(true)
'*****************************************************************************************************************
%>
<html>
<head>
	<title>Amministrazione aggiornamenti database</title>
	<link rel="stylesheet" type="text/css" href="../stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0" onload="window.focus();">
<%
dim conn, rst, rsd

set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application(request("ConnString")), "", ""
%>
<table width="740" cellspacing="1" cellpadding="0" border="0" class="tabella_madre" style="border-bottom:0px;">
	<caption style="border:0px;">
		<a style="float:right;" href="javascript:close();" class="menu" name="top">CHIUDI</a>
	</caption>
	<tr><th colspan="2">DATI GENERALI DATABASE</th></tr>
    <tr>
        <td class="label">database:</td>
        <td class="content"><%= GetDatabaseName(conn) %></td>
    </tr>
    <tr>
        <td class="label">dimensione totale:</td>
        <td class="content"><%= DatabaseSize(conn) %></td>
    </tr>
	<tr><th colspan="2">DIMENSIONI TABELLE DEL DATABASE</th></tr>
</table>

<% 'recupera elenco tabelle
set rst = conn.OpenSchema(adSchemaTables)
%>
<table width="740" cellspacing="1" cellpadding="0" border="0" class="tabella_madre" style="margin-bottom:20px;">
    <tr>
        <td class="label_no_width" colspan="4">
            Trovati n&deg; <%= CounrRecordsetRow(rst) %> oggetti nel database
        </td>
    </tr>
    <tr>
        <th class="L2">NOME OGGETTO</th>
        <th class="L2" width="15%">TIPO</th>
        <th class="l2_center" width="15%">NUMERO RIGHE</th>
        <th class="l2_center" width="15%">DIMENSIONE FISICA (KB)</th>
    </tr>
    <% while not rst.eof %>
        <tr>
            <td class="content"><%= rst("table_name") %></td>
            <td class="content"><%= rst("table_type") %></td>
            <% if rst("table_type") = "TABLE" then 
                'recupera dati sulla dimensione dell'oggetto
                set rsd = conn.Execute("EXEC sp_spaceused [" & rst("table_name") & "]")%>
                <td class="content_right" title="<%= rsd("rows") %>">
                    <%= FormatPrice(cIntero(rsd("rows")), 0, true) %>
                </td>
                <td class="content_right" title="<%= rsd("data") %>">
                    <%= FormatPrice(cIntero(replace(rsd("data"), " KB", "")), 0, true) %>
                </td>
            <% else %>
                <td class="content_right">- -</td>
                <td class="content_right">- -</td>
            <% end if %>
        </tr>
        <% rst.movenext
    wend%>
    <tr>
        <td class="footer" colspan="4">
            <input type="button" class="button" name="chiudi" value="CHIUDI" onclick="window.close();">
        </td>
    </tr>
</table>    
    

    
<% 

function CounrRecordsetRow(rs)
    CounrRecordsetRow = 0
    if not rs.eof then
        while not rs.eof
            CounrRecordsetRow = CounrRecordsetRow + 1
            rs.movenext
        wend
        rs.movefirst
    end if
end function

 %>