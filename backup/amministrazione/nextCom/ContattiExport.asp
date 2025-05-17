<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = true %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/TOOLS4ADMIN.ASP" -->
<!--#INCLUDE FILE="Tools_export.ASP" -->
<%

if request("esporta")="" then%>
	<html>
		<head>
			<title>Export dati ricercati</title>
			<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
			<META NAME="copyright" CONTENT="Copyright &copy;2003 - next-aim.com">
			<meta name="robots" content="noindex,nofollow" />
			<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
			<link rel="stylesheet" type="text/css" href="../library/stili.css">
			<SCRIPT LANGUAGE="javascript"  src="../library/utils.js" type="text/javascript"></SCRIPT>
		</head>
		
		<body onload="window.focus();" leftmargin="4" topmargin="3">
			<form action="" method="post" target="export" id="export" name="Export">
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						<table width="100%" border="0" cellspacing="0">
							<tr>
								<td class="caption">Export dei dati</td>
								<td align="right" style="padding-right:5px;"><a class="button" href="javascript:window.close();">CHIUDI</a></td>
							</tr>
						</table>
					</caption>
					<tr><th colspan="2">FORMATO DEI DATI</td></tr>
					<tr>
						<td class="content" width="15%"><INPUT  value="<%=FORMAT_EXCEL_XML%>" class="checkbox" type="radio" name="export"></td>
						<td class="content">foglio Excel XP o successivo</td>
					</tr>
					<tr>
						<td class="content"><INPUT  value="<%=FORMAT_EXCEL_FILE%>" class="checkbox" type="radio" name="export"></td>
						<td class="content">foglio Excel 2000 o precedente</td>
					</tr>
					<tr>
						<td class="content"><INPUT  value="<%=FORMAT_HTML%>" class="checkbox" type="radio" name="export"></td>
						<td class="content">pagina HTML</td>
					</tr>
					<tr>
						<td class="content"><INPUT  value="<%=FORMAT_TXT%>" class="checkbox" type="radio" name="export"></td>
						<td class="content">file di testo separato da &quot;;&quot;</td>
					</tr>
					<tr>
						<td class="content"><INPUT  value="<%=FORMAT_XML%>" class="checkbox" type="radio" name="export"></td>
						<td class="content">file XML</td>
					</tr>
					<tr>
						<td class="footer" colspan="2">
							<input onclick="window.close();" type="button" name="esporta" value="ANNULLA" class="button" style="width: 49%;">
							<input type="submit" name="esporta" value="ESPORTA DATI" class="button" style="width: 49%;">
						</td>
					</tr>
				</table>
			</form>
		</body>
	</html>
<%else
	dim conn, rs, rsv, sql, sessionSQL
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open Application("DATA_ConnectionString")
	set rs = Server.CreateObject("ADODB.RecordSet")
	set rsv = Server.CreateObject("ADODB.RecordSet")
	
	'recupera tipi di valori
	rsv.open "tb_TipNumeri", conn, adOpenStatic, adLockReadOnly, adCmdTable
	
	sessionSQL = request.querystring("sql")
	if request.querystring("sql") = "" then
		sessionSQL = "SQL_ELENCO"
	end if
	
	sql = "SELECT IDElencoIndirizzi " & right(Session(sessionSQL), len(Session(sessionSQL)) + 1 - instr(1, Session(sessionSQL), "FROM", vbTextCompare))
	if instrRev(sql, "ORDER BY", vbTrue,vbTextCompare) > 0 then
		sql = left(sql, instrRev(sql, "ORDER BY", vbTrue,vbTextCompare) - 1)
	end if
	sql = "SELECT (IDElencoIndirizzi) AS ID, " &_
				  "(TitoloElencoIndirizzi) AS titolo, " & _
				  "(NomeElencoIndirizzi) AS nome, " & _
				  "(CognomeElencoIndirizzi) AS cognome, " & _
				  "(NomeOrganizzazioneElencoIndirizzi) AS [ente-organizzazione], " & _
				  "(QualificaElencoIndirizzi) AS [ruolo / qualifica], " & _
				  "(IndirizzoElencoIndirizzi) AS indirizzo, " & _
				  "(CAPElencoIndirizzi) AS cap, " & _
				  "(LocalitaElencoIndirizzi) AS [Localita], " & _
				  "(CittaElencoIndirizzi) AS citta, " & _
				  "(statoProvElencoIndirizzi) AS provincia, " & _
				  "(CountryElencoIndirizzi) AS Nazione, " & _
				  "(DTNASCElencoIndirizzi) AS [Data di nascita], " & _
				  "(LuogoNascita) AS [Luogo di nascita], " & _
				  "(CF) AS [Codice fiscale], " & _
				  "(partita_iva) AS [Partita IVA], " & _
				  "NoteElencoIndirizzi AS [Note]" & _
				  "FROM tb_indirizzario WHERE IDElencoIndirizzi IN (" &_
				  sql & ") ORDER BY ModoRegistra"
	rs.Open sql, conn, AdOpenforwardOnly, adLockReadOnly, adCmdText
	
	Server.ScriptTimeout = 3600
	  
	Select case request("Export")
		case FORMAT_EXCEL_XML
			CALL Export_ExcelXP(conn, rs, sql, rsv)
		case FORMAT_EXCEL_FILE
			CALL Export_Excel2000(conn, rs, sql, rsv)
		case FORMAT_HTML
			CALL Export_HTML(conn, rs, sql, rsv)
		case FORMAT_TXT
			CALL Export_TXT(conn, rs, sql, rsv)
		case FORMAT_XML
			CALL Export_XML(rs, false)
		case else
			response.redirect "ContattiExport.asp"
	end select
	
	rs.close
	rsv.close
	conn.close 
	set rs = nothing
	set rsv = nothing
	set conn = nothing
end if

%>