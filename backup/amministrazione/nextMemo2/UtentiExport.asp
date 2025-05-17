<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = true %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/TOOLS4ADMIN.ASP" -->
<!--#INCLUDE FILE="Tools_exportUsers.ASP" -->
<%

if request("esporta")="" then%>
	<html>
		<head>
			<title>Export dati ricercati</title>
			<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
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
						<td class="content" width="15%"><INPUT  value="XLS_XP" class="checkbox" type="radio" name="export"></td>
						<td class="content">foglio Excel XP o successivo</td>
					</tr>
					<tr>
						<td class="content"><INPUT  value="XLS_2000" class="checkbox" type="radio" name="export"></td>
						<td class="content">foglio Excel 2000 o precedente</td>
					</tr>
					<tr>
						<td class="content"><INPUT  value="HTML" class="checkbox" type="radio" name="export"></td>
						<td class="content">pagina HTML</td>
					</tr>
					<tr>
						<td class="content"><INPUT  value="TXT" class="checkbox" type="radio" name="export"></td>
						<td class="content">file di testo separato da &quot;;&quot;</td>
					</tr>
					<tr>
						<td class="content"><INPUT  value="XML" class="checkbox" type="radio" name="export"></td>
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
		sessionSQL = "SQL_UTENTI"
	end if
	
	sql = "SELECT IDElencoIndirizzi " & right(Session(sessionSQL), len(Session(sessionSQL)) + 1 - instr(1, Session(sessionSQL), "FROM tb_indirizzario ", vbTextCompare))
	if instrRev(sql, "ORDER BY", vbTrue,vbTextCompare) > 0 then
		sql = left(sql, instrRev(sql, "ORDER BY", vbTrue,vbTextCompare) - 1)
	end if
	sql = "SELECT (IDElencoIndirizzi) AS ID, " &_
				  "(TitoloElencoIndirizzi) AS titolo, " & _
				  "(NomeElencoIndirizzi) AS nome, " & _
				  "(CognomeElencoIndirizzi) AS cognome, " & _
				  "(NomeOrganizzazioneElencoIndirizzi) AS [ente-organizzazione], " & _
				  "(IndirizzoElencoIndirizzi) AS indirizzo, " & _
				  "(CAPElencoIndirizzi) AS cap, " & _
				  "(LocalitaElencoIndirizzi) AS [Localita], " & _
				  "(CittaElencoIndirizzi) AS citta, " & _
				  "(statoProvElencoIndirizzi) AS provincia, " & _
				  "(CountryElencoIndirizzi) AS Nazione, " & _
				  "(DTNASCElencoIndirizzi) AS [Data di nascita], " & _
				  "(LuogoNascita) AS [Luogo di nascita], " & _
				  "(CF) AS [Codice fiscale - Partita IVA], " & _
				  "(ut_id) AS [ID Utente], " & _
				  "(ut_login) AS [Login], " & _
				  "(ut_password) AS [Password], " & _
				  "(ut_abilitato) AS [Abilitato], " & _
				  "(ut_ScadenzaAccesso) AS [Scadenza Accesso] " & _
				  "FROM tb_indirizzario INNER JOIN tb_utenti ON tb_indirizzario.IDElencoIndirizzi=tb_utenti.ut_NextCom_ID WHERE IDElencoIndirizzi IN (" &_
				  sql & ") ORDER BY ModoRegistra"
				  
	rs.Open sql, conn, AdOpenforwardOnly, adLockReadOnly, adAsyncFetch
	
	Server.ScriptTimeout = 3600
	  
	Select case request("Export")
		case "XLS_XP"
			CALL Export_ExcelXP(conn, rs, sql, rsv)
		case "XLS_2000"
			CALL Export_Excel2000(conn, rs, sql, rsv)
		case "HTML"
			CALL Export_HTML(conn, rs, sql, rsv)
		case "TXT"
			CALL Export_TXT(conn, rs, sql, rsv)
		case "XML"
			CALL Export_XML(rs, false)
		case else
			response.redirect "UtentiExport.asp"
	end select
	
	rs.close
	rsv.close
	conn.close 
	set rs = nothing
	set rsv = nothing
	set conn = nothing
end if

%>
