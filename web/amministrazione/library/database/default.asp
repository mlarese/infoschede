<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools4DataBase.asp" -->
<!--#INCLUDE FILE="../class_testata.asp" -->
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="../ClassCryptography.asp"-->
<%
dim conn, rs, sql
Set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
Set rs = Server.CreateObject("ADODB.RecordSet")

if Session("UTENTE_MANUTENZIONE") = "" then
	if uCase(request("EXECUTE")) = "ACCEDI" OR IsLocal() then
		dim login, password
		
		if IsLocal() then
			login = "NEXTAIM"
			password = "nextaim"
		else
			login = Ucase(Trim(Request.Form("textLogin")))
			password = Ucase(Trim(Request.Form("passLogin")))
		end if
		
		if login<>"" AND CheckChar(login, LOGIN_VALID_CHARSET) then
			if password<>"" AND CheckChar(password, LOGIN_VALID_CHARSET) then
				
				sql = " SELECT * FROM (rel_admin_sito INNER JOIN tb_siti ON rel_admin_sito.sito_id = tb_siti.id_sito) " &_
					  " INNER JOIN tb_admin ON rel_admin_sito.admin_id=tb_admin.id_admin " &_
					  " WHERE "
				if isLocal() then
					sql = sql & " (admin_login LIKE 'NEXTAIM' OR admin_login LIKE 'COMBINARIO') "
				else
					sql = sql & " admin_login LIKE '" & ParseSql(login, adChar) & "'"
				end if
				sql = Sql & " AND tb_siti.id_sito = " & NEXTPASSPORT & _
					  " ORDER BY sito_id "
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				
				if rs.recordcount>0 then
					'if UCASE(password) = UCASE(rs("admin_password")) OR _
					if EncryptPassword(password) = rs("admin_password") OR _
					   (IsLocal() AND  (login = "NEXTAIM" OR login="COMBINARIO")) then
						if IsDate(rs("admin_scadenza")) AND cString(rs("admin_scadenza"))<>"" then
							if rs("admin_scadenza") < Date() then
								Session("ERRORE") = "Accesso non consentito: Profilo non pi&ugrave; valido."
							end if
						end if
						
						if Session("ERRORE") = "" then
							'accesso consentito
							
							Session("LOGIN_4_LOG") = login
							Session("UTENTE_MANUTENZIONE") = login
							session("ID_ADMIN") = rs("admin_id")
							response.redirect "ConsoleManutenzione.asp"
						end if
					else
						Session("ERRORE") = "Accesso non consentito: Password errata!"
					end if
				else
					Session("ERRORE") = "Accesso non consentito: Login errato!"
				end if
				rs.close
			else
				Session("ERRORE") = "Accesso non consentito: Password mancante!"
			end if
		else
			Session("ERRORE") = "Accesso non consentito: Login mancante!"
		end if
	end if
else
	response.redirect "ConsoleManutenzione.asp"
end if
%>
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
	<title>Amministrazione aggiornamenti database</title>
	<link rel="stylesheet" type="text/css" href="../stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body onload="FocusOnFirstInput(form1)">
<!-- barra alta -->
<div id="Layer0" style="position:absolute; left:0px; top:0px; width:740px; height:400px; z-index:0">
<!-- barra alta -->
<table width="740" cellspacing="0" cellpadding="0" border="0">
	<caption style="border:0px;">
		<table width="100%" cellspacing="0" cellpadding="0" border="0">
	  		<tr>
	  			<td class="logout">
					<a href="../logout.asp" class="menu" title="esci dall'appplicazione e torna all'area di login" <%= ACTIVE_STATUS %>>AMMINISTRAZIONE</a>
					&nbsp;
					&nbsp;
					&nbsp;
					<a href="../" class="menu" title="vai all'home page del sito" <%= ACTIVE_STATUS %>>HOME</a>
				</td>
	  		</tr>
		</table>
	</caption>
    <% CALL WriteChiusuraIntestazione("") %>
    <% 	dim header
    set header = New testata 
    header.iniz_sottosez(0)
    header.sezione = "gestione database"
    header.scrivi_con_sottosez() %>
</table>
</div>
<form method="post" action="default.asp" id="form1" name="form1">
<div id="content" style="position:relative; top:150px;text-align:center;">
	<center>
		<table border="0" cellspacing="1" cellpadding="0" class="tabella_madre" style="width:40%;">
			<caption class="border warning">Login area manutenzione</caption>
			<tr>
				<td class="content" colspan="2" style="padding:5px; padding-left:6px;">
					Scrivi un <u>login</u> ed una <u>password</u> validi per accedere all'area amministrativa.
				</td>
			</tr>
			<tr>
				<td class="label_login">
					LOGIN
				</td>
				<td class="content" style="padding:2px;">
					<input type="text" id="textLogin" name="textLogin" value="" size="20" style="text-transform:uppercase;" tabindex="1">
				</td>
			</tr>
			<tr>
				<td class="label_login">
					PASSWORD
				</td>
				<td class="content" style="padding:2px;">
					<input type="password" id="passLogin" name="passLogin" value="" size="20" tabindex="2">
				</td>
			</tr>
				<tr>
					<td class="content">&nbsp;</td>
					<td class="content" style="padding:6px; padding-left:2px;">
						<input type="submit" value="ACCEDI" name="EXECUTE" class="button" tabindex="3">
						<input type="reset" value="RESET" name="resetAll" class="button">
					</td>
				</tr>
		</table>
	</center>
</div>
</form>
</body>
</html>
<%
conn.close
set rs = nothing
set conn = nothing
%>