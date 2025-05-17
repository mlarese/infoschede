<%@ Language=VBScript CODEPAGE=65001%>
<% option explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/class_testata.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<html>
<head>
	<title>Modifica password</title>
	<link rel="stylesheet" type="text/css" href="../library/stili.css">
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" onload="window.focus()">
<% 
	dim intestazione
	'disegna intestazione
	set intestazione= New testata
	intestazione.sezione = "Gestione utenti area amministrativa - password"
	intestazione.scrivi_ridotta()
	
dim conn, rs, sql, saved, password
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT * FROM tb_admin WHERE ID_admin=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
saved = false

if request("salva")<>"" then
	if request("tft_admin_password")="" OR  (not CheckChar(request("tft_admin_password"), LOGIN_VALID_CHARSET)) then
		Session("ERRORE") = "Password mancante o non valida! Utilizzare solo caratteri alfanumerici o &quot;_&quot;"
	elseif uCase(request("tft_admin_password")) <> uCase(request("conferma_password")) then
		Session("ERRORE") = "Errore nella conferma della password!"
	else
		'esegue la modifica
		password = UCASE(request("tft_admin_password"))
		'cripto la password
		password = EncryptPassword(password)
		rs("admin_password") = password
		rs.Update
		
		'registra su log la modifica
		CALL LogModificaUtente(conn, "AdminPassword", rs("id_admin") , rs("admin_login"), rs("admin_email"))
		
		saved = true
	end if
end if
%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Utente: &quot;<%= rs("admin_cognome") %>&nbsp;<%= rs("admin_nome") %>&quot;</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="javascript:window.close();">
							CHIUDI
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<% if saved then %>
			<tr><th colspan="2">NUOVA PASSWORD SALVATA</th></tr>
			<tr>
				<td class="content_b" colspan="2">
					La password &egrave; stata cambiata correttamente.
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="2">
					<input onclick="window.close();" type="button" class="button" name="annulla" value="OK">
				</td>
			</tr>
		<% else %>
			<% if Session("ERRORE")<>"" then %>
				<tr>
					<td colspan="2" class="errore"><%= Session("ERRORE") %></td>
				</tr>
				<% Session("ERRORE") = ""
			else%>
				<tr><th colspan="2">NUOVA PASSWORD</th></tr>
			<%end if %>
			<tr>
				<td class="label" style="width:30%;">Password:</td>
				<td class="content" style="width:70%;">
					<input type="password" class="text" name="tft_admin_password" value="<%= request("tft_admin_password") %>" maxlength="50" size="20">
				</td>
			</tr>
			<tr>
				<td class="label">Conferma password:</td>
				<td class="content">
					<input type="password" class="text" name="conferma_password" value="<%= request("conferma_password") %>" maxlength="50" size="20">
				</td>
			</tr>
			<tr>
				<td class="note" colspan="2">
					Per la composizione della password utilizzare solo caratteri alfanumerici o &quot;_&quot; 
					indifferentemente con lettere minuscole o maiuscole, ma senza spazi bianchi.
					<span style="letter-spacing:2px;">(<%= LOGIN_VALID_CHARSET %>)</span>
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="2">
					<input type="submit" class="button" name="salva" value="CAMBIA PASSWORD">
					<input onclick="window.close();" type="button" class="button" name="annulla" value="ANNULLA">
				</td>
			</tr>
		<% end if %>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>