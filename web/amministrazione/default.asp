<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="library/InitSex.asp" -->
<!--#INCLUDE FILE="library/class_testata.asp" -->

<%

'...............................................................................
'Esegue il redirect al dominio di amministrazione
if cSTring(Application("AMMINISTRAZIONE_SERVER_NAME"))<>"" then
	if instr(1, GetCurrentFullUrl(), Application("AMMINISTRAZIONE_SERVER_NAME"), vbTextCompare) < 1 then
		response.redirect "http://" + Application("AMMINISTRAZIONE_SERVER_NAME") + "/amministrazione"
	end if
end if
'...............................................................................


dim conn, rs, sql, check_pass
dim permessi, permessiTemp, sito, diff, i, id, punto, sites_list, login, password, destination, site,cookie_permanenti, nomeApplicativo

Set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
Set rs = Server.CreateObject("ADODB.RecordSet")


Select case uCase(cString(request("LINGUA")))
	case "ITALIANO"
		Session("LINGUA") = LINGUA_ITALIANO
		CALL SetCookieLingua()
	case "INGLESE"
		Session("LINGUA") = LINGUA_INGLESE
		CALL SetCookieLingua()
end Select

Select case uCase(request("EXECUTE"))	
	case "ACCEDI","ENTER","NETACCESS"
		if request.form("EXECUTE") = "ACCEDI" or request.form("EXECUTE") = "ENTER" then
			login = Ucase(Trim(Request.Form("textLogin")))
			password = Ucase(Trim(Request.Form("passLogin")))
			cookie_permanenti = (Ucase(request("storeLogin"))="YES")
		elseif request("EXECUTE") = "NETACCESS" then		
			LoginPassword_DecodeFromNET login,password,request.querystring("DATAFROMNET")
		else
			dim credentials
			credentials = ParseSQL(request.querystring("INFO"), adChar)
			if credentials <>"" then
				CALL LoginString_Decode(request.querystring("INFO"), login, password, destination, site)
			end if
		end if
		
		if login<>"" AND CheckChar(login, LOGIN_VALID_CHARSET) then
			if password<>"" AND CheckChar(password, LOGIN_VALID_CHARSET) then
				'Verifica permessi di accesso
				
				sql = " SELECT * FROM (rel_admin_sito INNER JOIN tb_siti ON rel_admin_sito.sito_id = tb_siti.id_sito) " &_
					  " INNER JOIN tb_admin ON rel_admin_sito.admin_id=tb_admin.id_admin " &_
					  " WHERE admin_login LIKE '" & ParseSql(login, adChar) & "' " &_
					  " ORDER BY sito_id "
				'response.write sql
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				if rs.recordcount>0 then
					if credentials <>"" then
						check_pass = password
					else
						check_pass = EncryptPassword(password)
					end if
					if check_pass = rs("admin_password") then

						if IsDate(rs("admin_scadenza")) AND cString(rs("admin_scadenza"))<>"" then
							if rs("admin_scadenza") < Date() then
								Session("ERRORE") = "Accesso non consentito: Profilo non pi&ugrave; valido."
							end if
						end if
						if Application("AMMINISTRAZIONE_FILTRO_IP")<>"" then
							if uCase(Login) <> "COMBINARIO" AND uCase(Login) <> "NEXTAIM" AND request.servervariables("REMOTE_HOST") <> Application("AMMINISTRAZIONE_FILTRO_IP") then
								Session("ERRORE") = "Accesso consentito solo dagli uffici."
							end if
						end if

						if Session("ERRORE") = "" then
							'ACCESSO CONSENTITO
							'compone il cookie
							permessi = login
							sito = 0
							rs.MoveFirst
							while not rs.eof
								if sito <> rs("id_sito") then
									diff = rs("id_sito") - sito
									if diff > 1 then
										for i = 2 to diff
											permessi = permessi & ";;"
										next
									end if
									sito = rs("id_sito")
									permessi = permessi & ";;"
								end if
								permessi = permessi & rs("sito_p"&rs("rel_as_permesso")) & ","
								rs.movenext
							wend
							
							'imposta il cookie
							'if cookie_permanenti then
								CALL SetCookie(permessi)
							'end if
							CALL SetCookieLingua()
							
							Session("PERMESSI_TEMPORANEI") = permessi
							if cInteger(site)>0 AND destination<>"" then
								CALL INITSEX(site)
								response.redirect destination
							end if
							
							'se ho un solo sito entro direttamente
							dim aux
							set aux = server.createObject("ADODB.recordset")
							sql = " SELECT sito_dir FROM tb_siti WHERE id_sito IN (" + _
								  " SELECT sito_id FROM rel_admin_sito INNER JOIN tb_admin ON rel_admin_sito.admin_id=tb_admin.id_admin " + _
								  " WHERE admin_login LIKE '" & login & "') "
							aux.open sql, conn, adOpenStatic, adLockOptimistic
							if aux.recordCount = 1 and not Application("ADMIN_MULTILINGUA") then
								response.redirect aux("sito_dir")
							end if
							set aux = nothing
							
							rs.MoveFirst
						end if
					else
						Session("ERRORE") = ChooseValueByAllLanguages(Session("LINGUA"), "Accesso non consentito: Password errata!", "Access denied: wrong Password!", "", "", "", "", "", "")
					end if
				else
					Session("ERRORE") = ChooseValueByAllLanguages(Session("LINGUA"), "Accesso non consentito: Login errato!", "Access denied: wrong Login!", "", "", "", "", "", "")
				end if
				rs.close
			else
				Session("ERRORE") = ChooseValueByAllLanguages(Session("LINGUA"), "Accesso non consentito: Password mancante!", "Access denied: Password missing!", "", "", "", "", "", "")
			end if
		else
			Session("ERRORE") = ChooseValueByAllLanguages(Session("LINGUA"), "Accesso non consentito: Login mancante!", "Access denied: Login missing!", "", "", "", "", "", "")
		end if
		if Session("ERRORE") <> "" then
			CALL WriteLogAdmin(conn,"",0,"login:"&login&"; password:"&password,Session("ERRORE"))
		end if
	case "LOGOUT"
		'esegue il logout dell'applicativo
		CALL CookieLogout()
		Session.Abandon
  		response.redirect "default.asp"
end select

%>
<!DOCTYPE html>
<html>
<head>
	<meta http-equiv=Content-Type content="text/html; charset=utf-8">
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
	<title><%= ChooseValueByAllLanguages(Session("LINGUA"), "Utente registrati", "User login", "", "", "", "", "", "")%></title> 
	<link rel="stylesheet" type="text/css" href="library/stili.css">
	<SCRIPT LANGUAGE="javascript"  src="library/utils.js" type="text/javascript"></SCRIPT>
</head>
<body onload="FocusOnFirstInput(form1)">
<!-- barra alta -->

<div id="Layer0" style="position:absolute; left:0px; top:0px; width:740px; height:400px; z-index:0">
<table width="740" cellspacing="0" cellpadding="0" border="0">
	<caption class="menu">
		<table width="100%" cellspacing="0" cellpadding="0" border="0">
	  		<tr>
	  			<td width="10">&nbsp;</td>
				<td class="logout"><a href="../" class="menu" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Vai all'home page del sito", "Go to home page", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>HOME</a></td>
	  		</tr>
		</table>
	</caption>
    <% CALL WriteChiusuraIntestazione(cString(Application("BARRA_ALTA_AMMINISTRAZIONE"))) %>
</table>
<div><table>
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = ChooseValueByAllLanguages(Session("LINGUA"), "Accesso area amministrativa", "Access administrative area", "", "", "", "", "", "")
dicitura.puls_new = ""
dicitura.link_new = ""
dicitura.scrivi_con_sottosez() %>
</div>
<% if application("AMMINISTRAZIONE_DISABLED") <> "" then %>
	<div id="content" style="position:relative; top:150px;text-align:center;">
		<table align="center" style="width:400px;" cellspadding="0" cellspacing="0" class="tabella_madre">
			<caption class="border">Area amministrativa in aggiornamento</caption>
			<tr>
				<td class="content_center" style="height:250px; vertical-align:middle;">
					<b>
						Area temporaneamente disattivata per aggiornamento del sistema.<br>
						<%=application("AMMINISTRAZIONE_DISABLED")%><br>
						Contattare il supporto tecnico per maggiorni informazioni.
					</b>
					<br><br>
					Supporto:<br>
					<b>COMBINARIO</b><br>
					tel: 041 8877149<br>
					<a href="mailto:supporto@combinario.com">supporto@combinario.com</a>
				</td>
			</tr>
		</table>
	</div>
<% else %>
	<form method="post" action="default.asp" id="form1" name="form1">
	<div id="content" style="position:relative; top:150px;text-align:center;">
		<table align="center" style="width:300px;" cellspadding="0" cellspacing="0">
			<% if Application("ADMIN_MULTILINGUA") then %>
				<tr>
					<td style="text-align:right; padding-bottom:3px;">
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "lingua: ", "language: ", "", "", "", "", "", "")%>
						<a style="text-decoration:none !important;" href="?LINGUA=ITALIANO" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "italiano", "Italian", "", "", "", "", "", "")%>">
							<img src="<%=GetAmministrazionePath()%>grafica/flag_mini_it.jpg">
						</a>
						<a style="text-decoration:none !important;" href="?LINGUA=INGLESE" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "inglese", "English", "", "", "", "", "", "")%>">
							<img src="<%=GetAmministrazionePath()%>grafica/flag_mini_en.jpg">
						</a>
					</td>
				</tr>
			<% end if %>
			<tr>
				<td>
					<table border="0" cellspacing="1" cellpadding="0" class="tabella_madre" width="100%">
						<%if isCookieValid(true) then%>
							<caption class="border"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Applicazioni abilitate", "Enabled applications", "", "", "", "", "", "")%></caption>
							<%
							permessi = GetCookie(true)
							id = 0
							sites_list = ""
							'calcola lista siti con permessi
							do while permessi <> ""
								id = id+1
								punto = instr(1, permessi, ";;")
								if punto = 0 then
									exit do
								end if
								permessi = right(permessi, len(permessi)-punto-1)
								if left(permessi, 1) <> ";" then
									sites_list = sites_list & id & " ,"
								end if
							loop
							if sites_list<>"" then
								sites_list = left(sites_list, len(sites_list)-2)
								sql = "SELECT * FROM tb_siti WHERE id_sito IN (" & sites_list & ")" &_
									  " ORDER BY sito_nome"
								rs.open sql, conn, adOpenStatic, adLockOptimistic
								rs.MoveFirst%>
								<tr>
									<td class="content">
										<ul class="login">
											<%do while not rs.eof%>
												<%  if Session("LINGUA") = "en" then
														if rs("sito_nome_en") <> "" then
															nomeApplicativo = rs("sito_nome_en")
														else
															nomeApplicativo = rs("sito_nome")
														end if
													else
														nomeApplicativo = rs("sito_nome")
													end if
												%>
												<li class="login">
													<a href="<%=rs("sito_dir")%>" <% if instr(1, rs("sito_dir"), "://", vbTextCompare) then %>target="_blank"<% end if %> class="content" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Entra nell'applicazione ", "Enter ", "", "", "", "", "", "")%><%=nomeApplicativo%>" <%= ACTIVE_STATUS %>>
														<%=nomeApplicativo%>
													</a>
												</li>
												<%rs.MoveNExt
											loop%>
										</UL>
									</td>
								</tr>
								
								<%rs.close
							else%>
								<tr>
									<td class="content_b" style="padding:10px;">
										<%= ChooseValueByAllLanguages(Session("LINGUA"), "Accesso alle applicazioni non consentito.", "Access to applications is not allowed.", "", "", "", "", "", "")%>
									</td>
								</tr>
							<%end if%>
							<tr>
								<td class="content_center" style="padding:7px;">
									<input type="submit" value="LOGOUT" name="EXECUTE" class="button">
								</td>
							</tr>
							
						<% else 
							'utente non loggato nel sistema 
							%>
							<caption class="border">Login</caption>
							<tr>
								<td class="content" colspan="2" style="padding:5px; padding-left:6px;">
									<%= ChooseValueByAllLanguages(Session("LINGUA"), "Scrivi un <u>login</u> ed una <u>password</u> validi per accedere all'area amministrativa.", "Write <u>login</u> and <u>password</u> for access the administration.", "", "", "", "", "", "")%>
								</td>
							</tr>
							<tr>
								<td class="label_login">
									LOGIN
								</td>
								<td class="content" style="padding:2px; text-align:left;">
									<input type="text" id="textLogin" name="textLogin" value="" size="20" style="text-transform:uppercase;" tabindex="1">
								</td>
							</tr>
							<tr>
								<td class="label_login">
									PASSWORD
								</td>
								<td class="content" style="padding:2px; text-align:left;">
									<input type="password" id="passLogin" name="passLogin" value="" size="20" tabindex="2">
								</td>
							</tr>
							<!--
							<tr>
									<td class="label_login">&nbsp;</td>
									<td class="content" style="padding:2px;">
										<table width="100%" cellspacing="0" cellpadding="0" border="0">
										<tr>
											<td align="right" valign="top"><input type="checkbox" class="checkbox" name="storeLogin" id="storeLogin" value="yes" checked></td>
											<td>Ricodati di me per tutta la giornata su questo computer</td>
										</tr>
										</table>
									</td>
							</tr>
							-->
							<tr>
								<td class="content">&nbsp;</td>
								<td class="content" style="padding:6px; padding-left:2px; text-align:left;">
									<input type="submit" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "ACCEDI", "ENTER", "", "", "", "", "", "")%>" name="EXECUTE" class="button" tabindex="3">
									<input type="reset" value="RESET" name="resetAll" class="button">
								</td>
							</tr>
						<% end if %>
					</table>
				</td>
			</tr>
		</table>
	</div>
	</form>
<% end if %>
</body>
</html>
<%
conn.close
set rs = nothing
set conn = nothing
%>