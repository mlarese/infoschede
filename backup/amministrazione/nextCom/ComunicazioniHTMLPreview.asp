<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<!--#INCLUDE FILE="../library/tools.asp" -->
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->

<%
if Session("LOGIN_4_LOG")="" AND request.form("HTML_by_post") <> "1" then	
	'utente non loggato
	response.write "Accesso all'email vietato! - Comunicazioni HTML Previewer"
	response.end
end if

dim conn, MessageType, Messaggio
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set Messaggio = new Mailer

MessageType = cIntero(request("type"))

if cIntero(request("HTML_by_post")) > 0 then
	response.write Messaggio.LoadHTMLCode(conn, request.form("HTML_title"), request.form("HTML_body"), "")
else
	response.write Messaggio.LoadHTMLCode(conn, ComunicazioniNew_Wizard_Session_GetField(messageType, "tft_email_object"), _
												ComunicazioniNew_Wizard_Session_GetField(MessageType, "tft_email_text"), "")
end if


%>
						
