<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/tools.asp" --> 
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->  
<% 
if Session("LOGIN_4_LOG")="" AND cString(request("KEY"))="" then	
	'utente non loggato
	response.write "Accesso all'email vietato!"
	response.end
end if

CALL ViewMessage()

%>