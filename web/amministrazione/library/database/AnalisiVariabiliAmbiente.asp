<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools4DataBase.asp" -->
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<% 
'*****************************************************************************************************************
'verifica dei permessi
CALL VerificaPermessiUtente(true)
'*****************************************************************************************************************
%>
<% dim var, var2 %>
<html>
<head>
	<title>Amministrazione aggiornamenti database</title>
	<link rel="stylesheet" type="text/css" href="../stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0" onload="window.focus();">
<table width="100%" cellspacing="1" cellpadding="0" style="margin-bottom:10px;">
    <caption style="border:0px;">
        <a style="float:right;" href="javascript:close();" class="menu" name="top">CHIUDI</a>
	</caption>
</table>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
    <caption>Variabili di sessione</caption>
	<tr>
        <th>Nome</th>
        <th>Valore</th>
    </tr>
	<%for each var in Session.contents%>
	    <tr>
		    <% if isObject(Session(var)) then %>
			    <td class="content warning"><%= var %></td>
				<td class="content warning">Object (<%= TypeName(Session(var)) %>)</td>
			<% elseif isArray(Session(var)) then %>
				<td class="content ok"><%= var %></td>
				<td class="content ok"><% ListArray(Session(var)) %></td>
			<%else %>
			    <td class="content"><%= var %></td>
			    <td class="content"><%= Session(var) %></td>
			<% end if %>
		</tr>
	<% next %>
</table>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
    <caption>Variabili di applicazione</caption>
	<tr>
        <th>Nome</th>
        <th>Valore</th>
    </tr>
	<%for each var in Application.contents%>
	    <tr>
		    <% if isObject(Application(var)) then %>
			    <td class="content warning"><%= var %></td>
				<td class="content warning">Object (<%= TypeName(Application(var)) %>)</td>
			<% elseif isArray(Application(var)) then %>
				<td class="content ok"><%= var %></td>
				<td class="content ok"><% ListArray(Application(var)) %></td>
			<%else %>
			    <td class="content"><%= var %></td>
			    <td class="content"><%= Application(var) %></td>
			<% end if %>
		</tr>
	<% next %>
</table>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
    <caption>Cookies</caption>
	<tr>
        <th>Cookie</th>
        <th>Valore</th>
    </tr>
	<%for each var in request.Cookies%>
        <tr>
            <td class="content" rowspan="2"><%= var %></td>
            <td class="content"><%= request.Cookies(var) %></td>
        </tr>
        <tr>
            <td>
                <table cellspacing="0" cellpadding="1" border="1">
                    <%for each var2 in request.Cookies(var) %>
                        <tr>
                            <td class="content"><%= var2 %></td>
				            <td class="content"><%= request.Cookies(var)(var2) %></td>
			            </tr>
		            <% next %>
            	</table>
            </td>
        </tr>
    <% next %>
</table>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
    <caption>ServerVariables</caption>
	<tr>
        <th>Nome</th>
        <th>Valore</th>
    </tr>
	<%for each var in request.ServerVariables%>
	    <tr>
		    <td class="content"><%= var %></td>
			<td class="content"><%= request.ServerVariables(var) %></td>
		</tr>
	<% next %>
</table>

<br>
<br>
<br>
</body>
</html>

