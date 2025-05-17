
<%CALL CheckAutentication(Session("WEB_ADMIN")<>"" OR session("WEB_POWER") <> "" OR session("WEB_USER") <> "")
%>

<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/class_testata.asp" -->
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<!DOCTYPE html>
<html>
	<head>
		<title><%= Session("NOME_APPLICAZIONE") %></title>
		<link rel="stylesheet" type="text/css" href="../library/stili.css">
		<SCRIPT LANGUAGE="javascript"  src="../library/utils.js" type="text/javascript"></SCRIPT>
		<meta http-equiv=Content-Type content="text/html; charset=utf-8">
		<meta name="robots" content="noindex,nofollow" />
		<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
	</head>
<body>
<!-- barra alta -->
<div id="Layer0" style="position:absolute; left:0px; top:0px; width:740px; height:400px; z-index:0">
<table width="740" cellspacing="0" cellpadding="0" border="0">
	<caption class="menu">
		<table width="100%" cellspacing="0" cellpadding="0" border="0">
	  		<tr>
	  			<td width="10">&nbsp;</td>
				<td width="85">
					<a href="Siti.asp" class="menu" title="gestione dei siti amministrati dal sistema" <%= ACTIVE_STATUS %>>GESTIONE SITI</a>
				</td>
				<% if index.ChkPrm(prm_indice_accesso, 0) then %>
					<td width="100">
						<a href="IndexGenerale.asp" class="menu" title="gestione dell'indice generale dei contenuti" <%= ACTIVE_STATUS %>>INDICE GENERALE</a>
					</td>
				<%end if%>
				<td align="right">
					<% if Session("AZ_ID")<>"" then %>
						<table border="0" cellspacing="0" cellpadding="0" style="padding-left: 5px; padding-right: 5px; width:98%; text-align:center;">
							<tr>
								<td>
                                    <span style="width:100%; height:12px; overflow:hidden; text-align:right;">
									<% 	if index.ChkPrm(prm_siti_gestione, 0) then %>
										<a href="SitoMod.asp?ID=<%= Session("AZ_ID") %>" class="menu" title="dati principali del sito <%= Session("NOME_WEBS") %>" <%= ACTIVE_STATUS %>><%= Session("NOME_WEBS") %></a>
									<% 	else %>
										<%= Session("NOME_WEBS") %>
									<% 	end if %>
                                    </span>
								</td>
                                
								<% 	if index.ChkPrm(prm_menu_accesso, 0) then %>
								<td style="border-left:1px solid gray;"><a href="SitoMenu.asp" class="menu" title="gestione dei menu del sito  <%= Session("NOME_WEBS") %>" <%= ACTIVE_STATUS %>>MENU</a></td>
								<% 	end if %>
								<td style="border-left:1px solid gray;"><a href="SitoFileManager.asp" class="menu" title="gestione di tutti i files del sito <%= Session("NOME_WEBS") %>" <%= ACTIVE_STATUS %>>FILES</a></td>
                                <%if IsNextAim() then 
                                 	if index.ChkPrm(prm_plugin_accesso, 0) then %>
                                        <td style="border-left:1px solid gray;"><a href="SitoPlugin.asp" class="menu" title="Gestione delle proprieta' dei plugin dinamici utilizzati nel sito <%= Session("NOME_WEBS") %>" <%= ACTIVE_STATUS %>>PLUGIN</a></td>
                                    <%end if
                                end if %>                                
								<% 	if index.ChkPrm(prm_template_accesso, 0) then %>
									<td style="border-left:1px solid gray;"><a href="SitoTemplate.asp" class="menu" title="gestione template di base per il sito <%= Session("NOME_WEBS") %>" <%= ACTIVE_STATUS %>>TEMPLATES</a></td>
								<% 	end if %>
									<td style="border-left:1px solid gray;"><a href="SitoPagine.asp" class="menu" title="gestione delle pagine che compongono il sito <%= Session("NOME_WEBS") %>" <%= ACTIVE_STATUS %>>PAGINE</a></td>
								<% if index.ChkPrm(prm_stili_accesso, 0) then %>
									<td style="border-left:1px solid gray;"><a href="SitoStili.asp" class="menu" title="gestione degli stili del sito  <%= Session("NOME_WEBS") %>" <%= ACTIVE_STATUS %>>STILI</a></td>
								<% 	end if %>
							</tr>
						</table>
					<% else %>
						&nbsp;
					<% end if %>
				</td>
				<td width="120" class="logout" style="<% if Session("AZ_ID")<>"" then %>border-left:1px solid gray;<% end if %>"><% CALL WriteLogoutLink(Session("NOME_APPLICAZIONE")) %></td>
	  		</tr>
		</table>
	</caption>
    <% CALL WriteChiusuraIntestazione("barra_nextweb.jpg") %>