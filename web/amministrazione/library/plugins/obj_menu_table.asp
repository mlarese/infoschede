<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="../TOOLS.ASP"-->
<!--#INCLUDE FILE="../TOOLS4Plugin.ASP"-->
<!--#INCLUDE FILE="../ClassConfiguration.asp"-->
<!--#INCLUDE FILE="obj_menu_TOOLS.ASP"-->
<% 
dim Config
set Config = new Configuration
'impostazione delle proprieta' di default
'proprieta' comuni ai due metodi
Config.AddDefault "MenuID", ""
Config.AddDefault "MenuStyleGroup", ""
Config.AddDefault "MenuGroup", ""		'indica il gruppo di appartenenza del menu per mantenere l'elemento selezionato
Config.AddDefault "ShowSpacer", "false"
Config.AddDefault "ShowTitle", "true"

'caricamento proprieta' specifiche
Config.SetConfigurationString(Session("CONFSTR"))

dim href
dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.Recordset")
conn.open Application("l_conn_ConnectionString"), "", ""

'variabili per la gestione della selezione dei menu correnti
dim CssClass, CssId, MenuItemID, Gruppo, Title
Gruppo = Config("MenuGroup")
MenuItemID = getMenuItemSelected(conn, rs, Config, Gruppo)

sql = "SELECT * From tb_menuItem INNER JOIN tb_links ON tb_menuItem.id_link=tb_links.id " + _
	  "WHERE (tb_links.id = " & cInteger(Config("MenuID")) & ") AND attivo_mi ORDER BY ordine_menuItem"
rs.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch %>
<% if Config("MenuStyleGroup")<>"" then %>
	<div id="<%= Config("MenuStyleGroup") %>">
<% end if
	if instr(1, Config("ShowTitle"), "false", vbTextCompare)<1 then%>
		<h1><%= CBL(rs, "nomelink") %></h1>
	<% end if %>
	<table cellpadding="<%= cIntero(Config("cellpadding")) %>" cellspacing="<%= cIntero(Config("cellspacing")) %>">
		<tr>
			<%while not rs.eof
				if CInteger(rs("id_MenuItem")) = MenuItemID then
					CssClass = "class=""selected"" "
				else
					CssClass = ""
				end if
				if rs.AbsolutePosition = 1 then
					CssId = "id=""first"" "
				elseif rs.AbsolutePosition = rs.recordcount then
					CssId = "id=""last"" "
				else
					CssId = ""
				end if
				title = cString(CBL(rs, "tag_title"))
				if title = "" then
					title = CBL(rs, "titolo_menuItem")
				end if%>
				<td <%= CssClass %> <%= CssId %>>
					<a <%= CssClass %> href="<%= getHREF(Config, rs, gruppo) %>" <% if rs("link_target")<>"" then %> target="<%= rs("link_target") %>"<% end if %> title="<%= title %>" lang="<%= Config.lingua %>">
					   	<% if cString(CBL(rs, "image_menuItem"))<>"" then %>
							<img border="0" src="<%= Config.ImageUrl & CBL(rs, "image_menuItem") %>" alt="<%= title %>"></a>
						  <% else %>
						  	<%= CBL(rs, "titolo_menuItem") %></a>
						<% end if %>
				</td>
				<%rs.movenext
				if not rs.eof AND _
					instr(1, Config("ShowSpacer"), "true", vbTextCompare)>0 then %>
						<td class="spacer">&nbsp;</td>
				<%end if
			wend %>
		</tr>
	</table>
<% if Config("MenuStyleGroup")<>"" then %>
	</div>
<% end if %>
<% rs.close
conn.close
set Config = nothing
set rs = nothing
set conn = nothing
%>