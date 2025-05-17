<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="../tools.asp"-->
<!--#INCLUDE FILE="../tools4plugin.asp"-->
<!--#INCLUDE FILE="../ClassConfiguration.asp"-->
<% 
'configuarazione proprietà oggetto
dim Config
set Config = new Configuration
'impostazione delle proprietà di default
Config.AddDefault "display", ""
Config.AddDefault "LabelLenght", ""
Config.AddDefault "PageToGo", ""
Config.AddDefault "label_IT", ""
Config.AddDefault "label_EN", ""
Config.AddDefault "label_FR", ""
Config.AddDefault "label_DE", ""
Config.AddDefault "label_ES", ""
' Bandiere grafiche 
Config.AddDefault "flag_bmp_file", ""
Config.AddDefault "flag_file_ext", ".jpg"

'caricamento proprieta' specifiche
Config.SetConfigurationString(Session("CONFSTR"))

'parametri speciali: corregge il problema che in alcuni siti LingueDisabilitate &egrave; usato con lo stesso significato
'di lingue nascoste.
if not Config.Exists("LingueNascoste") then
	Config.AddDefault "LingueNascoste", Config("LingueDisabilitate")
	Config.AddDefault "LingueDisabilitate", ""
end if

'simulo il post
dim var, html_hidden, i, pagina, querystring
if request.serverVariables("REQUEST_METHOD") = "POST" then
	'scrivo in una var un elenco di hidden da postare
	for each var in request.form
		html_hidden = html_hidden &"<input type='hidden' name='"& var &"' value='"& request.form(var) &"'>"
	next
end if
'simulo il querystring
for each var in request.querystring
	if uCase(var) <> "PAGINA" AND uCase(var) <> "LINGUA" then
		querystring = querystring & "&"& var &"="& request.querystring(var)
	end if
next

dim flags,fgif ' File per le bandiere
if Config("flag_bmp_file")<>"" then
		flags = array("_it","_en","_fr","_es","_de")
end if

dim conn, rs, sql, label, PageToGo
set conn = server.createobject("ADODB.connection")
set rs = server.CreateObject("ADODB.Recordset")
conn.open Application("L_conn_ConnectionString")
PageToGo = cInteger(Config("PageToGo"))
pagina = cIntero(Session("CURRENT_PAGINA"))
sql = " SELECT * FROM tb_pagineSito WHERE " & _
	  " id_pagDyn_it=" & pagina & _
	  " OR id_pagDyn_en=" & pagina & _
	  " OR id_pagDyn_fr=" & pagina & _
	  " OR id_pagDyn_es=" & pagina & _
	  " OR id_pagDyn_de=" & pagina & _
	  " OR id_pagStage_it=" & pagina & _
	  " OR id_pagStage_en=" & pagina & _
	  " OR id_pagStage_fr=" & pagina & _
	  " OR id_pagStage_es=" & pagina & _
	  " OR id_pagStage_de=" & pagina
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText %>

<table cellspacing="0" cellpadding="0" class="cambio_lingua">
<%
label = CBL(Config, "label")
if label <> "" then	%>
	<tr>
		<td class="cambio_lingua_caption"><%= label %></td>
	<%	if UCase(config("display")) <> "INLINE" then %>
		</tr>
	<%	end if
elseif UCase(config("display")) = "INLINE" and label = "" then %>
	<tr>
<% end if

for i = lBound(application("LINGUE")) to uBound(application("LINGUE"))
	if instr(1, Config("LingueNascoste"), application("LINGUE")(i), vbTextCompare)<1 then
		if rs.eof then
			pagina = "default.asp?"
		elseif PageToGo>0 then
			pagina = "default.asp?PS=" & Config("PageToGo") & "&"
		else
			pagina = "dynalay.asp?PAGINA=" & rs("id_pagDyn_"& application("LINGUE")(i)) & "&"
		end if
		pagina = pagina & "LINGUA=" & application("LINGUE")(i)
	
		if UCase(config("display")) <> "INLINE" then
			response.write "<tr>"
		end if %>
			<td id="<%= application("LINGUE")(i) %>">
				<form action="<%= pagina %><%= querystring %>" method="post" name="frmLingua_<%= application("LINGUE")(i) %>" id="frmLingua_<%= application("LINGUE")(i) %>">
					<%= html_hidden %>
					<% if instr(1, Config("LingueDisabilitate"), application("LINGUE")(i), vbTextCompare)<1 then
						if request.serverVariables("REQUEST_METHOD") = "POST" then %>
							<a href="#" onclick="javascript:document.frmLingua_<%= application("LINGUE")(i) %>.submit()"
						<% else %>	
							<a href="<%= pagina %><%= querystring %>"
						<% end if
					else %>
						<a
					<%end if%>
						title="<%= LINGUE_NAMES(i) %>" 
						id="<%= application("LINGUE")(i) %>"
					<%= IIF(config.lingua = application("LINGUE")(i), "class=""selected"" ", "") %>>
					<% if Config("flag_bmp_file")<>"" then %>
						<% fgif = Config("flag_bmp_file") & flags(i) & Config("flag_file_ext") %>
						<img border="0" src="<%= config.imageURL & fgif  %>" alt="<%=LINGUE_NAMES(i) %>" id="fl_<%=LINGUE_NAMES(i) %>"  /> 
					<% else %>
						<%= IIF(cInteger(Config("LabelLenght"))>0, left(LINGUE_NAMES(i), cInteger(Config("LabelLenght"))),LINGUE_NAMES(i)) %>
					<% end if %>
					</a>
				</form>
			</td>
		<%if UCase(config("display")) <> "INLINE" then
			response.write "</tr>"
		end if
	end if
next

if UCase(config("display")) = "INLINE" then %>
	</tr>
<% end if

rs.close
set rs = nothing
conn.close
set conn = nothing
%>
</table>