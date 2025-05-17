<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="../TOOLS.ASP"-->
<!--#INCLUDE FILE="../TOOLS4plugin.ASP"-->
<!--#INCLUDE FILE="../ClassConfiguration.asp"-->
<% 
dim Config
set Config = new Configuration
'impostazione delle proprieta' di default
Config.AddDefault "color", ""
Config.AddDefault "image", ""
Config.AddDefault "position", ""
Config.AddDefault "repeat", ""
Config.AddDefault "border", ""
Config.AddDefault "ShowBack", "false"
Config.AddDefault "id", ""
'caricamento proprieta' specifiche
Config.SetConfigurationString(Session("CONFSTR"))

dim style
style = ""
style = style + IIF(Config("color")<>"", "background-color:" & Config("color") & ";", "")
style = style + IIF(Config("image")<>"", "background-image:url(" & Config.ImageUrl & Config("image") & ");", "")
style = style + IIF(Config("position")<>"", "background-position:" & Config("position") & ";", "")
style = style + IIF(Config("repeat")<>"", "background-repeat:" & Config("repeat") & ";", "")
style = style + IIF(Config("border")<>"", "border:" & Config("border") & ";", "")
%>
<div <%= IIF(Config("id")<>"", "id=""" & Config("id") & """ ","") %> style="font-size:1px; width:<%= Session("LAYER_WIDTH") %>px; height:<%= Session("LAYER_HEIGHT") %>px; <%= style %>">
	&nbsp;
	<% if instr(1, Config("ShowBack"), "true", vbTextCompare)>0 then 
		if request.ServerVariables("HTTP_REFERER")<>"" then
			CALL WriteBackURL(request.ServerVariables("HTTP_REFERER"))
		else
			CALL WriteBackLink()
		end if
	 end if %>
</div>

<%set Config = nothing%>