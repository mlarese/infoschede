<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="../Tools.asp"-->
<!--#INCLUDE FILE="../Tools4Plugin.asp"-->
<!--#INCLUDE FILE="../ClassConfiguration.asp"-->
<%
dim Config
set Config = new Configuration
'impostazione delle proprieta' di default
Config.AddDefault "movie", ""
Config.AddDefault "id", ""
Config.AddDefault "play", ""
Config.AddDefault "loop", ""
Config.AddDefault "quality", ""
Config.AddDefault "scale", ""
Config.AddDefault "devicefont", ""
Config.AddDefault "bgcolor", ""
Config.AddDefault "allowScriptAccess", ""
Config.AddDefault "align", ""
'caricamento proprieta' specifiche
Config.SetConfigurationString(Session("CONFSTR"))

dim movie
if instr(1, Config("movie"), "http", vbTextCompare)>0 then
	movie = Config("movie")
else
	movie = Config.imageURl + Config("movie")
end if
%>
<!-- URL's used in the movie-->
<!-- text used in the movie-->
<!--web studio & creazioni multimediali next-Aim   -->
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="https://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"<%= write_property("width", SESSION("LAYER_WIDTH")) %><%= write_property("height", SESSION("LAYER_HEIGHT")) %><%= write_property("align", Config("align")) %><%= write_property("id", Config("id")) %>>
<%= write_parameter("movie", movie) %>
<%= write_parameter("play", Config("play")) %>
<%= write_parameter("loop", Config("loop")) %>
<%= write_parameter("quality", Config("quality")) %>
<%= write_parameter("scale", Config("scale")) %>
<%= write_parameter("devicefont", Config("devicefont")) %>
<%= write_parameter("bgcolor", Config("bgcolor")) %>
<%= write_parameter("allowScriptAccess", Config("allowScriptAccess")) %>
<embed <%= write_property("src", movie) %><%= write_property("name", Config("id")) %><%= write_property("play", Config("play")) %><%= write_property("loop", Config("loop")) %><%= write_property("quality", Config("quality")) %><%= write_property("scale", Config("scale")) %><%= write_property("devicefont", Config("devicefont")) %><%= write_property("bgcolor", Config("bgcolor")) %><%= write_property("allowScriptAccess", Config("AllowScriptAccess")) %><%= write_property("width", SESSION("LAYER_WIDTH")) %><%= write_property("height", SESSION("LAYER_HEIGHT")) %><%= write_property("align", Config("align")) %> TYPE="application/x-shockwave-flash" PLUGINSPAGE="https://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash">
</embed>
</object>

<% 
function write_property(name, value)
	if cString(value)<>"" then
		write_property = " " & name & "=""" & value & """"
	else 
		write_property = ""
	end if
end function

function write_parameter(name, value)
	if cString(value)<>"" then
		write_parameter = "<param name=""" & name & """ value=""" & value & """ />"
	else
		write_parameter = ""
	end if
end function
%>