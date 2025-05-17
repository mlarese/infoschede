<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="../TOOLS.ASP"-->
<!--#INCLUDE FILE="../TOOLS4plugin.ASP"-->
<!--#INCLUDE FILE="../ClassConfiguration.asp"-->
<% 
dim Config
set Config = new Configuration
'impostazione delle proprieta' di default
Config.AddDefault "width", "425"
Config.AddDefault "height", "344"
Config.AddDefault "allowFullScreen", "true"
Config.AddDefault "allowscriptaccess", "always"
Config.AddDefault "url", ""
'caricamento proprieta' specifiche
Config.SetConfigurationString(Session("CONFSTR"))

%>
<div>
<% dim width,height,url
	width = Config("width")
	height = Config("height")
	url = Config("url")
	if url<>"" then
		url = "http://www.youtube.com/v/"&url&"&hl=it&fs=1"
 %>
<object width="<%= width %>" height="<%= height %>">
	<param name="movie" value="<%= url %>"></param>
	<param name="allowFullScreen" value="<%=Config("allowFullScreen") %>"></param>
	<param name="allowscriptaccess" value="<%=Config("allowscriptaccess") %>"></param>
	<embed src="<%= url %>" 
			type="application/x-shockwave-flash" 
			allowscriptaccess="<%=Config("allowscriptaccess") %>" 
			allowfullscreen="<%=Config("allowFullScreen") %>" 
			width="<%= width %>" 
			height="<%= height %>"
	>
	</embed>
</object></div>
<% end if %>
<%set Config = nothing%>