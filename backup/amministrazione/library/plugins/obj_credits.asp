<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="../Tools.asp"-->
<!--#INCLUDE FILE="../ClassConfiguration.asp"-->
<%
	'configuarazione proprieta oggetto
	dim Config
	set Config = new Configuration
	'impostazione delle proprieta' di default
	if cInt("0" & Session("VERSION")) <> 4 then
		Config.AddDefault "coloreLink", "black"
		Config.AddDefault "coloreLinkHover", "blue"
		Config.AddDefault "coloreTesto", "black"
		Config.AddDefault "allineaTesto", "left"
		%>
		<style type="text/css">
		A.cre_link{
			color: <%= Config("coloreLink") %>;
			font-size:9px;
		}
		
		A.cre_link:hover{
			color: <%= Config("coloreLinkHover") %>;
		}
		SPAN.cre_text {
			font-family: Arial, Helvetica, sans-serif;
			color: <%= Config("coloreTesto") %>;
			text-align: <%= Config("allineaTesto") %>
			font-size:9px;
		}
		</style>
		<span class="cre_text">
			powered by <a class="cre_link" target="_blank" href="http://www.next-aim.com" title="NEXT-AIM - eBusiness, eCommerce, software on demand, siti internet, applicazioni web">NEXT-AIM</a>
		</span>
	<% else %>
		<div id="NEXTAIM_credits">
			<a target="_blank" href="http://www.next-aim.com" title="NEXT-AIM - eBusiness, eCommerce, software on demand, siti internet, applicazioni web">powered by NEXT-AIM</a>
		</div>
	<% end if %>