<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_Memo2.asp" -->
<%
'--------------------------------------------------------
sezione_testata = "Spedizione avvisi via mail" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->

<%'----------------------------------------------------- 
dim conn
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
conn.BeginTrans
if cIntero(request("ID"))>0 then
	CALL SendAvvisoImpegno(conn,cIntero(request("ID")),cIntero(Session("ID_PAGINA_AVVISO")))
end if
conn.commitTrans
%>
<div id="content_ridotto">
<form action="" method="post" id="ricerca" name="ricerca">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption style="border-bottom:1px solid #919191;">
		Avvisi via e-mail
	</caption>
	<tr>
		<td class="content_center">
			<br><b>
			<% if cIntero(request("ID"))>0 then %>
				Gli avvisi sono stati spediti agli utenti interessati.
			<% else %>
				ERRORE! Avvisi non spediti!
			<% end if %>
			</b><br>&nbsp;
		</td>
	</tr>
		<tr>
			<td bgcolor="#e6e6e6" style="text-align:right; padding-left:10px; border-top:1px solid #919191">
				<a href="#" onclick="window.close();" class="button" title="Chiudi la finestra" <%= ACTIVE_STATUS %>>
					CHIUDI
				</a>
			</td>
		</tr>
</table>

</body>
</html>