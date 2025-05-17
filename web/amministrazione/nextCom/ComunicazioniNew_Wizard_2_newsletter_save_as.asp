<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
dim nextWeb_Conn, rs, sql, lingua, value
set nextWeb_Conn = Server.CreateObject("ADODB.Connection")
nextWeb_Conn.open Application("l_conn_ConnectionString"),"",""

if request("salva")<>"" then
	sql = " UPDATE tb_email " & _
		  " SET email_newsletter_tipo_id = "&cIntero(request("tipo_newsletter_id"))&" WHERE email_id = " & cIntero(request("EMAIL_ID"))
	nextWeb_Conn.execute(sql)	%>
	<script language="JavaScript">
		opener.location.reload();
		window.close()
	</script>
<% end if

'--------------------------------------------------------
sezione_testata = "Salva email come newsletter" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->
<%'----------------------------------------------------- 

sql = "SELECT * FROM tb_newsletters "

%>

<div id="content_ridotto">
<form action="" method="post" name="form1" id="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<%= sezione_testata %>
		</caption>
		<tr>
			<th colspan="2">SELEZIONA IL TIPO NEWSLETTER</th>
		</tr>
		<tr>
			<td class="label">tipo:</td>
			<td class="content">
				<% CALL dropDownAdvanced(nextWeb_Conn, sql, "nl_id", "nl_nome_it", "tipo_newsletter_id", cIntero(request("TIPO_ID")), false, " style=""width:100%;""", "< non salvare come newsletter >", "< Nessun tipo newsletter trovato >") %>
			</td>
		</tr>
		<!--
		<tr>
			<td class="content notes">
				La selezione della lingua permette di cambiare la lingua della pagina che verr&agrave; spedita.
			</td>
		</tr>
		-->
		<tr>
			<td class="footer" colspan="2">
				<input type="button" class="button" name="chiudi" value="ANNULLA" onclick="window.close();">
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
</form>
</div>
</body>
</html>

<%
nextWeb_Conn.close
set nextWeb_Conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>