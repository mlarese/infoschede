<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->

<%
dim conn, sql, rs, rsd, rsi, voucher_page, url
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.RecordSet")
set rsi = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT lingua, IDElencoIndirizzi, codiceInserimento " & _
	  " FROM tb_Utenti INNER JOIN tb_Indirizzario ON tb_Utenti.ut_NextCom_ID = tb_Indirizzario.IDElencoIndirizzi" & _
	  " WHERE IDElencoIndirizzi=" & cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText



url = GetPageSiteUrl(conn, cIntero(Session("PAGINA_AVVISO_ABILITAZIONE_UTENTE")), rs("lingua")) & "&ID_ADMIN="&Session("ID_ADMIN")&"&IDCNT=" & rs("IDElencoIndirizzi") & "&KEY=" & rs("codiceInserimento") & "&HTML_FOR_EMAIL=1&SEND=1"


'--------------------------------------------------------
sezione_testata = "Gestione prenotazioni" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'------------------------------------------------------ 
%>
<div id="content_ridotto">
<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Gestione utenti</caption>

		<% if url<>"" then %>
			<tr>
				<td colspan="2" class="content notes">
					Se la pagina di anteprima dell'avviso viene visualizzata correttamente l'operazione richiesta &egrave; andata a buon fine.
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<iframe src="<%= URL %>" name="esito_spedizione" width="100%" height="400" id="esito_spedizione"></iframe>
				</td>
			</tr>
			<script language="JavaScript" type="text/javascript">
				window.setTimeout("opener.location.reload(true);", 3000);
			</script>
			<!-- <%=url%> -->
		<% else %>
			<tr>
				<td colspan="2" class="content_b content_center">
					<br/>
					IMPOSSIBILE SPEDIRE AVVISO
					<br/>
					<br/>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td colspan="2" class="footer">
				<input type="button" onclick="window.close();" class="button" name="chiudi" value="CHIUDI">
			</td>
		</tr>
	</table>
	&nbsp;
</form>
</div>
</body>
</html>