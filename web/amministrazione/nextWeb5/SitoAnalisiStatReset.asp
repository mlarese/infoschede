<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<!--#INCLUDE FILE="SitoAnalisiStat_TOOLS.asp" -->
<%
dim conn, rs, sql

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")

if request("conferma")<>"" AND cIntero(request("WEB_ID"))>0 then
	'aggiunge allo storico la situazione attuale
	conn.beginTrans
	
	CALL StatisticheArchiviaAzzera(conn, request("WEB_ID"))
		
	conn.commitTrans
end if

sql = " SELECT * FROM tb_webs WHERE id_webs=" & cIntero(request("WEB_ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adAsyncFetch
%>

<%'--------------------------------------------------------
sezione_testata = "Statistiche di accesso - archiviazione ed azzeramento" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>
				Archiviazione statistiche ed azzeramento contatori
			</caption>
			<% if request("conferma") = "" then %>
				<tr><th colspan="2">DATI DEL SITO</th></tr>
				<tr>
					<td class="label_no_width" style="width:28%;">sito:</td>
					<td class="content"><%= rs("nome_webs") %></td>
				</tr>
				<tr>
					<td class="label_no_width">ultimo azzeramento:</td>
					<td class="content"><%= DateTimeIta(rs("contRes")) %></td>
				</tr>
				<tr><th colspan="2">RICHIESTA DI CONFERMA</th></tr>
				<tr>
					<td colspan="2" class="note">
						N.B.: Le operazioni di archiviazione ed azzeramento avverranno sia sulle pagine che sull'indice.
					</td>
				</tr>
				<tr>
					<td class="content_b" colspan="2">
						Attenzione: l'operazione sar&agrave; irreversibile.
					</td>
				</tr>
				<tr>
					<td class="content" colspan="2">
						<table cellpadding="0" cellspacing="0" class="conferme">
							<tr>
								<td class="content_center">
									<input type="submit" class="button" name="conferma" value="PROCEDI" tabindex="1" id="primo_elemento">
								</td>
								<td class="content_center">
									<input type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close()" tabindex="2">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			<% else 'andato a buon fine %>
				<tr>
					<th>ESITO OPERAZIONI</th>
				</tr>
				<tr>
					<td colspan="2" class="content_b">
						Archiviazione statistiche ed azzeramento contatori eseguite correttamente.
					</td>
				</tr>
				<tr>
					<td colspan="2" class="note">
						Questa finestra si chiuder&agrave; automaticamente tra 5 secondi.
					</td>
				</tr>
				<script language="JavaScript">
					opener.location.reload(true);
					window.setTimeout("close();", 5000);
				</script>
			<% end if %>
			<tr>
				<td class="footer" colspan="2">
					<input type="button" class="button" name="chiudi" value="CHIUDI" onclick="window.close();">
				</td>
			</tr>
		</form>
	</table>
</div>
</body>
</html>

<%
conn.close
set rs = nothing
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
	PageOnLoad_FocusSet();
//-->
</script>
