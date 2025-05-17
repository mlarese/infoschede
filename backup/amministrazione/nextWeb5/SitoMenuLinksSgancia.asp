<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<% 
'check dei permessi dell'utente
if NOT index.ChkPrm(prm_menu_accesso, 0) then %>
	<script type="text/javascript">
		window.close()
	</script>
<%
end if

dim conn, rsm, sql
set conn = index.conn
set rsm = Server.CreateObject("ADODB.Recordset")

dim esito
if request("sgancia")<>"" then
	'sgancia menu da eventuale voce
	sql = "UPDATE tb_menu SET m_index_id=NULL WHERE m_id=" & cIntero(request("ID"))
	CALL conn.Execute(sql, , adExecuteNoRecords)
	
	Esito = "OK"
end if

sql = " SELECT *, " + _
	  " (SELECT COUNT(*) FROM tb_menuItem WHERE tb_menuItem.mi_menu_id = tb_menu.m_id) AS N_LINK " + _
	  " FROM tb_menu WHERE m_id=" & cIntero(request("ID"))
rsm.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>

<%'--------------------------------------------------------
sezione_testata = "Gestione siti - menu - sgancia menu" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption class="border">
				Gestione menu &quot;<%= rsm("m_nome_it") %>&quot;
			</caption>
			<%if Esito <>"" then%>
				<tr>
					<td colspan="3" class="content">
						<br>
						<strong>Menu sganciato dalla voce.</strong><br>
						L'inserimento dei link pu&ograve; ora essere fatto manualmente<br>
						<br>
					</td>
				</tr>
				<tr>
					<td colspan="3" class="note">
						Questa finestra si chiuder&agrave; automaticamente tra 5 secondi.
					</td>
				</tr>
				<script language="JavaScript">
					opener.location.reload(true);
					window.setTimeout("close();", 5000);
				</script>
				<tr>
					<td class="footer" colspan="3">
						<input style="width:23%;" type="button" class="button" name="annulla" value="CHIUDI" onclick="window.close();">
					</td>
				</tr>
			<%else%>
				<tr>
					<td class="label_no_width" style="width:23%;">
						voce agganciata:
					</td>
					<td class="content_b">
						<%= index.NomeCompleto(rsm("m_index_id"))  %>
					</td>
				</tr>
				<tr>
					<td colspan="2">
						<table cellpadding="0" cellspacing="0" class="conferme">
							<tr>
								<td class="content_center">
									<input type="submit" class="button" name="sgancia" value="SGANCIA" tabindex="1" id="primo_elemento">
								</td>
								<td class="content_center">
									<input type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close()" tabindex="2">
								</td>
							</tr>
						</table>
					</td>
				</tr>
				
				<tr>
					<td class="footer" colspan="2">
						<input type="button" class="button" name="chiudi" value="CHIUDI" onclick="window.close();">
					</td>
				</tr>
			<% end if %>
		</form>
	</table>
</div>
</body>
</html>

<%
conn.close
set conn = nothing
%>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>