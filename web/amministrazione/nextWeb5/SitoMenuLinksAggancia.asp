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
if request("aggancia")<>"" then
	if cInteger(request("aggancia_index"))=0 then
		Session("ERRORE") = "Voce dell'indice non selezionata: impossibile agganciare il menu."
	else
		if cInteger(request("N_link"))>0 then
			'cancella link precedenti
			sql = "DELETE FROM tb_menuItem WHERE mi_menu_id = "& cIntero(request("ID"))
			CALL conn.Execute(sql, , adExecuteNoRecords)
		end if
		
		sql = "UPDATE tb_menu SET m_index_id=" & cIntero(request("aggancia_index")) & " WHERE m_id=" & cIntero(request("ID"))
		CALL conn.Execute(sql, , adExecuteNoRecords)
		Esito = "OK"
	end if
end if

sql = " SELECT *, " + _
	  " (SELECT COUNT(*) FROM tb_menuItem WHERE tb_menuItem.mi_menu_id = tb_menu.m_id) AS N_LINK " + _
	  " FROM tb_menu WHERE m_id=" & cIntero(request("ID"))
rsm.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>

<%'--------------------------------------------------------
sezione_testata = "Gestione siti - menu - aggamcia menu" %>
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
						<b>Menu agganciato correttamente alla voce.</b><br>
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
						aggancia alla voce:
					</td>
					<td class="content">
						<% 	CALL index.WritePicker("", "", "form1", "aggancia_index", rsm("m_index_id"), Session("AZ_ID"), false, false, "47", false, true) %>
					</td>
				</tr>
				<% if cInteger(rsm("N_LINK"))>0 then %>
					<input type="hidden" name="n_link" value="<%= rsm("N_LINK") %>">
					<tr>
						<td class="note" colspan="3">
							ATTENZIONE:<br>
							Agganciando il menu ad una voce dell'indice i link attualmente presenti verranno cancellati.
						</td>
					</tr>
				<% end if %>
				<tr>
					<td class="footer" colspan="3">
						<input style="width:30%;" type="submit" class="button" name="aggancia" value="AGGANCIA ALLA VOCE">
						<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
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