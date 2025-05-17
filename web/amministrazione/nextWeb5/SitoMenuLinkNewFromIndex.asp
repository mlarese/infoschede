<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_menu_accesso, 0))

dim conn, rsm, rsi, sql, lingua
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	
	set rsi = Server.CreateObject("ADODB.Recordset")
	set rsm = Server.CreateObject("ADODB.Recordset")
	
	if cIntero(request("tfn_mi_index_id"))>0 then
		
		sql = "SELECT * FROM v_indice WHERE idx_id=" & cIntero(request("tfn_mi_index_id"))
		rsi.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
		
		if not rsi.eof then
			
			sql = "SELECT * FROM tb_menuItem WHERE mi_menu_id=" & cIntero(request("MENU"))
			rsm.open sql, conn, adOpenKeySet, adLockOptimistic
			rsm.addnew
			rsm("mi_menu_id") = request("MENU")
			rsm("mi_index_id") = request("tfn_mi_index_id")
			rsm("mi_attivo") = true
			rsm("mi_figli") = false
			rsm("mi_ordine") = rsi("co_ordine")
			for each lingua in Application("LINGUE")
				rsm("mi_titolo_" & lingua) = rsi("co_titolo_" & lingua)
				rsm("mi_tag_title_" & lingua) = rsi("idx_alt_" & lingua)
			next
			
			rsm.update
			
			response.redirect "SitoMenuLinkMod.asp?MENU=" & request("MENU") & "&ID=" & rsm("mi_id")
			
			rsm.close
			
		end if
		
		rsi.close
		
	else
		Session("ERRORE") = "Scegli la voce dell'indice."
	end if
	
	set rsi = nothing
	set rsm = nothing
	
end if

'--------------------------------------------------------
sezione_testata = "Gestione siti - menu - copia nuovo link da voce dell'indice" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Copia nuovo link da voce dell'indice</caption>
		<tr><th colspan="4">VOCE DELL'INDICE</th></tr>
		<tr>
			<td class="label" rowspan="2">scegli la voce</td>
			<td class="content">
				<% 	CALL index.WritePicker("", "", "form1", "tfn_mi_index_id", request("tfn_mi_index_id"), Session("AZ_ID"), false, false, "77", false, false) %>
			</td>
		</tr>
		<tr>
			<td class="note">
				Seleziona la voce dell'indice da cui copiare i dati per creare il nuovo link di menu.
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				<input style="width:23%;" type="submit" class="button" name="importa" value="INSERISCI LINK">
			</td>
		</tr>
	</table>
</div>
</body>
</html>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>
<%
conn.Close
set conn = nothing
%>