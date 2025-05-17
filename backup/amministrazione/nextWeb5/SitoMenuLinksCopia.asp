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
if request("copia") <> "" then
	'esegue copia dei link
	
	if request("cancella") = "1" then
		'cancella link presenti
		sql = "DELETE FROM tb_menuItem WHERE mi_menu_id = "& cIntero(request("ID"))
		CALL conn.Execute(sql, , adExecuteNoRecords)
	end if
	
	'sgancia menu da eventuale voce
	if cInteger(request("m_index_id"))>0 then
		sql = "UPDATE tb_menu SET m_index_id=NULL WHERE m_id=" & cIntero(request("ID"))
		CALL conn.Execute(sql, , adExecuteNoRecords)
	end if
	
	'esegue copia dei link
	if cInteger(request("copia_menu"))>0 then
		'da altro menu
		CALL Copy_MenuFromMenu(conn, request("ID"), request("copia_menu"))
		
	elseif CIntero(request("copia_index")) > 0 then
		'da voce dell'indice
		CALL Copy_MenuFromIndex(conn, request("ID"), request("copia_index"), request("chk_mi_figli")<>"")
		
	end if
	
	esito = "OK"
end if

sql = " SELECT *, " + _
	  " (SELECT COUNT(*) FROM tb_menuItem WHERE tb_menuItem.mi_menu_id = tb_menu.m_id) AS N_LINK " + _
	  " FROM tb_menu WHERE m_id=" & cIntero(request("ID"))
rsm.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>

<%'--------------------------------------------------------
sezione_testata = "Gestione siti - menu - copia links" %>
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
						<b>Menu copiato correttamente.</b><br>
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
					<td class="label_no_width" rowspan="5" style="width:30%;">
						copia links da:
					</td>
					<td class="content_center" rowspan="2" width="4%;">
						<input type="radio" class="noBorder" name="rad_copia" id="rad_copia_0" onclick="Abilita()" value="0" checked>
					</td>
					<td class="content">altro menu</td>
				</tr>
				<tr>
					<td class="content">
						<% 	sql = " SELECT m_id, m_nome_it FROM tb_menu"& _
								  " WHERE m_id <> "& cIntero(request("ID")) &" AND m_id_webs=" & session("AZ_ID") & " ORDER BY m_nome_it"
							CALL dropDown(conn, sql, "m_id", "m_nome_it", "copia_menu", "", true, "", LINGUA_ITALIANO) %>
					</td>
				</tr>
				<tr>
					<td class="content_center" rowspan="3">
						<input type="radio" class="noBorder" name="rad_copia" id="rad_copia_1" onclick="Abilita()" value="1">
					</td>
					<td class="content">voce dell'indice</td>
				</tr>
				<tr>
					<td class="content">
						<% 	CALL index.WritePicker("", "", "form1", "copia_index", "", Session("AZ_ID"), false, false, "37", true, false) %>
					</td>
				</tr>
				<tr>
					<td class="content">
						<input type="checkbox" name="chk_mi_figli" id="chk_mi_figli" value="1" class="checkbox">
						abilita visualizzazione delle voci figlie
					</td>
				</tr>
				<script type="text/javascript">
					function Abilita() {
						var oCopia = document.getElementById("rad_copia_0");
						EnableIfChecked(oCopia , document.getElementById("copia_menu"));
						
						DisablePickerIfChecked(oCopia, document.getElementById("copia_index"));
						DisableIfChecked(oCopia , document.getElementById("chk_mi_figli"));
					}
					
					Abilita();
				</script>
				
				<% if cInteger(rsm("m_index_id"))>0 then %>
					<input type="hidden" name="m_index_id" value="<%= rsm("m_index_id") %>">
					<tr>
						<td class="note" colspan="3">
							ATTENZIONE:<br>
							Copiando i link il menu verr&agrave; sganciato automaticamente dall'attuale voce dell'indice.
						</td>
					</tr>
				<% elseif cInteger(rsm("N_LINK"))>0 then %>
					<tr>
						<td class="label">cancellare i links presenti?</td>
						<td class="content" colspan="2" style="padding-left: 0px;">
							<input type="radio" class="noBorder" name="cancella" value="1" checked> si
							<input type="radio" class="noBorder" name="cancella" value="0"> no
						</td>
					</tr>
				<% end if %>
				<tr>
					<td class="footer" colspan="3">
						<input style="width:23%;" type="submit" class="button" name="copia" value="COPIA LINKS">
						<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
					</td>
				</tr>
			<% end if %>
		</form>
	</table>
</div>
</body>
</html>

<% 	conn.close
	set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>