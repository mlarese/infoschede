<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1073741824 %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))
%>

<%'--------------------------------------------------------
sezione_testata = "attivazione lingua" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'-----------------------------------------------------
%>
<%
dim conn, connPage, ID_S, ID_D, rs, rsw, sql, default
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsw = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT * FROM tb_pagineSito WHERE id_web = " & request("WEB_ID")
rs.open sql, conn, adOpenStatic, adLockOptimistic, adAsyncFetch
sql = "SELECT * FROM tb_webs WHERE id_webs = " & request("WEB_ID")
rsw.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

if request("conferma") <> "" then
	conn.beginTrans

	if not rsw.eof then
		rsw("lingua_" & request("LINGUA_D")) = 1
		rsw.update
	end if
	
	if request("pagina_copia") <> "" AND request("lingua_s") <> "" then
		do while not(rs.eof)
			ID_S = cInteger(rs("id_pagStage_" & request("LINGUA_D")))
			ID_D = cInteger(rs("id_pagDyn_" & request("LINGUA_D")))
			
			if request("pagina_forza") <> "" or ID_S <= 0 and ID_D <= 0 then
				'aggiorna il titolo solo se la pagina non esiste o se occorre forzare la copia
				rs("nome_ps_" & request("LINGUA_D")) = cString(rs("nome_ps_" & request("lingua_s")))
				rs.update
			end if
			
			if ID_S <= 0 then
				'pagina di lavoro non esistente, la crea
				ID_S = Create_page(conn, rs("nome_ps_" & request("lingua_s")), rs("id_web"), rs("id_pagineSito"), request("LINGUA_D"))
				rs("id_pagStage_" & request("LINGUA_D")) = ID_S
				rs.update
				
				'pagina di lavoro: copia dei dati della pagina e dei layers
				CALL Copy_page(conn, rs("id_pagStage_" & request("lingua_s")), ID_S, false)
				
			elseif request("pagina_forza") <> "" then
				'pagina di lavoro: copia dei dati della pagina e dei layers
				CALL Copy_page(conn, rs("id_pagStage_" & request("lingua_s")), ID_S, false)
			end if

			if ID_D <= 0 then
				'pagina pubblica non esistente, la crea
				ID_D = Create_page(conn, rs("nome_ps_" & request("lingua_s")), rs("id_web"), rs("id_pagineSito"), request("LINGUA_D"))
				CALL conn.execute("UPDATE tb_pagineSito SET id_pagDyn_"&request("LINGUA_D")&"="&ID_D&" WHERE id_pagineSito=" & rs("id_pagineSito"))
				'rs("id_pagDyn_" & request("LINGUA_D")) = ID_D
				'rs.update
				
				'imposta dati di visibilita' della pagina se non vuota.
				'COMMENTATO il 13/10/2011
				'NICOLA: è totalmente inefficente, esegue troppi aggiornamenti sull'indice. Va comunque aggiornato l'indice a fine lavoro di aggiunta della lingua. Ma da n'altra parte.
				'CALL Index_UpdateItem(conn, "tb_pagineSito", rs("id_paginesito"), false)
				
				'pagina pubblica: copia dei dati della pagina e dei layers
				CALL Copy_page(conn, rs("id_pagDyn_" & request("lingua_s")), ID_D, false)
				
			elseif request("pagina_forza") <> "" then
				'pagina di lavoro: copia dei dati della pagina e dei layers
				CALL Copy_page(conn, rs("id_pagDyn_" & request("lingua_s")), ID_D, false)
			end if

		rs.movenext
		loop
	end if
end if
%>


<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption class="border">Attivazione della lingua '<%=cString(request("LINGUA_D"))%>' e copia pagine (facoltativa)</caption>
			<% if request("conferma") = "" then %>
				<tr>
					<td class="content_b" colspan="2">
						Attenzione: l'operazione sar&agrave; irreversibile
					</td>
				</tr>
				<tr>
					<th class="L2" colspan="2">Copia pagine</th>
				</tr>
				<tr>
					<td class="label_no_width" style="width:20%;" nowrap>
						crea pagine copiandole da italiano o inglese:
					</td>
					<td class="content">
						<input class="checkbox" type="checkbox" name="pagina_copia"
							onclick="EnablePickerIfChecked(this, form1.pagina_forza);
									 //EnablePickerIfChecked(this, form1.pagina_pubblica);
									 EnablePickerIfChecked(this, form1.lingua_s);
									 if (!this.checked) {
									 	 form1.pagina_forza.checked = false;
										 //form1.pagina_pubblica.checked = false;
									 	 form1.lingua_s.options[0].selected = true; }">
					</td>
				</tr>
				<tr>
					<td class="label_no_width" style="width:20%;" nowrap>
						forza copia se le pagine esistono gi&agrave;:
					</td>
					<td class="content">
						<input class="checkbox" disabled type="checkbox" name="pagina_forza">
					</td>
				</tr>
				<!--<tr>
					<td class="label_no_width" style="width:20%;" nowrap>
						pubblica pagine:
					</td>
					<td class="content">
						<input class="checkbox" disabled type="checkbox" name="pagina_pubblica">
					</td>
				</tr>-->
				<tr>
					<td class="label_no_width" rowspan="2" style="width:20%;" nowrap>
						seleziona la lingua da cui copiare le pagine:
					</td>
					<td class="content">
						<select name="lingua_s" disabled id="lingua_s">
							<option value=""
								<% if cString(request("lingua_s") = "") then %> selected <% end if %>>
								scegli</option>
							<option value="<%=LINGUA_ITALIANO%>"
								<% if cString(request("lingua_s") = LINGUA_ITALIANO) then %> selected <% end if %>>
								Italiano</option>
							<option value="<%=LINGUA_INGLESE%>"
								<% if cString(request("LINGUA_D") = LINGUA_INGLESE) then %> disabled <% end if %>
								<% if cString(request("lingua_s") = LINGUA_INGLESE) then %> selected <% end if %>>
								Inglese</option>
						</select>
					</td>
				</tr>
				<tr>
					<td class="content" colspan="2">
						<table cellpadding="0" cellspacing="0" class="conferme">
							<tr>
								<td class="content_center">
									<input type="submit" class="button" name="conferma" value="CONFERMA" tabindex="1">
									
								</td>
								<td class="content_center">
									<input type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close()" tabindex="2">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			<% else 'copia a buon fine
				if err.number = 0 then
					conn.CommitTrans%>
					<tr>
						<td colspan="2" class="content_b">
							Lingua attivata.
							<% if request("pagina_copia") <> "" then %>
								Copia pagina eseguita correttamente.
							<% end if %>
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
				<% else
					conn.RollBackTrans %>
					<tr>
						<td colspan="2" class="content_b">
							Errore nella copia: trasferimento annullato<br>
							<%= Err.Number & " - " & Err.Description %>
						</td>
					</tr>
				<% end if
			end if %>
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
set rsw = nothing
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
	PageOnLoad_FocusSet();
//-->
</script>