<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%'--------------------------------------------------------
if request("template") <> "" then
	sezione_testata = "Gestione siti - templates - "& lcase(request("AZIONE"))
elseif request("nextmail") <> "" then
	sezione_testata = "nuova next-email - "& lcase(request("AZIONE"))
else
	sezione_testata = "Gestione siti - indice delle pagine - "& lcase(request("AZIONE")) & " la pagina"
end if %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>
<%
dim conn, ID_S, ID_D, rs, i, lingua, sql, default, pubblica, controllo, is_template

ID_S = cInteger(request("ID_S"))
ID_D = cInteger(request("ID_D"))
if request("sorgente") = "template" then
	ID_S = cInteger(request("ID_S_template"))
end if
pubblica = request("nextmail") <> ""
is_template = request("template") <> ""
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")


'check dei permessi dell'utente
sql = QryPagineSito("*", ID_D)
rs.open sql, conn, adOpenStatic, adLockReadOnly	'seleziono la paginaSito di destinazione

if not is_template then
	if DB_Type(conn) = DB_ACCESS then
		controllo = cBoolean(ID_D = 0 OR rs("id_pagDyn_it") = ID_D OR rs("id_pagDyn_en") = ID_D OR rs("id_pagDyn_fr") = ID_D _
					OR rs("id_pagDyn_es") = ID_D OR rs("id_pagDyn_de") = ID_D, false)
	else
	'response.write sql
		controllo = cBoolean(ID_D = 0 OR rs("id_pagDyn_it") = ID_D OR rs("id_pagDyn_en") = ID_D OR rs("id_pagDyn_fr") = ID_D _
					OR rs("id_pagDyn_es") = ID_D OR rs("id_pagDyn_de") = ID_D OR rs("id_pagDyn_ru") = ID_D _
					OR rs("id_pagDyn_cn") = ID_D OR rs("id_pagDyn_pt") = ID_D, false)
	end if
end if 

if rs.eof then		'template
	if NOT index.ChkPrm(prm_template_accesso, 0) then
		rs.close
		conn.close
		set conn = nothing %>
		<script language="JavaScript">
			window.close()
		</script>
	<% end if
elseif controllo then 'se sto pubblicando, controllo piu sicuro di request...
	if NOT ChkPrmPages(ID_D) then
		rs.close
		conn.close
		set conn = nothing %>
		<script language="JavaScript">
			window.close()
		</script>
	<% end if
else
	if NOT ChkPrmPages(ID_D) then
		rs.close
		conn.close
		set conn = nothing %>
		<script language="JavaScript">
			window.close()
		</script>
	<% end if
end if

if request("conferma") = "" AND NOT rs.eof then
	if ID_D = rs("id_pagDyn_en") OR ID_D = rs("id_pagStage_en") then
		if pubblica then
			default = rs("id_pagDyn_it")
		else
			default = rs("id_pagStage_it")
		end if
	elseif ID_D <> rs("id_pagDyn_it") AND ID_D <> rs("id_pagStage_it") then
		if pubblica then
			default = rs("id_pagDyn_en")
		else
			default = rs("id_pagStage_en")
		end if
	end if
else
	default = ID_S
end if
rs.close

if request("conferma") <> "" then
	conn.beginTrans
	
	'controllo esistenza pagina pubblica se in pubblicazione
	if request("AZIONE") = "PUBBLICA" then
		sql = " SELECT * FROM tb_PagineSito WHERE id_pagStage_" & request("LINGUA") & "=" & ID_S
		rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		if ID_D<=0 then
			'pagina pubblica non esistente, la crea
			ID_D = Create_page(conn, rs("nome_ps_" & request("LINGUA")), rs("id_web"), rs("id_pagineSito"), request("LINGUA"))
			rs("id_pagDyn_" & request("LINGUA")) = ID_D
			rs.update
			
			'imposta dati di visibilita' della pagina se non vuota.
			CALL Index_UpdateItem(conn, "tb_pagineSito", rs("id_paginesito"), false)
		end if
		
		rs.close
		set rs = nothing
	end if
	
	if ID_S = 0 then		'ho scelto la pagina da cui copiare
		CALL Copy_page(conn, request("pagina"), ID_D, false)
	else	'pagina sorgente gia impostata (template)
		'copia dei dati della pagina e dei layers
		CALL Copy_page(conn, ID_S, ID_D, false)
	end if
	
	'se template copio anche i layer del template della pagina
	if CIntero(request("template")) > 0 AND CIntero(request("pagina")) > 0 then
		dim idTemplate
		idTemplate = CIntero(GetValueList(conn, NULL, "SELECT id_template FROM tb_pages WHERE id_page = "& request("pagina")))
		if idTemplate > 0 then
			CALL Copy_page(conn, idTemplate, ID_D, true)
		end if
	end if
	
	if Session("ERRORE") <>"" then
		response.redirect request.ServerVariables("HTTP_REFERER")
	end if
end if
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption class="border">
				<%Select Case request("AZIONE")
					case "RESET" %>
						Ripristino della pagina da ultima versione pubblicata
					<% case "COPIA" %>
						Copia da
					<% case "PUBBLICA" %>
						Pubblicazione della pagina di lavoro
				<% end select %>
			</caption>
			<% if request("conferma") = "" then %>
				<tr>
					<td class="content_b" colspan="2">
						Attenzione: l'operazione sar&agrave; irreversibile
					</td>
				</tr>
				<% if request("AZIONE") <> "COPIA" then %>
					<tr>
						<td class="content" colspan="2">
							<table cellpadding="0" cellspacing="0" class="conferme">
								<tr>
									<td class="content_center">
										<input type="submit" class="button" name="conferma" value="PUBBLICA" tabindex="1" id="primo_elemento">
									</td>
									<td class="content_center">
										<input type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close()" tabindex="2">
									</td>
								</tr>
							</table>
						</td>
					</tr>
				<% else%>
					<tr>
						<th class="L2" colspan="2">SELEZIONA LA PAGINA SORGENTE DA CUI COPIARE</th>
					</tr>
					<% if request("template") <> "" OR request("nextmail")<>"" then %>
						<tr>
							<td class="label_no_width" nowrap>
								<input class="noborder" type="radio" name="sorgente" value="template" checked onclick="DisablePickerIfChecked(this, form1.pagina);EnablePickerIfChecked(this, form1.ID_S_template)">
								template:
							</td>
							<td class="content">
								<% sql = QryElencoTemplate(" AND id_page<>"& cIntero(request("template")), false)
								CALL dropDown_WithViewer("VISUALIZZA", "dynalay.asp?PAGINA=", _
										   				 "form1", IIF(request("nextmail")<>"", "", 250), conn, sql, "id_page", "name", "ID_S_template", "", true)%>
							</td>
						</tr>
					<% end if %>
					<tr>
						<td class="label_no_width" style="width:20%;" nowrap>
							<% 	if request("template") <> "" OR request("nextmail")<>"" then %>
								<input class="noborder" type="radio" name="sorgente" value="pagina" onclick="DisablePickerIfChecked(this, form1.ID_S_template);EnablePickerIfChecked(this, form1.pagina)">
							<% 	end if %>
							pagina:
						</td>
						<td class="content">
						<% 	CALL DropDownPagesAdvanced(conn, "form1", "100%", IIF(request("nextmail")<>"" OR cBoolean(Session("COPIA_PAGINE_TRA_SITI"),false), 0,  Session("AZ_ID")), "pagina", default, false, true, pubblica, ID_D)
							if request("template") <> "" OR request("nextmail")<>"" then %>
								<script language="JavaScript">
									DisablePicker(form1.pagina, true)
								</script>
							<% 	end if %>
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
				<% end if 'fine azione <> COPIA
			else 'copia a buon fine
				if err.number=0 then
					conn.CommitTrans%>
					<tr>
						<td colspan="2" class="content_b">
							<%	Select Case request("AZIONE")
									case "RESET" %>
									Ripristino pagina eseguito correttamente
							<% 		case "COPIA" %>
									Copia pagina eseguita correttamente.
							<% 		case "PUBBLICA" %>
									Pubblicazione eseguita correttamente
							<% 	end select %>
						</td>
					</tr>
					<tr>
						<td colspan="2" class="note">
							Questa finestra si chiuder&agrave; automaticamente tra 5 secondi.
						</td>
					</tr>
					<script language="JavaScript">
						<% if request("nextmail")="" then %>
							opener.location.reload(true);
						<% else %>
							opener.SetPreview( <%= ID_D %>);
						<% end if %>
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
<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
	PageOnLoad_FocusSet();
//-->
</script>
</body>
</html>
<%
conn.close
set rs = nothing
set conn = nothing %>