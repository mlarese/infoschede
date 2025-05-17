<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_menu_accesso, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoMenuLinkSalva.asp")
end if

'--------------------------------------------------------
sezione_testata = "Gestione siti - menu - nuovo link" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim conn, i, lingua
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_mi_menu_id" value="<%= request("MENU") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo link</caption>
		<tr><th colspan="4">DATI DEL LINK</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
			lingua = Application("LINGUE")(i)
			if Session("LINGUA_" & lingua) then%>
				<tr>
				<% 	if i = 0 then %>
					<td colspan="2" style="width:17%;" class="label" rowspan="<%= Session("LINGUE_ATTIVE") %>">titolo:</td>
				<% 	end if %>
					<td class="content" colspan="2">
						<img src="../grafica/flag_<%= lingua %>.jpg">
						<input type="text" class="text" name="tft_mi_titolo_<%= lingua %>" value="<%= request("tft_mi_titolo_"& lingua) %>" maxlength="255" style="width:90%;">
						<% 	if lingua = LINGUA_ITALIANO then response.write "(*)" end if %>
					</td>
				</tr>
			<%end if
		next %>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
			lingua = Application("LINGUE")(i)
			if Session("LINGUA_" & lingua) then%>
				<tr>
				<% 	if i = 0 then %>
					<td colspan="2" class="label" rowspan="<%= Session("LINGUE_ATTIVE") %>">tag title:</td>
				<% 	end if %>
					<td class="content" colspan="2">
						<img src="../grafica/flag_<%= lingua %>.jpg">
						<input type="text" class="text" name="tft_mi_tag_title_<%= lingua %>" value="<%= request("tft_mi_tag_title_"& lingua) %>" maxlength="255" style="width:90%;">
					</td>
				</tr>
			<%end if
		next %>
		<tr>
			<td colspan="2" class="label" rowspan="2">dati pubblicazione:</td>
			<td class="label">visibile:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_mi_attivo" <%= chk(request("tfn_mi_attivo")<>"0")%>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_mi_attivo" <%= chk(request("tfn_mi_attivo")="0")%>>
				no
			</td>
		</tr>
		<tr>
			<td class="label">ordine</td>
			<td class="content">
				<input type="text" class="text" name="tfn_mi_ordine" value="<%= request("tfn_mi_ordine") %>" maxlength="4" size="4">
			</td>
		</tr>
		<tr><th colspan="4">COLLEGAMENTO</th></tr>
		<script language="JavaScript" type="text/javascript">
			function SetStato(){
				var tipo = document.getElementById("tipo_link_interno");
				<%for each lingua in Application("LINGUE")
					if Session("LINGUA_" & lingua) then%>
						DisableIfChecked(tipo, form1.tft_mi_link_<%= lingua %>);
					<% end if
				next %>
				DisablePicker(form1.tfn_mi_index_id, !tipo.checked);
				DisablePicker(document.getElementById('chk_mi_figli'), !tipo.checked);
			}
		</script>
		<tr>
			<td class="label_no_width" style="width:13%;" rowspan="<%= 4 + Session("LINGUE_ATTIVE") %>">link a:</td>
			<td class="content_center" rowspan="3">
				<input type="radio" class="noBorder" name="tipo_link" id="tipo_link_interno" value="" <%= chk(request("tipo_link")<>"0")%> onclick="SetStato()">
			</td>
			<td class="content" colspan="2">voce dell'indice (risorsa interna)</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<% 	CALL index.WritePicker("", "", "form1", "tfn_mi_index_id", request("tfn_mi_index_id"), Session("AZ_ID"), false, false, "77", false, false) %>
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="checkbox" name="chk_mi_figli" id="chk_mi_figli" value="1" class="checkbox" <%= Chk(request("chk_mi_figli")) %>>
				abilita visualizzazione delle voci figlie
			</td>
		</tr>
		<tr>
			<td class="content_center" rowspan="<%= 1 + Session("LINGUE_ATTIVE") %>">
				<input type="radio" class="noBorder" name="tipo_link" id="tipo_link_esterno" value="0" <%= chk(request("tipo_link")="0")%> onclick="SetStato()">
			</td>
			<td class="content" colspan="2">risorsa esterna</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
			lingua = Application("LINGUE")(i)
			if Session("LINGUA_" & lingua) then%>
				<tr>
					<td class="content" colspan="2">
						<table cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td width="30px"><img src="../grafica/flag_<%= lingua %>.jpg"></td>
								<td>
									<input type="text" class="text" name="tft_mi_link_<%= lingua %>" value="<%= request("tft_mi_link_"& lingua) %>" maxlength="255" style="width:95%;">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			<%end if
		next %>
		<tr>
			<td colspan="2" class="label">target:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_mi_target" value="<%= request("tft_mi_target") %>" maxlength="50" size="14">
			</td>
		</tr>
		<tr><th colspan="4">IMMAGINI</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
			lingua = Application("LINGUE")(i)
			if Session("LINGUA_" & lingua) then%>
				<tr>
				<% 	if i = 0 then %>
					<td class="label" colspan="2" rowspan="<%= Session("LINGUE_ATTIVE") %>">immagine associata:</td>
				<% 	end if %>
					<td class="content" colspan="2">
						<table cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td width="30px"><img src="../grafica/flag_<%= lingua %>.jpg"></td>
								<td><% CALL WriteFilePicker_Input(Session("AZ_ID"), "images", "form1", "tft_mi_image_"& lingua, request("tft_mi_image_"& lingua) , "width:374px;", false) %></td>
							</tr>
						</table>
					</td>
				</tr>
			<%end if
		next %>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori
				<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
				<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
			</td>
		</tr>
	</table>
	</form>
</div>
</body>
</html>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
	SetStato();
</script>
<%
conn.Close
set conn = nothing
%>