<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoPagineSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - indice delle pagine - nuova pagina"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = IIF(request("FROM")=FROM_ALBERO, "SitoPagineAlbero.asp", "SitoPagine.asp")
dicitura.scrivi_con_sottosez() 

dim i, lingua
dim conn, sql, rs
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_id_web" value="<%=Session("AZ_ID")%>">
	<input type="hidden" name="tfn_archiviata" value="0">
	<% for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE")) %>
		<input type="hidden" name="tfn_id_pagStage_<%= Application("LINGUE")(i) %>" value="0">
		<input type="hidden" name="tfn_id_pagDyn_<%= Application("LINGUE")(i) %>" value="0">
	<% next %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>Inserimento nuova pagina</caption>
		<tr><th colspan="4">DATI DELLA PAGINA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
			lingua = Application("LINGUE")(i)
			if Session("LINGUA_" & lingua) then%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label"  style="width:18%;" rowspan="<%= Session("LINGUE_ATTIVE") %>">titolo:</td>
			<% 	end if %>
				<td class="content" colspan="3">
					<img src="../grafica/flag_<%= lingua %>.jpg" alt="" border="0">
					<input type="text" class="text" name="tft_nome_ps_<%= lingua %>" value="<%= request("tft_nome_ps_"& lingua) %>" maxlength="250" style="width:90%;">
					(*)
				</td>
			</tr>
			<%end if
		next
		if index.GetTable("tb_pagineSito") > 0 then %>
			<tr>
				<td class="label" rowspan="2">collega la pagina a:</td>
				<td class="content" colspan="3">
					<% 	CALL index.WritePicker("", "", "form1", "idx", IIF(request("idx") <> "", request("idx"), request("indiceNew")), Session("AZ_ID"), false, false, "86", false, false) %>
				</td>
			</tr>
			<tr>
				<td class="content notes" colspan="3">
					Non immettendo il valore la pagina potr&agrave; essere collegata all'indice in un secondo momento.
				</td>
			</tr>
		<% end if %>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<tr><th class="L2" colspan="4">stato della pagina</th></tr>
		<% if IsAreaRiservataActive(conn) then %>
			<tr>
				<td class="label_no_width" style="width:18%;">
					<img src="../grafica/padlock.gif" border="0" alt="Pagina appartenente all'area protetta">
					protetta
				</td>
				<td class="content">
					<input type="radio" class="checkbox" value="1" name="tfn_riservata" <%= chk(cInteger(request("tfn_riservata"))=1) %>>
					si
					<input type="radio" class="checkbox" value="0" name="tfn_riservata" <%= chk(cInteger(request("tfn_riservata"))=0) %>>
					no
				</td>
				<td class="content_right" colspan="2">
					<span class="note">
						Proteggendo la pagina se ne permette la visualizzazione solo agli utenti dell'area riservata.
					</span>
				</td>
			</tr>
		<% else %>
			<input type="hidden" class="checkbox" value="0" name="tfn_riservata">
		<% end if %>
		<tr>
			<td class="label_no_width" style="width:18%;">
				<img src="../grafica/indicizzazione.gif" border="0" alt="Pagina indicizzabile">
				indicizzabile
			</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_indicizzabile" <%= chk(cInteger(request("tfn_indicizzabile"))=1 OR CString(request("tfn_indicizzabile"))="") %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_indicizzabile" <%= chk(cInteger(request("tfn_indicizzabile"))=0 AND NOT CString(request("tfn_indicizzabile"))="") %>>
				no
			</td>
			<td class="content_right" colspan="2">
				<span class="note">
					Scegliere se rendere indicizzabile questa pagina dai motori di ricerca come Google, Bing, Yahoo, ecc.
				</span>
			</td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<tr><th colspan="4">ASSOCIAZIONE DEL TEMPLATE ALLA PAGINA</th></tr>
		<tr>
			<td class="label" style="width:18%;" rowspan="<%= 3 + IIF(cInteger(Session("LINGUE_ATTIVE"))>1, 1 + Session("LINGUE_ATTIVE"), 0) %>">template da associare:</td>
			<td class="content_center" style="width:3%;">
				<input type="radio" class="noBorder" name="selezione_template" id="selezione_template_nessuno" onclick="Template_SetState()" value="" <%= chk(request("selezione_template")="") %>>
			</td>
			<td class="content" colspan="2">nessuno - lascia la pagina bianca (template vuoto)</td>
		</tr>
		<tr>
			<td class="content_center" rowspan="2">
				<input type="radio" class="noBorder" name="selezione_template" id="selezione_template_unico" onclick="Template_SetState()" value="unico" <%= chk(request("selezione_template")="unico") %>>
			</td>
			<td class="content" colspan="2">uguale per tutte le lingue</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<%sql = QryElencoTemplate("", false)
				CALL dropDown(conn, sql, "id_page", "name", "sel_template_unico", request("sel_template_unico"), TRUE, "", LINGUA_ITALIANO)%>
			</td>
		</tr>
		<% if cInteger(Session("LINGUE_ATTIVE"))>1 then %>
			<tr>
				<td class="content_center" rowspan="<%= 1 + cInteger(Session("LINGUE_ATTIVE")) %>">
					<input type="radio" class="noBorder" name="selezione_template" id="selezione_template_lingue" onclick="Template_SetState()" value="lingue" <%= chk(request("selezione_template")="lingue") %>>
				</td>
				<td class="content" colspan="2">per ogni lingua:</td>
			</tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
				lingua = Application("LINGUE")(i)
				if Session("LINGUA_" & lingua) then%>
				<tr>
					<td class="content" colspan="2">
						<img src="../grafica/flag_<%= lingua %>.jpg" alt="" border="0">
						<%sql = QryElencoTemplate("", false)
						CALL dropDown(conn, sql, "id_page", "name", "sel_template_" & lingua, request("sel_template_" & lingua), TRUE, "", LINGUA_ITALIANO)%>
					</td>
				</tr>
				<%end if
			next
		end if %>
		<script language="JavaScript" type="text/javascript">
			function Template_SetState(){
				var template_unico = document.getElementById("selezione_template_unico");
				EnableIfChecked(template_unico, form1.sel_template_unico);
				
				<% if cInteger(Session("LINGUE_ATTIVE"))>1 then %>
					var template_lingue = document.getElementById("selezione_template_lingue");
					<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
						lingua = Application("LINGUE")(i)
						if Session("LINGUA_" & lingua) then%>
							EnableIfChecked(template_lingue, form1.sel_template_<%= lingua %>);
						<%end if
					next
				end if %>
			}
			
			Template_SetState();
		</script>
		
		<tr><th colspan="4">COPIA PAGINA DA:</th></tr>
			<tr>
				<td class="label">copia da:</td>
				<td class="content" colspan="3">
					<% CALL DropDownPages(conn, "form1", "345", IIF(cBoolean(Session("COPIA_PAGINE_TRA_SITI"),false), 0, Session("AZ_ID")), "pagina_da_copiare", request("pagina_da_copiare"), false, false) %>
				</td>
			</tr>
		<tr>
		
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA &raquo;">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<%
conn.close
set conn = nothing
%>