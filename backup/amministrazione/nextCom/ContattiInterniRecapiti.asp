<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_contatti.ASP" --> 
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<%
'--------------------------------------------------------
sezione_testata = "Recapiti del contatto interno" 
testata_elenco_pulsanti = "SCHEDA"
testata_elenco_href = "ContattiInterniMod.asp?ID=" & request("ID")%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim conn, rs, rsc, sql, rubriche_visibili

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

'recupera rubriche visibili all'utente
rubriche_visibili = GetList_Rubriche(conn, rs)

if request("SALVA")<>"" then
	'controlli di correttezza (vengono bypassatti se il record e' sincronizzato)
	if request("ValoreNumero")="" AND request("SyncroField")="" then Session("ERRORE")="Manca il recapito!!"
	
	if cLng(request("id_TipoNumero"))=VAL_EMAIL AND request("SyncroField")="" then	'se email controlla che sia corretta
		if not isEmail(request("ValoreNumero")) then _
			Session("ERRORE") = Session("ERRORE") & "L'email inserita non &egrave; corretta. "
	end if
	if Session("ERRORE")="" then
		sql = "SELECT * FROM tb_ValoriNumeri WHERE id_ValoreNumero=" & cInteger(request("RID"))
		rs.Open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		if cInteger(request("RID"))=0 then
			rs.AddNew
			rs("id_Indirizzario") = request("ID")
		end if
		if request("SyncroField")="" then
			rs("ValoreNumero") = request("ValoreNumero")
			rs("id_TipoNumero") = request("id_TipoNumero")
			rs("protetto_privacy") = request("protetto_privacy")<>""
		end if
		rs("email_default") = request("email_default")<>"" and cInt(request("id_TipoNumero"))<>VAL_URL
		if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then
			rs("email_newsletter") = (request("email_newsletter")<>"") and cInt(request("id_TipoNumero"))=VAL_EMAIL
		end if
		rs.Update
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
		response.redirect "ContattiInterniRecapiti.asp?ID=" & request("ID")
	end if
end if

%>

<script language="JavaScript" type="text/javascript">
	function cambio_tiponumero(sender){
		document.form1.email_default.disabled = sender.options[sender.selectedIndex].value==<%= VAL_URL %>;
		document.form1.protetto_privacy.disabled = sender.options[sender.selectedIndex].value==<%= VAL_URL %>;
		<% if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then %>
			document.form1.email_newsletter.disabled = sender.options[sender.selectedIndex].value!=<%= VAL_EMAIL %>;
		<% end if %>
	}
</script>

<div id="content_ridotto">
<form action="" method="post" id="form1" name="form1">
<table cellspacing="1" cellpadding="0" class="tabella_madre">
	<% sql = "SELECT * FROM tb_indirizzario WHERE IDElencoIndirizzi=" & cIntero(request("ID"))
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText%>
	<caption>Recapiti del contatto interno &ldquo;<%= ContactName(rs) %>&rdquo;</caption>
	<% if Session("ERRORE")<>"" then %>
		<tr><td colspan="5" class="errore"><%= Session("ERRORE") %></td></tr>
		<%Session("ERRORE") = "" 
	end if 
	rs.close 
	sql = " SELECT tb_ValoriNumeri.*, tb_tipNumeri.nome_tipoNumero FROM tb_tipNumeri INNER JOIN tb_ValoriNumeri " &_
		  " ON tb_tipNumeri.id_tipoNumero = tb_ValoriNumeri.id_TipoNumero " &_
		  " WHERE id_indirizzario=" & request("ID") & _
		  " ORDER BY tb_ValoriNumeri.id_TipoNumero, email_default"
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText %>
	<tr>
		<th width="23%">TIPO</th>
		<th>RECAPITO</th>
		<th class="center" style="width:9%;">PRED.</th>
		<th class="center" style="width:7%;">PRIVACY</th>
		<% if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then %>
			<th class="center" style="width:8%;">NEWSLETTER</th>
		<% end if %>
		<th class="center" colspan="2" style="width: 25%;">OPERAZIONI</th>
	</tr>
		<% while not rs.eof 
			if cInteger(request("RID")) = rs("id_ValoreNumero") then
				'modifica del numero
				%>
				<input type="hidden" name="SyncroField" value="<%= rs("SyncroField") %>">
				<tr>
					<td class="content">
				<% 	sql = " SELECT * FROM tb_TipNumeri"
					CALL dropDown(conn, sql, "id_tipoNumero", "nome_tipoNumero", "id_TipoNumero", rs("id_TipoNumero"), true, disable(cString(rs("SyncroField"))<>"") & "style=""width:100%;"" onchange=""cambio_tiponumero(this)""", LINGUA_ITALIANO)
					if cString(rs("SyncroField"))<>"" then %>
						<input type="hidden" name="id_tipoNumero" value="<%= rs("id_TipoNumero") %>">
				<%	end if %>
					</td>
					<td class="content">
						<input type="text" class="<%= IIF(cString(rs("SyncroField"))<>"", "text_disabled", "text") %>" name="ValoreNumero" value="<%= rs("ValoreNumero") %>" maxlength="250" style="width:100%;" <%= disable(cString(rs("SyncroField"))<>"") %>>
					</td>
					<td class="Content_center">
						<input type="Checkbox" class="noBorder" name="email_default" <%= chk(rs("email_default"))%> <%= disable(rs("id_tipoNumero")=VAL_URL) %>>
					</td>
					<td class="Content_center">
						<input type="Checkbox" class="noBorder" name="protetto_privacy" <%=chk(rs("protetto_privacy"))%> <%=disable(rs("id_tipoNumero")=VAL_URL)%>>
					</td>
					<% if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then %>
						<td class="Content_center">
							<input type="Checkbox" class="noBorder" name="email_newsletter" <%= chk(rs("email_newsletter"))%> <%=disable(rs("id_tipoNumero")<>VAL_EMAIL)%>>
						</td>
					<% end if %>
					<td class="Content_center" style="vertical-align: middle;">
						<input style="width:100%;" type="submit" class="button" name="salva" value="SALVA">
					</td>
					<td class="Content_center" style="vertical-align: middle;">
						<input style="width:100%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="document.location='ContattiInterniRecapiti.asp?ID=<%= request("ID") %>';">
					</td>
				</tr>
			<% else %>
				<tr>
					<td class="content"><%= rs("nome_tipoNumero") %></td>
					<td class="content">
						<%= rs("ValoreNumero") %>
						<% CALL Check_DuplicatiRecapito(conn, rsc, rs("id_Indirizzario"), rs("ValoreNumero"), rubriche_visibili) %>
					</td>
					<td class="Content_center">
						<% if rs("id_tipoNumero")<>VAL_URL then%>
							<input type="Checkbox" class="checkbox" <%= chk(rs("email_default")) %> disabled>
						<% else %>
							&nbsp;
						<% end if %>
					</td>
					<td class="Content_center">
						<% if rs("id_tipoNumero")<>VAL_URL then%>
							<input type="Checkbox" class="checkbox" <%= chk(rs("protetto_privacy")) %> disabled>
						<% else %>
							&nbsp;
						<% end if %>
					</td>
					<% if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then %>
						<td class="Content_center">
							<% if rs("id_tipoNumero")<>VAL_URL then%>
								<input type="Checkbox" class="checkbox" <%= chk(rs("email_newsletter")) %> disabled>
							<% else %>
								&nbsp;
							<% end if %>
						</td>
					<% end if %>
					<td class="Content_center">
						<% if cString(rs("SyncroField"))="" OR rs("id_TipoNumero")=VAL_EMAIL then%>
							<a class="button_L2" href="?interno=<%= request.querystring("interno") %>&RID=<%= rs("id_ValoreNumero") %>&ID=<%= request("ID") %>">
								MODIFICA
							</a>
						<% else %>
							<a class="button_L2_disabled" title="contatto sincronizzato con dei dati di un'applicazione esterna.">
								MODIFICA
							</a>
						<% end if %>
					</td>
					<td class="Content_center">
						<% if cString(rs("SyncroField"))<>"" then%>
							<a class="button_L2_disabled" title="contatto sincronizzato con dei dati di un'applicazione esterna.">
								CANCELLA
							</a>
						<% else %>
							<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('RECAPITI','<%= rs("id_ValoreNumero") %>');" >
								CANCELLA
							</a>
						<% end if %>
					</td>
				</tr>
			<% end if
			rs.movenext
		wend
		if request("RID")="" then%>
			<tr>
				<td class="content">
					<% sql = " SELECT * FROM tb_TipNumeri"
					CALL dropDown(conn, sql, "id_tipoNumero", "nome_tipoNumero", "id_TipoNumero", request("id_tipoNumero"), true, "style=""width:100%;"" onchange=""cambio_tiponumero(this)""", LINGUA_ITALIANO)%>
				</td>
				<td class="content"><input type="text" class="text" name="ValoreNumero" value="<%= request("ValoreNumero") %>" maxlength="250" style="width:100%;"></td>
				<td class="Content_center">
					<input type="Checkbox" class="noBorder" name="email_default" <%=chk(request("email_default")<>"")%> <%=disable(cInteger(request("id_tipoNumero"))=VAL_URL)%>>
				</td>
				<td class="Content_center">
					<input type="Checkbox" class="noBorder" name="protetto_privacy" <%=chk(request("protetto_privacy")<>"" OR request("protetto_privacy")="")%> <%=disable(cInteger(request("id_tipoNumero"))=VAL_URL)%>>
				</td>
				<% if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then %>
					<%
					dim disabled 
					disabled = ""
					sql = Replace(sql, "SELECT *", "SELECT TOP 1 id_TipoNumero")
					if cIntero(GetValueList(conn, NULL, sql)) <> VAL_EMAIL then
						disabled = "disabled"
					end if
					%>
					<td class="Content_center">
						<input type="Checkbox" class="noBorder" name="email_newsletter" <%=chk(request("email_newsletter")<>"")%> <%=disabled%>>
					</td>
				<% end if %>
				<td class="Content_center" style="vertical-align: middle;">
					<input style="width:100%;" type="submit" class="button" name="salva" value="AGGIUNGI">
				</td>
				<td class="Content_center" style="vertical-align: middle;">
					<input style="width:100%;" type="button" class="button" name="reset" value="RESET" onclick="document.location='ContattiInterniRecapiti.asp?ID=<%= request("ID") %>';">
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="footer" <%=IIF(Session("ATTIVA_RECAPITI_NEWSLETTER")="True","colspan=""8""","colspan=""7""")%>>
				Tutti i campi sono obbligatori.
				<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
			</td>
		</tr>
</table>
</form>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set rsc = nothing
set conn = nothing
%>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this)
</script>