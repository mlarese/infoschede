<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
'call listsession()


'array per definire quanti e quali tipiNumero nei nuovi inserimenti dei recapiti
Dim tb_tipNumeri(2)
tb_tipNumeri(0) = 1
tb_tipNumeri(1) = 5
tb_tipNumeri(2) = 6
		
dim conn, rs, rsc, sql, record, i, value
dim isLocked, rubriche_visibili, isSocieta

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

'recupera rubriche visibili all'utente
rubriche_visibili = GetList_Rubriche(conn, rs)

if request("SALVA")<>"" OR Request.ServerVariables("REQUEST_METHOD")="POST" then
	dim RID
	
	'salvo recapiti già presenti
	sql = " SELECT id_ValoreNumero FROM tb_ValoriNumeri WHERE id_indirizzario=" & cIntero(request("ID")) 
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	while not rs.eof
		CALL SalvaRecapito(conn, rs("id_ValoreNumero"), "_"&rs("id_ValoreNumero"))
		rs.moveNext
	wend
	rs.close
	
	'salvo eventuali nuovi recapiti
	for i = 0 to ubound(tb_tipNumeri)
		if request("ValoreNumero_"&i) <> "" then
			CALL SalvaRecapito(conn, 0, "_"&i)
		end if
	next
	
	'conn.close
	'set rs = nothing
	'set conn = nothing
	
	if Session("ERRORE") = "" then
		%>
		<script language="JavaScript" type="text/javascript">
			parent.document.form1.submit();
			<% if request("salva_elenco")<>"" then %>
				parent.window.close();
			<% end if %>
		</script>
		<%
	else
		'response.end
		'response.redirect "ContattiRecapiti_iFrame.asp?ID=" & request("ID") & "&MODE=" & request("MODE")
	end if
end if

function SalvaRecapito(conn, RID, suffix)
	'controlli di correttezza (vengono bypassatti se il record e' sincronizzato)
	if request("ValoreNumero"&suffix)="" AND request("SyncroField"&suffix)="" then Session("ERRORE")="Manca il recapito!!"
	
	dim email
	email = Trim(request("ValoreNumero"&suffix))
	if cLng(request("id_TipoNumero"&suffix))=VAL_EMAIL AND request("SyncroField"&suffix)="" then	'se email controlla che sia corretta
		if not isEmail(email) then _
			Session("ERRORE") = Session("ERRORE") & "L'email inserita non &egrave; corretta."
	end if
	if Session("ERRORE")="" then
		sql = "SELECT * FROM tb_ValoriNumeri WHERE id_ValoreNumero=" & cInteger(RID)
		rsc.Open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		if cInteger(RID)=0 then
			rsc.AddNew
			rsc("id_Indirizzario") = request("ID")
		end if
		if request("SyncroField"&suffix)="" then
			rsc("ValoreNumero") = email
			rsc("id_TipoNumero") = request("id_TipoNumero"&suffix)
			rsc("protetto_privacy") = request("protetto_privacy"&suffix)<>""
		end if
		rsc("email_default") = (request("email_default"&suffix)<>"") and cInt(request("id_TipoNumero"&suffix))<>VAL_URL
		if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then
			rsc("email_newsletter") = (request("email_newsletter"&suffix)<>"") and cInt(request("id_TipoNumero"&suffix))=VAL_EMAIL
		end if
		rsc.Update
		rsc.close
	end if
end function


%>
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<% '*******************************************************************************************************************************
ParentFrameName = "IFrameRecapiti" %>
<!--#INCLUDE FILE="../library/Intestazione_iframe.asp" -->
<% '*******************************************************************************************************************************
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

<form action="" method="post" id="form1" name="form1"  style="margin-bottom:0px;">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-right:0px; border-left:0px;border-bottom:0px;" >
		<% if session("ERRORE")<> "" then %>
			<tr>
				<td class="errore" colspan="6"><%= Session("ERRORE")%></th>
			</tr>
			<% Session("ERRORE")="" %>
		<% end if %>
		<tr>
			<th colspan="6">RECAPITI</th>
		</tr>
		<%		
		sql = " SELECT tb_ValoriNumeri.*, tb_tipNumeri.nome_tipoNumero FROM tb_tipNumeri INNER JOIN tb_ValoriNumeri " &_
			  " ON tb_tipNumeri.id_tipoNumero = tb_ValoriNumeri.id_TipoNumero " &_
			  " WHERE id_indirizzario=" & cIntero(request("ID")) & _
			  " ORDER BY tb_ValoriNumeri.id_TipoNumero, email_default"
		rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText %>
		<tr>
			<th class="L2" width="22%">TIPO</th>
			<th class="L2">RECAPITO <%= IIF(isLocked<>"", "(**)", "") %></th>
			<th class="l2_center" style="width:7%;" title="PREDEFINITO">PREDEF.</th>
			<th class="l2_center" style="width:6%;">PRIVACY</th>
			<% if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then %>
				<th class="l2_center" style="width:7%;">NEWSLETTER</th>
			<% end if %>
			<th class="l2_center" style="width:8%;">OPERAZIONI</th>
		</tr>
		<% while not rs.eof 
			%>
			<input type="hidden" name="SyncroField_<%=rs("id_ValoreNumero")%>" value="<%= rs("SyncroField") %>">
			<tr>
				<td class="content">
					<% 
					if request("ValoreNumero_"&rs("id_ValoreNumero")) <> "" then
						value = request("id_TipoNumero_"&rs("id_ValoreNumero"))
					else
						value = rs("id_TipoNumero")
					end if
					%>
					<% 	sql = " SELECT * FROM tb_TipNumeri"
						CALL dropDown(conn, sql, "id_tipoNumero", "nome_tipoNumero", "id_TipoNumero_"&rs("id_ValoreNumero"), rs("id_TipoNumero"), true, disable(cString(rs("SyncroField"))<>"") & "style=""width:100%;"" onchange=""cambio_tiponumero(this)""", LINGUA_ITALIANO)
						if cString(rs("SyncroField"))<>"" then %>
							<input type="hidden" name="id_tipoNumero_<%=rs("id_ValoreNumero")%>" value="<%= value %>">
					<%	end if %>
				</td>
				<td class="content">
					<% 
					if request("ValoreNumero_"&rs("id_ValoreNumero")) <> "" then
						value = request("ValoreNumero_"&rs("id_ValoreNumero"))
					else
						value = rs("ValoreNumero")
					end if
					%>
					<input type="text" class="<%= IIF(cString(rs("SyncroField"))<>"", "text_disabled", "text") %>" name="ValoreNumero_<%=rs("id_ValoreNumero")%>" value="<%= value %>" maxlength="250" style="width:90%;" <%= disable(cString(rs("SyncroField"))<>"") %>>&nbsp;
					<% CALL Check_DuplicatiRecapito(conn, rsc, rs("id_Indirizzario"), rs("ValoreNumero"), rubriche_visibili) %>
				</td>
				<td class="Content_center">
					<input type="Checkbox" class="noBorder" name="email_default_<%=rs("id_ValoreNumero")%>" <%= chk(rs("email_default"))%> <%=disable(rs("id_tipoNumero")=VAL_URL)%>>
				</td>
				<td class="Content_center">
					<input type="Checkbox" class="noBorder" name="protetto_privacy_<%=rs("id_ValoreNumero")%>" <%=chk(rs("protetto_privacy"))%> <%=disable(rs("id_tipoNumero")=VAL_URL)%>>
				</td>
				<% if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then %>
					<td class="Content_center">
						<input type="Checkbox" class="noBorder" name="email_newsletter_<%=rs("id_ValoreNumero")%>" <%= chk(rs("email_newsletter"))%> <%=disable(rs("id_tipoNumero")<>VAL_EMAIL)%>>
					</td>
				<% end if %>
				<td class="Content_center">
					<% if cString(rs("SyncroField"))<>"" then%>
						<a class="button_disabled" title="contatto sincronizzato con dei dati di un'applicazione esterna.">
							CANC
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('RECAPITI','<%= rs("id_ValoreNumero") %>');" >
							CANC
						</a>
					<% end if %>
				</td>
			</tr>
			<%
			rs.movenext
		wend
		
		for i = 0 to ubound(tb_tipNumeri)
			%>
			<tr>
				<td class="content">
					<% sql = " SELECT * FROM tb_TipNumeri"
					dim valore
					valore = tb_tipNumeri(i)
					if Request.ServerVariables("REQUEST_METHOD")="POST" then
						valore = request("id_TipoNumero_"&i)
					end if
					CALL dropDown(conn, sql, "id_tipoNumero", "nome_tipoNumero", "id_TipoNumero_"&i, valore, true, "style=""width:100%;"" onchange=""cambio_tiponumero(this)""", LINGUA_ITALIANO)%>
				</td>
				<td class="content"><input type="text" class="text" name="ValoreNumero_<%=i%>" value="<%= request("ValoreNumero_"&i) %>" maxlength="250" style="width:99%;">&nbsp;</td>
				<td class="Content_center">
					<input type="Checkbox" class="noBorder" name="email_default_<%=i%>" <%=chk(request("email_default_"&i)<>"")%> >
				</td>
				<td class="Content_center">
					<input type="Checkbox" class="noBorder" name="protetto_privacy_<%=i%>" <%=chk(request("protetto_privacy_"&i)<>"")%> <%=disable(cInteger(request("id_tipoNumero_"&i))=VAL_URL)%>>
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
						<input type="Checkbox" class="noBorder" name="email_newsletter_<%=i%>" <%=chk(request("email_newsletter_"&i)<>"")%> <%=disabled%>>
					</td>
				<% end if %>
				<td class="Content_center">&nbsp;</td>
				<!--
				<td class="Content_center" style="vertical-align: middle;">
					<input style="width:92%;" type="submit" class="button" name="salva" value="AGGIUNGI">
				</td>
				<td class="Content_center" style="vertical-align: middle;">
					<input style="width:97%;" type="button" class="button" name="reset" value="RESET" onclick="document.location='ContattiRecapiti_iFrame.asp?ID=<%= request("ID") %>&MODE=<%= request("MODE") %>';">
				</td>
				-->
			</tr>
		<% next %>
		<tr>
			<td <%=IIF(cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false),"colspan=""6""","colspan=""5""")%> class="note">
				<% if isLocked<>"" then %>
					(**) Alcuni valori dei recapiti potrebbero essere non modificabili perch&egrave; impostati dalle applicazioni che gestiscono il contatto.<br>
				<% end if %>
				All'invio di una nuova email dalla sezione mailing al contatto verr&agrave; inviata una copia del messaggio per ogni indirizzo impostato come predefinito.
			</td>
		</tr>
		<tr><td <%=IIF(cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false),"colspan=""6""","colspan=""5""")%> class="content">&nbsp;</td></tr>
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