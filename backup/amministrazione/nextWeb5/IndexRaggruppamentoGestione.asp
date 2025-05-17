<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_indice_accesso, 0))
	
if request.form("salva") <> "" then
	index.conn.BeginTrans
	
	dim co_id, campo
	'salvo il contenuto
	set index.content.dizionario = server.createobject("Scripting.Dictionary")
	index.content.co_F_key_id = 0
	index.content.co_F_table_id = index.GetTableNT(tabRaggruppamentoTable, tabRaggruppamento)
	index.content.dizionario("co_visibile") = 1
	for each campo in request.form
		if Left(campo, 3) = "co_" then
			index.content.dizionario(campo) = request.form(campo)
		end if
	next
	if request("ID") = "" then
		co_id = index.content.Salva(0)
	else
		co_id = index.content.Salva(GetValueList(index.conn, NULL, "SELECT idx_content_id FROM tb_contents_index WHERE idx_id = "& cIntero(request("ID"))))
	end if
	
	'salvo l'indice
	set index.dizionario = server.createobject("Scripting.Dictionary")
	for each campo in request.form
		if Left(campo, 4) = "idx_" then
			index.dizionario(campo) = request.form(campo)
		end if
	next
	index.dizionario("idx_content_id") = co_id
	CALL index.Salva(request("ID"))
	
	if session("ERRORE") = "" then
		index.conn.CommitTrans
		response.redirect IIF(request("FROM")=FROM_ALBERO, "IndexAlbero.asp", "IndexGenerale.asp")
	else
		index.conn.RollbackTrans
	end if
end if
	
	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
if request("ID") = "" then
	dicitura.sezione = "Indice generale - nuovo raggruppamento"
	dicitura.puls_new = "INDIETRO"
	dicitura.link_new = IIF(request("FROM")=FROM_ALBERO, "IndexAlbero.asp", "IndexGenerale.asp")
else
	dicitura.sezione = "Indice generale - modifica raggruppamento"
	dicitura.puls_new = "INDIETRO;VOCI COLLEGATE"
	dicitura.link_new = IIF(request("FROM")=FROM_ALBERO, "IndexAlbero.asp", "IndexGenerale.asp") &";IndexSottosezioni.asp?FROM="& request("FROM") &"&ID="& request("ID")
end if
dicitura.scrivi_con_sottosez()
	
	
	'codice simile a classIndex.Modifica
	dim dizionario, ID, conn, rs, linkPagina
	set conn = index.conn
	
	ID = CIntero(request("ID"))
	if ID > 0 AND request.servervariables("REQUEST_METHOD") <> "POST" then
		set rs = conn.execute(" SELECT * FROM tb_contents_index i"& _
							  " INNER JOIN tb_contents c ON i.idx_content_id = c.co_id"& _
							  " WHERE idx_id = "& ID)
		set dizionario = rs
	else
		set dizionario = request.form
	end if
	
	dim i %>
	<div id="content">
		<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>
				<% 	if ID = 0 then %>
					Inserimento nuovo raggruppamento
				<%	else %>
					Modifica raggruppamento
				<%	end if%>
			</caption>
			<tr><th colspan="4">DATI DEL RAGGRUPPAMENTO</th></tr>
			<tr>
				<td class="label" nowrap>voce collegata a:</td>
				<td class="content" colspan="3">
					<% 'sottocategoria creabile in tutte le categorie, anche con record associati
						CALL index.WritePicker("", "", "form1", "idx_padre_id", IIF(CIntero(dizionario("idx_padre_id")) > 0, dizionario("idx_padre_id"), request.querystring("idx_padre_id")), 0, false, false, "91", false, true) %>
				</td>
			</tr>
			
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
				<% 	if i = 0 then %>
					<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo:</td>
				<% 	end if %>
					<td class="content" colspan="3">
						<img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
						<input type="text" class="text" name="co_titolo_<%= Application("LINGUE")(i) %>" value="<%= dizionario("co_titolo_"& Application("LINGUE")(i)) %>" maxlength="255" size="85">
						<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "<span id=""obblN"">(*)</span>" end if %>
					</td>
				</tr>
			<%next %>
			
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
				<% 	if i = 0 then %>
					<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">codice univoco:</td>
				<% 	end if %>
					<td class="content" colspan="3">
						<img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
						<input type="text" class="text" name="co_chiave_<%= Application("LINGUE")(i) %>" value="<%= dizionario("co_chiave_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
					</td>
				</tr>
			<%next %>
			
				<tr><th class="L2" colspan="4">DATI DI PUBBLICAZIONE</th></tr>
				<tr>
					<td class="label">visibile:</td>
					<td class="content">
						<input type="radio" class="checkbox" value="1" name="co_visibile" <%= chk(dizionario("co_visibile") = "" OR CBool(dizionario("co_visibile"))) %>>
						si
						<input type="radio" class="checkbox" value="0" name="co_visibile" <%= chk(dizionario("co_visibile") <> "" AND CIntero(dizionario("co_visibile")) = 0 AND NOT CBool(dizionario("co_visibile"))) %>>
						no
					</td>
					<td class="label">ordine:</td>
					<td class="content">
						<input type="text" class="text" name="co_ordine" value="<%= dizionario("co_ordine") %>" maxlength="<%= index.OrdineLenght %>" size="3">
					</td>
				</tr>
				<tr>
					<td class="label">data pubblicazione:</td>
					<td class="content">
						<% CALL WriteDataPicker_Input("form1", "co_data_pubblicazione", dizionario("co_data_pubblicazione"), "", "/", true, false, LINGUA_ITALIANO) %>
					</td>
					<td class="label">data scadenza:</td>
					<td class="content">
						<% CALL WriteDataPicker_Input_Ex("form1", "co_data_scadenza", dizionario("co_data_scadenza"), "", "/", true, false, LINGUA_ITALIANO, "co_data_pubblicazione") %>
					</td>
				</tr>
				<tr><th class="L2" colspan="4">FOTO / IMMAGINI</th></tr>
				
				<tr>
					<td class="label">thumbnail:</td>
					<td class="content" colspan="3">
						<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "co_foto_thumb", dizionario("co_foto_thumb"), "width:482px", false) %>
					</td>
				</tr>
				<tr>
					<td class="label">zoom:</td>
					<td class="content" colspan="3">
						<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "co_foto_zoom", dizionario("co_foto_zoom"), "width:482px", false) %>
					</td>
				</tr>
				<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					<tr>
						<% 	if i = 0 then %>
							<td class="label_no_width" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo alternativo:</td>
						<% 	end if %>
						<td class="content" colspan="3">
							<img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
							<input type="text" class="text" name="idx_titolo_<%= Application("LINGUE")(i) %>" value="<%= dizionario("idx_titolo_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
						</td>
					</tr>
				<% next %>
			
			<%
			if IsCombinario() then
				for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					<tr>
					<% 	if i = 0 then %>
						<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">url manuale:</td>
					<% 	end if %>
						<td class="content" colspan="3">
							<img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
							<input type="text" class="text" name="idx_link_url_<%= Application("LINGUE")(i) %>" value="<%= dizionario("idx_link_url_"& Application("LINGUE")(i)) %>" maxlength="255" size="85">
						</td>
					</tr>
				<%next
			else
				for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					<input type="hidden" name="idx_link_url_<%= Application("LINGUE")(i) %>" value="<%= dizionario("idx_link_url_"& Application("LINGUE")(i)) %>">
				<%next
			end if%>
				<% 	'check dei permessi dell'utente
				if index.ChkPrm(prm_indice_permessi, 0) then
					'form di gestione dei permessi
					CALL index.prmForm(ID, "AL RAGGRUPPAMENTO")
				end if %>
			<tr>
				<td class="footer" colspan="4">
					(*) Campi obbligatori.
					<input type="submit" class="button" name="salva" value="SALVA">
				</td>
			</tr>
		</table>
		&nbsp;
		</form>
	</div>
	</body>
	</html>