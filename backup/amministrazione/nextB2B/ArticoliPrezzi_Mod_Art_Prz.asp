<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<% 
dim conn, rs, rsp, sql, tipo
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsp = Server.CreateObject("ADODB.Recordset")

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	'aggiornamento dei prezzi
	if request("prezzo")="" OR not isNumeric(request("prezzo")) then
		Session("ERRORE") = "Prezzo immesso non valido."
	end if
	if not GetModuleParam(conn, "LISTINI_PREZZI_INDIPENDENTI") then
		if request("aggiornamento_varianti") = "" OR  request("aggiornamento_listini")="" then
			Session("ERRORE") = "Scegliere il metodo di aggiornamento dei prezzi."
		end if	
	end if
	'if cBoolean(cString(Session("INIBISCI_PREZZO_A_ZERO")), false) AND cReal(request("prezzo")) = 0 then
	'	Session("ERRORE") = "Impossibile inserire un articolo con il prezzo uguale a zero"
	'end if
	if Session("ERRORE") = "" then
		dim prezzo, aggiornamento
		prezzo = cReal(request("prezzo"))
		
		conn.beginTrans
		sql = "SELECT * FROM gtb_articoli WHERE art_id=" & cIntero(request("ID"))
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		'aggiorna prezzo base dell'articolo
		rs("art_prezzo_base") = prezzo
		' Da verificare
		rs("art_modData") = now
		rs.update
		rs.close
		
		'aggiornamento varianti con prezzo dipendente o variante unica di sistema
		CALL AggiornaPrezziVarianti(conn, rs, request("ID"))
		
		'gestione varianti con prezzo indipendente
		if cInteger(request("aggiornamento_varianti"))>0 then
			'imposta prezzo corrente per tutte le varianti indipendenti
			sql = "SELECT * FROM grel_art_valori WHERE ISNULL(rel_prezzo_indipendente, 0)=1 AND rel_art_id=" & cIntero(request("ID"))
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			while not rs.eof
				rs("rel_prezzo") = prezzo
				rs("rel_var_euro") = 0
				rs("rel_var_sconto") = 0
				rs.update
				rs.movenext
			wend
			rs.close
		end if
		

		'if not GetModuleParam(conn, "LISTINI_PREZZI_INDIPENDENTI") then
			'aggiorna prezzi dei listini base
			aggiornamento = cInteger(request("aggiornamento_listini"))
			if aggiornamento = 0 OR aggiornamento = 2 then
				'aggiornamento diretto via sconti dei listini da varianti
				if aggiornamento = 2 then
					'elimina prima sconti e variaizoni
					sql = " UPDATE gtb_prezzi SET prz_var_sconto=0, prz_var_euro=0 " + _
						  " WHERE prz_variante_id IN (SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & cIntero(request("ID")) & ")" + _
						  " AND prz_listino_id IN (SELECT listino_id FROM gtb_listini WHERE listino_base=1) "

					CALL conn.execute(sql, , adExecuteNoRecords)
				end if
				
				'aggiorna i prezzi dei listini base e di tutti i listini successivi
				sql = "SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & cIntero(request("ID"))
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				while not rs.eof
					CALL AggiornaPrezzoListiniBaseForzabile(conn, rs("rel_id"), false)
					rs.movenext
				wend
				rs.close
			elseif aggiornamento = 1 then
				'aggiornamento sconti da prezzo base via prezzo variante
				'sostituisce i prezzi e ricalcola gli sconti per mantenere inalterati i prezzi.
				
				sql = " SELECT * FROM gtb_prezzi " + _
					  " WHERE prz_variante_id IN (SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & cIntero(request("ID")) & ") " + _
					  " AND prz_listino_id IN (SELECT listino_id FROM gtb_listini WHERE listino_base=1) "
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				while not rs.eof
					'recupera prezzo attuale variante
					sql = "SELECT rel_prezzo FROM grel_art_valori WHERE rel_id=" & rs("prz_variante_id")
					prezzo = cReal(GetValueList(conn, rsp, sql))
					if cReal(rs("prz_var_euro"))<>0 OR prezzo = 0 then
						rs("prz_var_sconto") = 0
						rs("prz_var_euro") = rs("prz_prezzo") - prezzo
					else
						rs("prz_var_sconto") = GetVarPercent(prezzo, rs("prz_prezzo"))
						rs("prz_var_euro") = 0
					end if
					rs.update
					rs.movenext
				wend
				rs.close
				
				'aggiorna prezzi dei listini non base
				sql = "SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & cIntero(request("ID"))
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				while not rs.eof
					CALL AggiornaPrezzoListiniDaListinoBase(conn, rs("rel_id"))
					rs.movenext
				wend
				rs.close
			end if
		'end if
 		conn.commitTrans
		%>
		<script language="JavaScript" type="text/javascript">
			opener.location.reload(true);
			window.close();
		</script>
	<% end if
end if


sql = "SELECT * FROM gtb_articoli WHERE art_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

if rs("art_se_bundle") then
	tipo = "del bundle"
elseif rs("art_se_confezione") then
	tipo = "della confezione"
elseif rs("art_varianti") then
	tipo ="dell'articolo con varianti"
else
	tipo ="dell'articolo singolo"
end if
%>

<%'--------------------------------------------------------
sezione_testata = "modifica prezzo " & tipo %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>
				Modifica del prezzo base <%= tipo %>
			</caption>
			<tr><th colspan="3">DATI BASE ARTICOLO</th></tr>
			<tr>
				<td class="label" rowspan="2">articolo</td>
				<td class="content" colspan="2"><%= rs("art_nome_it") %></td>
			</tr>
			<tr>
				<% if rs("art_se_bundle") then %>
					<td colspan="2" class="content bundle">bundle</td>
				<% elseif rs("art_se_confezione") then %>
					<td colspan="2" class="content confezione">confezione</td>
				<% elseif rs("art_varianti") then %>
					<td colspan="2" class="content varianti">articolo con varianti</td>
				<% else %>
					<td colspan="2" class="content">articolo singolo</td>
				<% end if %>
			</tr>
			<tr>
				<td class="label">prezzo</td>
				<td class="content" colspan="2"><input type="text" class="number" name="prezzo" value="<%= FormatPrice(rs("art_prezzo_base"), 2, false) %>" size="7"> &euro;&nbsp;&nbsp;&nbsp;&nbsp;(*)</td>
			</tr>
			<% if rs("art_varianti") then
				sql = "SELECT COUNT(*) FROM grel_Art_valori WHERE ISNULL(rel_prezzo_indipendente, 0)=1 AND rel_art_id=" & rs("art_id")
				if cInteger(GetValueList(conn, rsp, sql))>0 then %>
					<tr><th colspan="3">AGGIORNAMENTO VARIANTI CON PREZZI INDIPENDENTI</th></tr>
					<tr>
						<td class="label" rowspan="2">tipo:</td>
						<td class="content_center">
							<input type="radio" class="checkbox" name="aggiornamento_varianti" value="0" <%= chk(cInteger(request("aggiornamento_varianti"))=0) %>>
						</td>
						<td class="content">
							non aggiornare i prezzi delle varianti<br>
							<span class="note">
								I prezzi delle varianti calcolati sulla base di sconti dal prezzo articolo verranno comunque aggiornati
							</span>
						</td>
					</tr>
					<tr>
						<td class="content_center">
							<input type="radio" class="checkbox" name="aggiornamento_varianti" value="1" <%= chk(cInteger(request("aggiornamento_varianti"))=1) %>>
						</td>
						<td class="content">
							aggiorna anche i prezzi delle varianti<br>
							<span class="note">
								Sostituisce i prezzi delle varianti i cui prezzi sono indipendenti e ricalcola i prezzi delle altre.
							</span>
						</td>
					</tr>
				<% else %>
					<input type="hidden" name="aggiornamento_varianti" value="0">
					<tr><th colspan="3">AGGIORNAMENTO VARIANTI</th></tr>
					
					<tr>
						<td class="note" colspan="3">
							I prezzi delle varianti vengono ricalcolati automaticamente sulla base degli sconti applicati.
						</td>
					</tr>
				<% end if
			else 
				'se non ci sono varianti il prezzo va aggiornato direttamente
				%>
				<input type="hidden" name="aggiornamento_varianti" value="1">
			<% end if 
			
			if not GetModuleParam(conn, "LISTINI_PREZZI_INDIPENDENTI") then %>
				<tr><th colspan="3">AGGIORNAMENTO LISTINI BASE</th></tr>
				<tr>
					<td class="label" rowspan="3">tipo:</td>
					<td class="content_center">
						<input type="radio" class="checkbox" name="aggiornamento_listini" value="0" <%= chk(cInteger(request("aggiornamento_listini"))=0) %>>
					</td>
					<td class="content">
						ricalcola i prezzi<br>
						<span class="note">Il calcolo verr&agrave; effettuato sulla base delle variazioni (in &euro; o %) applicate.</span>
					</td>
				</tr>
				<tr>
					<td class="content_center">
						<input type="radio" class="checkbox" name="aggiornamento_listini" value="1" <%= chk(cInteger(request("aggiornamento_listini"))=1) %>>
					</td>
					<td class="content">
						lascia prezzi inalterati<br>
						<span class="note">il sistema ricalcoler&agrave; automaticamente le variazioni (in &euro; o %) applicate.</span>
					</td>
				</tr>
				<tr>
					<td class="content_center">
						<input type="radio" class="checkbox" name="aggiornamento_listini" value="2" <%= chk(cInteger(request("aggiornamento_listini"))=2) %>>
					</td>
					<td class="content">
						sostituisci i prezzi ed azzera variazioni (in &euro; o %) applicate.
					</td>
				</tr>
				<tr>
					<td class="note" colspan="3">ATTENZIONE: i listini delle offerte speciali e dei clienti verranno aggiornati solo nei prezzi: nessuna variazione verr&agrave; effettuata sulle variazioni (in &euro; o %) applicate.</td>
				</tr>
			<% else %>
				<tr><th colspan="3">AGGIORNAMENTO LISTINI BASE</th></tr>
				<tr>
					<td class="label" rowspan="2">tipo:</td>
					<td class="content_center">
						<input type="radio" class="checkbox" name="aggiornamento_listini" value="1" <%= chk(cInteger(request("aggiornamento_listini"))=1 OR cString(request("aggiornamento_listini"))="") %>>
					</td>
					<td class="content">
						lascia prezzi inalterati<br>
						<span class="note">il sistema ricalcoler&agrave; automaticamente le variazioni (in &euro; o %) applicate.</span>
					</td>
				</tr>
				<tr>
					<td class="content_center">
						<input type="radio" class="checkbox" name="aggiornamento_listini" value="2" <%= chk(cInteger(request("aggiornamento_listini"))=1) %>>
					</td>
					<td class="content">
						sostituisci i prezzi ed azzera variazioni (in &euro; o %) applicate.
					</td>
				</tr>
				<!--
				<tr>
					<td class="note" colspan="3">
						ATTENZIONE: la variazione di prezzo non verr&agrave; propagata nei listini perch&egrave; attiva la modalit&agrave; Listini Indipendenti.
					</td>
				</tr>
				-->
			<% end if %>
			<tr>
				<td class="footer" colspan="3">
					(*) Campi obbligatori.
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
					<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
				</td>
			</tr>
		</form>
	</table>
</div>
</body>
</html>
<% rs.close
conn.close
set rs = nothing
set rsp = nothing
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>