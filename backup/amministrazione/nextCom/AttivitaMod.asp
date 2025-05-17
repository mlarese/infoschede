<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_Attivita.asp" -->
<%
dim conn, rs, rsp, sql, conclusa, risposta

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.createobject("adodb.recordset")
set rsp = server.createobject("adodb.recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session(Session("ATT_PREFIX") & "SQL_ATTIVITA"), "att_id", "AttivitaMod.asp")
end if

'controllo accesso
if NOT AL(conn, cIntero(request("ID")), AL_ATTIVITA) then
	Session("ERRORE") = "Impossibile visualizzare l'attivit&agrave;: permessi non validi."
	response.redirect "Attivita.asp"
end if

if Request.ServerVariables("REQUEST_METHOD")="POST" AND request("MOD")<>"" then
	Server.Execute("AttivitaSalva.asp")
end if


'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
if Session("ATT_PRA_ID")<>"" then
	Titolo_sezione = "Pratiche - attivit&agrave; della pratica - modifica"
elseif Session("ATT_DOC_ID")<>"" then
	'attivita' collegate al documento
	Titolo_sezione = "Documenti - attivit&agrave; collegate al documento - modifica"
else
	Titolo_sezione = "Attivit&agrave; - modifica"
end if
	HREF = "Attivita.asp"
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->
<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************

sql = "SELECT tb_Attivita.*, admin_nome, admin_cognome, (SELECT TOP 1 att_id FROM tb_attivita r " & _
	  " WHERE r.att_domanda_id=tb_attivita.att_id AND r.att_id<>tb_attivita.att_id) AS RISPOSTA, " & _
	  " (SELECT COUNT(*) FROM tb_documenti INNER JOIN tb_allegati ON tb_documenti.doc_id = tb_allegati.all_documento_id " & _
	  "  WHERE "& AL_query(conn, AL_DOCUMENTI) & " AND all_attivita_id=" & cIntero(request("ID")) & ") AS ALLEGATI " & _
	  " FROM tb_attivita INNER JOIN tb_admin ON tb_Attivita.att_mittente_id = tb_admin.id_admin WHERE att_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

'GESTIONE MODIFICA
If (rs("att_conclusa") OR NOT rs("att_inSospeso") OR _
	rs("att_inSospeso") AND Session("ID_ADMIN") <> rs("att_mittente_id")) AND Session("COM_ADMIN") = "" then
	conclusa = "disabled"
end if
if rs("att_pratica_id") <> 0 then
	sql = "SELECT COUNT(*) FROM tb_pratiche WHERE " & SQL_isTrue(conn, "pra_archiviata") & " AND pra_id=" & rs("att_pratica_id")
 	if cInteger(GetValueList( conn, rsp, sql))>0 then
		conclusa = "disabled"
	end if
end if
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_att_mittente_id" value="<%= rs("att_mittente_id") %>">
	<% if (Session("ID_ADMIN")<>"" AND conclusa="") OR rs("att_inSospeso") OR session("COM_ADMIN") <> "" then  %>
		<input type="hidden" name="tfd_att_dataCrea" value="NOW">
	<% end if %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<% if Session(Session("ATT_PREFIX") & "SQL_ATTIVITA")<>"" then 
				'verifica se esiste elenco pratiche%>
				<table border="0" cellspacing="0" cellpadding="0" align="right">
					<tr>
						<td style="font-size: 1px; padding-right:1px;" nowrap>
							<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="attivit&agrave; precedente">
								&lt;&lt; PRECEDENTE
							</a>
							&nbsp;
							<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="attivit&agrave; successiva">
								SUCCESSIVA &gt;&gt;
							</a>
						</td>
					</tr>
				</table>
			<% end if %>
			Vedi / Modifica attivit&agrave;
		</caption>
		<%if cInteger(rs("RISPOSTA")) > 0 OR cInteger(rs("att_domanda_id"))>0 then %>
			<tr><th colspan="4">ATTIVIT&Agrave; COLLEGATE</th></tr>
			<tr>
				<td colspan="4" class="content_RIGHT" style="font-size:1px;">
					<% if cInteger(rs("att_domanda_id"))>0 AND rs("att_domanda_id")<>rs("att_id") then %>
						<a href="AttivitaMod.asp?ID=<%= rs("att_domanda_id") %>" class="button_L2" title="Visualizza l'attivita' a cui ho risposto">
							&lt;&lt; ATTIVITA' COLLEGATA PRECEDENTE
						</a>
					<% else %>
						<a class="button_L2_disabled">
							&lt;&lt; ATTIVITA' COLLEGATA PRECEDENTE
						</a>
					<% end if %>
					&nbsp;
					<%if cInteger(rs("RISPOSTA")) > 0 then%>
						<a href="AttivitaMod.asp?ID=<%= rs("RISPOSTA") %>" class="button_L2" title="Visualizza la risposta">
							ATTIVITA' COLLEGATA SUCCESSIVA &gt;&gt;
						</a>
					<% else %>
						<a class="button_L2_disabled" title="nessuna risposta">
							ATTIVITA' COLLEGATA SUCCESSIVA &gt;&gt;
						</a>
					<% end if %>
				</td>
			</tr>
		<% end if %>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<% if Session("ATT_PRA_ID")<>"" then
			CALL SelezionaPratica(conn, rsp, "ATT", rs("att_pratica_id"), false) 
		else
			CALL SelezionaPratica(conn, rsp, "ATT", rs("att_pratica_id"), conclusa="") 
		end if
		%>
		<tr>
			<td class="label">oggetto:</td>
			<td class="content" colspan="3">
				<% if conclusa<>"" then %>
					<input type="hidden" name="tft_Att_oggetto"value="<%= rs("att_oggetto") %>">
				<% end if %>
				<input <%= conclusa %> type="text" class="text" name="tft_att_oggetto" value="<%= rs("att_oggetto") %>" maxlength="255" size="75">
				<span id="oggetto">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">data scadenza:</td>
			<td class="content" colspan="3">
				<%If conclusa = "" then
					CALL WriteDataPicker_Input("form1", "tfd_att_dataS", rs("att_dataS"), "", "/", true, true, LINGUA_ITALIANO)
				else %>
					<%= DateIta(rs("att_dataS")) %>
				<%End If %>
			</td>
		</tr>
		<tr>
			<td class="label">prioritaria:</td>
			<td class="content">
				<% if conclusa<>"" then %>
					<input type="hidden" name="chk_att_priorita" value="<%= IIF(rs("att_priorita"), "1", "") %>">
					<input type="Checkbox" disabled class="noborder" <%= Chk(rs("att_priorita")) %>>
				<% else %>
					<input type="Checkbox" name="chk_att_priorita" value="1" class="noborder" <%= Chk(rs("att_priorita")) %>>
				<% end if %>
			</td>
			<td class="label">conclusa:</td>
			<td class="content">
				<input <%= IIF(rs("att_conclusa") AND Session("COM_ADMIN") = "" AND NOT rs("att_inSospeso"), " disabled ", "") %> type="Checkbox" name="chk_att_conclusa" value="1" class="noborder" <%= Chk(rs("att_conclusa")) %>>
			</td>
		</tr>
		<tr>
			<td class="label">mittente:</td>
			<td class="content"><%= rs("admin_cognome") %>&nbsp;<%= rs("admin_nome") %></td>
			<td class="label">data creazione:</td>
			<td class="content"><%= DateTimeIta(rs("att_DataCrea")) %></tr>
		</tr>
		<% if rs("att_conclusa") then %>
			<tr>
				<td class="label">conclusa da:</td>
				<td class="content">
					<% if cInteger(rs("att_utente_chiusura")) > 0 then 
						sql = "SELECT admin_cognome, admin_nome FROM tb_admin WHERE id_admin=" & rs("att_utente_chiusura")
						rsp.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText%>
						<%= rs("admin_cognome") %>&nbsp;<%= rs("admin_nome") %>
						<% rsp.close
					end if %>
				</td>
				<td class="label">data chiusura:</td>
				<td class="content"><%= DateTimeIta(rs("att_DataChiusa")) %></tr>
			</tr>
		<% end if %>
		<tr><th colspan="4">TESTO (*)</th></tr>
		<tr>
			<td class="content" colspan="4">
				<% if conclusa<>"" then %>
					<input type="hidden" name="tft_att_testo" value="<%= Server.HTMLEncode(rs("att_testo")) %>">
					<span class="overflow" style="height:170px;">
						<%= TextEncode(rs("att_testo"))%>
					</span>
				<% else %>
					<textarea <%= conclusa%> style="width:100%;" rows="12" name="tft_att_testo"><%= rs("att_testo")%></textarea>
				<% end if %>
			</td>
		</tr>
		<%If conclusa <> ""  then
			if rs("ALLEGATI") > 0 then %>
				<tr>
					<th colspan="4">
						DOCUMENTI ALLEGATI
					</th>
				</tr>
				<tr>
					<td class="content" colspan="4" style="padding-bottom:1px;">
						<a href="Documenti.asp?ATT_ID=<%= rs("att_id") %>" class="button_form" target="_balnk">
							VISUALIZZA ELENCO DOCUMENTI ALLEGATI
						</a>
					</td>
				</tr>
			<%end if
		Else
			CALL GestioneDocumentiCollegati(conn, rsp, rs("att_id"))
		End If
			
		if NOT rs("att_conclusa") AND rs("ALLEGATI") > 0 then
			dim utentiOrbi
			sql = "SELECT DISTINCT id_admin FROM tb_admin INNER JOIN tb_rel_dipGruppi ON " & _
				  "tb_admin.id_admin=tb_rel_dipGruppi.id_impiegato " & _
			  "WHERE (id_admin IN (SELECT al_utente_id FROM al_attivita_utenti " & _
					   		   	  "WHERE al_tipo_id="& cIntero(request("ID")) &") " & _
			  "OR id_admin IN (SELECT id_impiegato FROM al_attivita_gruppi t INNER JOIN " & _
					   		  "tb_rel_dipGruppi r ON t.al_gruppo_id=r.id_gruppo " & _
							  "WHERE al_tipo_id="& cIntero(request("ID")) &")) " & _
			  "AND (id_admin NOT IN (SELECT al_utente_id FROM al_documenti_utenti c INNER JOIN " & _
				  				 	"tb_allegati d ON c.al_tipo_id=d.all_documento_id " & _
				  		   		 	"WHERE all_attivita_id="& cIntero(request("ID")) &") " & _
		 	  "AND id_admin NOT IN (SELECT id_impiegato FROM (al_documenti_gruppi a INNER JOIN " & _
				 				   "tb_rel_dipGruppi b ON a.al_gruppo_id=b.id_gruppo) INNER JOIN " & _
								   "tb_allegati e ON a.al_tipo_id=e.all_documento_id " & _
								   "WHERE all_attivita_id="& cIntero(request("ID")) &")) " & _
			  "OR (id_admin NOT IN (SELECT al_utente_id FROM al_documenti_utenti e INNER JOIN " & _
				  				   "tb_allegati f ON e.al_tipo_id=f.all_documento_id " & _
				  		   		   "WHERE all_attivita_id="& cIntero(request("ID")) &") " & _
		 	  "AND id_admin NOT IN (SELECT DISTINCT id_impiegato FROM ((al_documenti_gruppi g INNER JOIN " & _
				 				   "tb_rel_dipGruppi h ON g.al_gruppo_id=h.id_gruppo) INNER JOIN " & _
								   "tb_allegati i ON g.al_tipo_id=i.all_documento_id) " & _
								   "WHERE all_attivita_id="& cIntero(request("ID")) &") "& _
			  "AND "& SQL_IsTrue(conn, "(SELECT att_pubblica FROM tb_attivita WHERE att_id="& cIntero(request("ID")) &")") &")"
			utentiOrbi = GetValueList(conn, rsp, sql)
			'utentiOrbi: se "" non ci sono documenti invisibili else bisogna controllare che i doc non siano pubblici
			if utentiOrbi <> "" then
				sql = "SELECT doc_id, doc_nome FROM tb_documenti d INNER JOIN tb_allegati a "& _
					  "ON d.doc_id=a.all_documento_id "& _
					  "WHERE doc_id NOT IN (SELECT al_tipo_id FROM al_documenti_utenti "& _
					  				       "WHERE al_utente_id IN ("& utentiOrbi &")) "& _
					  "AND doc_id NOT IN (SELECT al_tipo_id FROM al_documenti_gruppi aldg INNER JOIN "& _
					  					 "tb_rel_dipgruppi r ON aldg.al_gruppo_id=r.id_gruppo "& _
										 "WHERE id_impiegato IN ("& utentiOrbi &")) "& _
					  "AND all_attivita_id="& cIntero(request("ID")) &" AND ("& AL_query(conn, AL_DOCUMENTI) & _
					  " OR doc_creatore_id="& Session("ID_ADMIN") &") AND NOT "& SQL_IsTrue(conn, "doc_pubblica")
				rsp.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdtext
				if not rsp.eof then%>
					<tr><th class="error" colspan="4">ATTENZIONE: Allegati non visibili da alcuni destinatari</th></tr>
					<tr>
						<td colspan="4">
							<table cellpadding="0" cellspacing="1" width="100%">
								<tr>
									<td class="note" colspan="2">
										E' possibile salvare comunque, lasciando inalterati i permessi dei documenti interessati.
									</td>
								</tr>
								<%while not rsp.eof%>
									<tr>
										<td class="content"><%= rsp("doc_nome") %></td>
										<td width="24%" class="content_center">
											<a class="button_L2" href="javascript:void(0);" onclick="OpenPositionedScrollWindow('AccessList.asp?ctrl=si&ID=<%= rsp("doc_id") %>&tipo=DOCUMENTI', 'al', 50, 50, 700, 300, true)">
												MODIFICA PERMESSI DOCUMENTO
											</a>
										</td>
									</tr>
									<%rsp.movenext
								wend %>
							</table>
						</td>
					</tr>
				<%End If 
				rsp.close
			end if
		end if			'fine: se non e' conclusa%>
		<% if conclusa="" then %>
			<tr><th colspan="4">DESTINATARI DELL'ATTIVIT&Agrave;</th></tr>
			<tr>
				<td class="content" colspan="4" style="padding-bottom:1px;">
					<a href="javascript:void(0);" class="button_form" onclick="OpenAutoPositionedScrollWindow('AccessList.asp?ctrl=si&ID=<%= request("ID") %>&TIPO=ATTIVITA', 'AL', 700, 240, true);">
						MODIFICA DESTINATARI
					</a>
				</td>
			</tr>
		<% end if %>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
			<textarea style="width:100%;" rows="4" name="tft_att_note" <%= IIF(rs("att_conclusa"), " disabled ", "") %>><%= rs("att_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				<%If rs("att_conclusa") AND NOT rs("att_inSospeso") AND Session("COM_ADMIN") = "" then %>
					<a href="Attivita.asp" class="button">
						INDIETRO
					</a>
				<% Else %>
					(*) Campi obbligatori.
					<% If NOT rs("att_inSospeso") then 
						if cInteger(rs("RISPOSTA")) = 0 then%>
							<input type="button" class="button" name="rispondi" value="RISPONDI" onclick="document.location='AttivitaNew.asp?DOM_ID=<%= request("ID") %>'">
						<% end if
					else %>
						<input type="submit" class="button" name="mod" value="SALVA COME BOZZA">
					<% End If %>
					<input type="submit" class="button" name="mod" value="SALVA">
				<% End If %>
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<%
	set rs = nothing
	conn.close
	set conn = nothing
%>
