<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_Attivita.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" AND request("salva")<>"" then
	Server.Execute("AttivitaSalva.asp")
end if

dim conn, rs, rsd, sql, value
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.createobject("adodb.recordset")
set rsd = server.createobject("adodb.recordset")

if request("DOM_ID")<>"" then
	'apre recordset su domanda: Azione eseguita --> rispondi all'attivita'
	sql = " SELECT * FROM tb_attivita INNER JOIN tb_admin ON tb_attivita.att_mittente_id = tb_admin.id_admin " & _
		  " WHERE att_id=" & cIntero(request("DOM_ID"))
	rsd.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
end if

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
if Session("ATT_PRA_ID")<>"" then
	Titolo_sezione = "Pratiche - attivit&agrave; della pratica - nuova"
elseif Session("ATT_DOC_ID")<>"" then
	'attivita' collegate al documento
	Titolo_sezione = "Documenti - attivit&agrave; collegate al documento - nuova"
elseif rsd.state = adStateOpen then
	'azione: rispondi all'attivita'
	Titolo_sezione = "Attivit&agrave; - risposta all'attivit&agrave; &quot;" & rsd("att_oggetto") & "&quot;"
else
	Titolo_sezione = "Attivit&agrave; - nuova"
end if
'Indirizzo pagina per link su sezione 
		HREF = "Attivita.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>
<div id="content">
	<form action="<%= GetCurrentBaseUrl() %>/AttivitaNew.asp" method="post" id="form1" name="form1">
	<input type="hidden" name="tfd_att_dataCrea" value="NOW">
	<input type="hidden" name="tfn_att_mittente_id" value="<%= Session("ID_ADMIN") %>">
	<input type="hidden" name="tfn_att_sistema" value="0">
	<% if rsd.State = adStateOpen then
		'azione: rispondi all'attivita'%>
		<input type="hidden" name="tfn_att_domanda_id" value="<%= rsd("att_id") %>">
	<% else %>
		<input type="hidden" name="tfn_att_domanda_id" value="0">
	<% end if %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova attivit&agrave;</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<% if Session("ATT_PRA_ID")<>"" then
			CALL SelezionaPratica(conn, rs, "ATT", Session("ATT_PRA_ID"), false) 
		elseif rsd.State = adStateOpen then
			'azione: rispondi all'attivita'
			CALL SelezionaPratica(conn, rs, "ATT", rsd("att_pratica_id"), false) 
		else	
			CALL SelezionaPratica(conn, rs, "ATT", request("tfn_att_pratica_id"), request("ID")="") 
		end if %>
		<tr>
			<td class="label">oggetto:</td>
			<td class="content" colspan="3">
				<% If Request.ServerVariables("REQUEST_METHOD")="POST" then 
					'salvataggio non andato a buon fine
					value = request("tft_att_oggetto")
				elseif rsd.State = adStateOpen then
					'risposta ad un'altra attivita
					if instr(1, rsd("att_oggetto"), "Re: ", vbTextCompare)<1 then
						value = "Re: " & rsd("att_oggetto")
					else
						value = rsd("att_oggetto")
					end if
				else
					value = ""
				end if%>
				<input type="text" class="text" name="tft_att_oggetto" value="<%= value %>" maxlength="255" size="75">
				<span id="oggetto">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">data scadenza:</td>
			<td class="content" colspan="3">
				<% CALL WriteDataPicker_Input("form1", "tfd_att_dataS", request("tfd_att_dataS"), "", "/", true, true, LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">prioritaria:</td>
			<td class="content">
				<input type="Checkbox" name="chk_att_priorita" value="1" class="noborder" <%= Chk(request("chk_att_priorita")<>"") %>>
			</td>
			<td class="label">conclusa:</td>
			<td class="content">
				<input type="Checkbox" name="chk_att_conclusa" value="1" class="noborder" <%= Chk(request("chk_att_conclusa")<>"") %>>
			</td>
		</tr>
		<tr><th colspan="4">TESTO (*)</th></tr>
		<tr>
			<td class="content" colspan="4">
				<% If Request.ServerVariables("REQUEST_METHOD")="POST" then 
					'salvataggio non andato a buon fine
					value = request("tft_att_testo")
				elseif rsd.State = adStateOpen then
					'risposta ad un'altra attivita - genera testo del messaggio
					value = "" & vbCRLF & vbCRLF & vbCRLF & _
			  				"--------------------------------------------------------------------------------------------------------------------------------------------" & vbCRLF & _
			  				"data: " & _
								DateTimeITA(rsd("att_dataCrea")) & vbCRLF & _
			  				"da: " & _
			  					rsd("admin_nome") & " " & rsd("admin_cognome") & vbCrLF & _
							"testo: " & _
			  				rsd("att_testo") & vbCRLF
				else
					value = ""
				end if%>
				<textarea style="width:100%;" rows="12" name="tft_att_testo"><%= value %></textarea>
			</td>
		</tr>
		<% if rsd.State = adStateOpen AND Request.ServerVariables("REQUEST_METHOD")<>"POST" then
			'recupera documenti collegati dell'attivita' a cui si risponde
			CALL GestioneDocumentiCollegati(conn, rs, rsd("att_id"))
		else	
			'recupera documenti tramite form
			CALL GestioneDocumentiCollegati(conn, rs, "")
		end if %>
		<tr><th colspan="4">DESTINATARI DELL'ATTIVIT&Agrave;</th></tr>
		<tr>
			<td colspan="4">
				<% if rsd.State = adStateOpen AND Request.ServerVariables("REQUEST_METHOD")<>"POST" then
					'recupera permessi dell'attivita alla quale si risponde
					CALL AL_disegna(conn, rsd("att_id"), AL_ATTIVITA)
				else
					'recupera permessi da form
					CALL AL_disegna(conn, "", AL_ATTIVITA)
				end if%>
			</td>
		</tr>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="4" name="tft_att_note"><%=request("tft_att_note")%></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				<table cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td valign="top">(*) Campi obbligatori.</td>
					<td align="right">
						<%' se cambi nome pulsante vedi AttivitaSalva.asp "flag bozza" %>
						(La bozza &egrave; visibile solo a chi la crea.)
						<input type="submit" class="button" name="salva" value="SALVA COME BOZZA">
						<input type="submit" class="button" name="salva" value="SALVA">
					</td>
				</tr>
				</table>
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
set rs = nothing
set rsd = nothing
set conn = nothing
%>
