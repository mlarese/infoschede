<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("PraticaSalva.asp")
end if

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.Recordset")
conn.open Application("DATA_ConnectionString")


'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	if Session("PRA_CONTATTO_ID")<>"" then
		Titolo_sezione = "Anagrafica contatti - nuova pratica"
	else
		Titolo_sezione = "Pratiche - nuova"
	end if
'Indirizzo pagina per link su sezione 
HREF = "Pratiche.asp"
'Azione sul link: {BACK | NEW}
Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
		<caption>Inserimento nuova pratica</caption>
		<tr><th colspan="4">DATI PRATICA</th></tr>
		<% if Session("PRA_CONTATTO_ID") = "" then 
			'sezione "stand alone" disponibile per tutti i contatti%>
			<tr>
				<td class="label" rowspan="2">contatto:</td>
				<td class="content" colspan="3">
                    <% CALL WriteContactPicker_Input(conn, rs, "", "", "form1", "contatti", request("contatti"), "", true, true, false, "") %>
				</td>
			</tr>
			<tr>
				<td class="content notes" colspan="4">
					Selezionando pi&ugrave; contatti verr&agrave; creata una pratica per ciascuno di essi con le stesso nome ed eventuali note.
					In caso si inserisca anche la prima attivit&agrave; ne verr&agrave; generata una con le caratteristiche indicate per ogni pratica.
				</td>
			</tr>
		<% Else %>
			<tr>
				<td class="label">contatto:</td>
				<td class="content" colspan="3">
					<% sql = "SELECT IDElencoIndirizzi, isSocieta, NomeOrganizzazioneElencoIndirizzi, NomeElencoIndirizzi, CognomeElencoIndirizzi " & _
							 " FROM tb_Indirizzario WHERE IDElencoIndirizzi=" & Session("PRA_CONTATTO_ID")
					rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText %>
					<% ContactLinkedName(rs) %>
					<% rs.close %>
				</td>
			</tr>
			<input type="hidden" name="contatti" value="<%= Session("PRA_CONTATTO_ID") %>;">
		<% End If %>
		<% If Application("NextCom_codice") = "" then %>
		<tr>
			<td class="label">codice:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_pra_codice" value="<%= request("tft_pra_codice") %>" maxlength="50" size="50">
				<span id="codice">(*)</span>
			</td>
		</tr>
		<% Else  %>
		<input type="Hidden" name="tft_pra_codice" value=" ">
		<% End If %>
		<tr>
			<td class="label">nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_pra_nome" value="<%= request("tft_pra_nome") %>" maxlength="255" size="100">
				<span id="codice">(*)</span>
			</td>
		</tr>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="7" name="tft_pra_note"><%=request("tft_pra_note")%></textarea>
			</td>
		</tr>
		<input type="hidden" name="tfd_att_dataCrea" value="<%= Date %>">
		<input type="hidden" name="tfn_att_mittente_id" value="<%= Session("ID_ADMIN") %>">
		<input type="hidden" name="tfn_att_sistema" value="0">

		<tr><th colspan="4">PRIMA ATTIVIT&Agrave; DELLA PRATICA</th></tr>
		<tr><th class="L2" colspan="4">DATI ATTIVIT&Agrave;</th></tr>
		<tr>	
			<td colspan=4" class="content notes">
				Impostando almeno l'oggetto ed il testo dell'attivit&agrave; ne verr&agrave; creata una per ogni pratica generata.
			</td>
		</tr>
		<tr>
			<td class="label">oggetto:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_att_oggetto" value="<%= request("tft_att_oggetto") %>" maxlength="255" size="75">
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
		<tr><th class="L2" colspan="4">TESTO</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="7" name="tft_att_testo"><%=request("tft_att_testo")%></textarea>
			</td>
		</tr>
		<tr><th class="L2" colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="4" name="tft_att_note"><%=request("tft_att_note")%></textarea>
			</td>
		</tr>
		<tr><th colspan="4">PERMESSI DI BASE DELLA PRATICA</th></tr>
		<tr>
			<td colspan="4">
				<% CALL AL_disegna(conn, "", AL_PRATICHE)%>
			</td>
		</tr>
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

<%
conn.close
set rs = nothing
set conn = nothing
%>
