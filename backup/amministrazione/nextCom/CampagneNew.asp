<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<%
'controllo accesso
if Session("COM_ADMIN")="" AND Session("COM_POWER")="" then
	response.redirect "Contatti.asp"
end if
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("CampagneSalva.asp")
end if

dim conn, rs, sql

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")


'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action, i
'Titolo della pagina
	Titolo_sezione = "Campagne marketing - nuova"
'Indirizzo pagina per link su sezione 
		HREF = "Campagne.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfd_inc_insData" value="<%=Now()%>">
	<input type="hidden" name="tfn_inc_insAdmin_id" value="<%=Session("ID_ADMIN")%>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova campagna marketing</caption>
		<tr><th colspan="2">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label" style="width:22%;">nome campagna:</td>
			<td class="content">
				<input type="text" class="text" name="tft_inc_nome" value="<%= request("tft_inc_nome") %>" maxlength="250" size="70">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<tr><th colspan="2">CONTATTI ASSOCIATI</th></tr>
		<tr>
			<td class="label">singoli contatti:</td>
			<td class="content">
                <% CALL WriteContactPicker_Input(conn, rs, "", "", "form1", "contatti", request("contatti"), "", true, false, false, "") %>
			</td>
		</tr>
		<tr><th colspan="2">NOTE</th></tr>
		<tr>
			<td class="content" colspan="2">
				<textarea style="width:100%;" rows="5" name="tft_inc_note"><%=request("tft_inc_note")%></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="2">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva_avanti" value="SALVA">
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
set conn = nothing%>