<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("AmministratoriSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="Tools_Passport.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
if session("PASS_ADMIN") <> "" then
	dicitura.iniz_sottosez(3)
	dicitura.sottosezioni(1) = "APPLICAZIONI"
	dicitura.links(1) = "Applicazioni.asp"
	dicitura.sottosezioni(2) = "GRUPPI DI PARAMETRI"
	dicitura.links(2) = "ApplicazioniParamsGruppi.asp"
	dicitura.sottosezioni(3) = "PARAMETRI"
	dicitura.links(3) = "ApplicazioniParams.asp"
	
	dicitura.puls_new = "NUOVA APPLICAZIONE"
	dicitura.link_new = "ApplicazioniNew.asp"
else
	dicitura.iniz_sottosez(0)
end if
dicitura.sezione = "Gestione utenti area amministrativa - nuovo utente"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Amministratori.asp"
dicitura.scrivi_con_sottosez() 

dim i, conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo utente dell'area amministrativa</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label">Cognome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_admin_cognome" value="<%= request("tft_admin_cognome") %>" maxlength="250" size="70">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label" style="width:18%;">Nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_admin_nome" value="<%= request("tft_admin_nome") %>" maxlength="250" size="70">
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="2">Indirizzi email:</td>
			<td class="label_no_width">Email utente</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_admin_email" value="<%= request("tft_admin_email") %>" maxlength="50" size="50">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label_no_width">Email per newsletter</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_admin_email_newsletter" value="<%= request("tft_admin_email_newsletter") %>" maxlength="50" size="50">
				<span class="note"><br>Se non impostata viene usata l'email dell'utente.</span>
			</td>
		</tr>
		<tr>
			<td class="label">Telefono:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_admin_telefono" value="<%= request("tft_admin_telefono") %>" maxlength="250" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">Cellulare:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_admin_cell" value="<%= request("tft_admin_cell") %>" maxlength="250" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">Fax:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_admin_fax" value="<%= request("tft_admin_fax") %>" maxlength="250" size="50">
			</td>
		</tr>
		<tr><th colspan="4">ACCOUNT DI ACCESSO</th></tr>
		<tr>
			<td class="label">Login:</td>
			<td class="content">
				<input type="hidden" name="old_admin_login" value="">
				<input type="text" class="text" name="tft_admin_login" value="<%= request("tft_admin_login") %>" maxlength="50" size="20">
				(*)
			</td>
			<td class="label">Scadenza:</td>
			<td class="content">
				<% CALL WriteDataPicker_Input("form1", "tfd_admin_scadenza", request("tfd_admin_scadenza"), "", "/", true, true, LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">Password:</td>
			<td class="content">
				<input type="password" class="text" name="tft_admin_password" value="<%= request("tft_admin_password") %>" maxlength="50" size="20">
				(*)
			</td>
			<td class="note" colspan="2" rowspan="2" style="width:60%;">
				Per i valori di login e password utilizzare solo caratteri alfanumerici o &quot;_&quot; 
				indifferentemente con lettere minuscole o maiuscole, ma senza spazi bianchi.
				<span style="letter-spacing:2px;">(<%= LOGIN_VALID_CHARSET %>)</span>
			</td>
		</tr>
		<tr>
			<td class="label">Conferma password:</td>
			<td class="content">
				<input type="password" class="text" name="conferma_password" value="<%= request("conferma_password") %>" maxlength="50" size="20">
				(*)
			</td>
		</tr>
		<tr><th colspan="4">PROFILO DI ACCESSO</th></tr>
		<tr>
			<td colspan="4">
				<% CALL write_permessi(conn, rs, true, 0, 3) %>
			</td>
		</tr>
		<tr><th colspan="4">ACCESSO AI FILES</th></tr>
		<tr>
			<td class="label" rowspan="2">Directory di partenza:</td>
			<td class="content" colspan="3">
				<% CALL WriteFileSystemPicker_Input(application("AZ_ID"), FILE_SYSTEM_DIRECTORY, "images", "", "form1", "tft_admin_dir", request("tft_admin_dir"), "width:420px;", false, false) %>
			</td>
		</tr>
		<tr>
			<td class="content notes" colspan="3">
				Permette di limitare la gestione dei files da parte dell'utente all'interno della cartella selezionata.
			</td>
		</tr>
		<% if cInteger(Application("NextCom_DefaultWorkGroup"))=0 then %>
			<tr><th colspan="4">GRUPPO DI LAVORO</th></tr>
			<tr>
				<td colspan="4">
					<% sql = "SELECT id_gruppo, nome_gruppo, (NULL) AS id_rel_dipgruppi FROM tb_gruppi " & _
							 " ORDER BY nome_gruppo"
					CALL Write_Relations_Checker(conn, rs, sql, 2, "id_gruppo", "nome_gruppo", "id_rel_dipgruppi", "gruppi_di_lavoro") %>
				</td>
			</tr>
		<% else %>
			<input type="hidden" name="gruppi_di_lavoro" value="<%= Application("NextCom_DefaultWorkGroup") %>">
		<% end if %>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
                <input type="submit" class="button" name="salva_elenco" value="SALVA & TORNA AD ELENCO">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% conn.close
set rs = nothing
set conn = nothing
%>