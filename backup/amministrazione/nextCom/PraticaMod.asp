<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<%
dim conn, rs, sql

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.Recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session(Session("PRA_PREFIX") & "SQL_PRATICHE"), "pra_id", "PraticaMod.asp")
end if

'controllo accesso
if NOT AL(conn, request("ID"), AL_PRATICHE) then
	response.redirect "Pratiche.asp"
end if

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("PraticaSalva.asp")
end if

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	if Session("PRA_CONTATTO_ID")<>"" then
		Titolo_sezione = "Anagrafica contatti - modifica pratica"
	else
		Titolo_sezione = "Pratiche - modifica"
	end if
'Indirizzo pagina per link su sezione 
	HREF = "Pratiche.asp;Attivita.asp?PRA_ID=" & request("ID") & ";Documenti.asp?PRA_ID=" & request("ID")
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO;ATTIVITA' PRATICA;DOCUMENTI PRATICA"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
sql = "SELECT tb_pratiche.*, IDElencoIndirizzi, isSocieta, NomeOrganizzazioneElencoIndirizzi, NomeElencoIndirizzi, " & _
	  " CognomeElencoIndirizzi, tb_admin.admin_nome, tb_admin.admin_cognome " & _
	  " FROM (tb_pratiche INNER JOIN tb_Indirizzario ON tb_pratiche.pra_cliente_id=tb_indirizzario.IDElencoIndirizzi) " & _
	  "	INNER JOIN tb_admin ON tb_pratiche.pra_creatore_id = tb_admin.id_admin " & _
	  " WHERE pra_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText

'permessi modifica
dim disabled
if Session("COM_ADMIN") = "" AND Session("COM_POWER") = "" AND Session("ID_ADMIN") <> rs("pra_creatore_id") then
	disabled = "disabled"
end if
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<% if Session(Session("PRA_PREFIX") & "SQL_PRATICHE")<>"" then 
				'verifica se esiste elenco pratiche%>
				<table border="0" cellspacing="0" cellpadding="0" align="right">
					<tr>
						<td style="font-size: 1px; padding-right:1px;" nowrap>
							<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="pratica precedente">
								&lt;&lt; PRECEDENTE
							</a>
							&nbsp;
							<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="pratica successiva">
								SUCCESSIVA &gt;&gt;
							</a>
						</td>
					</tr>
				</table>
			<% end if %>
			Modifica pratica
		</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label">contatto:</td>
			<td class="content" colspan="3">
				<% ContactLinkedName(rs) %>
			</td>
		</tr>
		<tr>
			<td class="label">codice:</td>
			<td class="content_b" colspan="3">
				<%= rs("pra_codice") %>
			</td>
		</tr>
		<tr>
			<td class="label">creatore:</td>
			<td class="content"><%= rs("admin_cognome") %>&nbsp;<%= rs("admin_nome") %></td>
			<td class="label">data creazione:</td>
			<td class="content"><%= DateTimeITA(rs("pra_dataI")) %></td>
		</tr>
		<tr>
			<td class="label">nome:</td>
			<td class="content" colspan="3">
				<input <%= disabled %> type="text" class="text" name="tft_pra_nome" value="<%= rs("pra_nome") %>" maxlength="255" size="100">
				<span id="codice">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">archiviata:</td>
			<td class="content" colspan="3">
				<input <%= disabled %> type="Checkbox" name="chk_pra_archiviata" value="true" class="noborder" <%= Chk(rs("pra_archiviata")) %>>
			</td>
		</tr>
		<% 	If disabled = "" then %>
			<tr><th colspan="4">PERMESSI DI DEFAULT</th></tr>
			<tr>
				<td class="content" colspan="4">
					<table width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td width="24%">
								<a href="javascript:void(0);" class="button_form" onclick="OpenAutoPositionedScrollWindow('AccessList.asp?ID=<%= request("ID") %>&TIPO=DEFAULT', 'AL', 700, 240, true);">
									MODIFICA PERMESSI DI BASE
								</a>
							</td>
							<td class="content notes">
								ATTENZIONE: Cambiando l'access list di default della pratica cambiano anche i permessi delle attivit&agrave;
								e dei documenti contenuti e che ereditano il comportamento dalla pratica.
							</td>
						</tr>
					</table>
				</td>
			</tr>
		<% 	End If %>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea <%= disabled %> style="width:100%;" rows="7" name="tft_pra_note"><%= rs("pra_note")%></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<% 	If disabled = "" then %>
				<input type="submit" class="button" name="mod" value="SALVA">
				<% 	Else %>
					<a href="Pratiche.asp" class="button">
						INDIETRO
					</a>
				<% 	End If %>
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
