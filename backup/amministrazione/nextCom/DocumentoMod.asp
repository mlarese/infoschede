<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_DocumentiFiles.asp" -->
<%
if request.form("mod") <> "" then
	Server.Execute("DocumentoSalva.asp")
end if

dim conn, sql, rs, rsd

set conn = Server.CreateObject("ADODB.Connection")
set rs = server.CreateObject("ADODB.RecordSet")
set rsd = server.CreateObject("ADODB.RecordSet")
conn.open Application("DATA_ConnectionString")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session(Session("DOC_PREFIX") & "SQL_DOCUMENTI"), "doc_id", "DocumentoMod.asp")
end if


'controllo accesso
if NOT AL(conn, request("ID"), AL_DOCUMENTI) then
	response.redirect "documenti.asp"
end if

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = "Documenti - modifica"
'Indirizzo pagina per link su sezione 
		HREF = "Documenti.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************

sql = "SELECT * FROM tb_documenti WHERE doc_id="& cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText

'permessi modifica
dim disabled, FileList, Tipologia, i
if Session("COM_ADMIN") = "" AND Session("COM_POWER") = "" AND Session("ID_ADMIN") <> rs("doc_creatore_id") then
	disabled = "disabled"
end if

%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfd_doc_mod_data" value="NOW">
	<input type="hidden" name="tfn_doc_mod_utente" value="<%= Session("ID_ADMIN") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<% if Session(Session("DOC_PREFIX") & "SQL_DOCUMENTI")<>"" then 
				'verifica se esiste elenco documenti%>
				<table border="0" cellspacing="0" cellpadding="0" align="right">
					<tr>
						<td style="font-size: 1px; padding-right:1px;" nowrap>
							<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="documento precedente">
								&lt;&lt; PRECEDENTE
							</a>
							&nbsp;
							<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="documento successivo">
								SUCCESSIVO &gt;&gt;
							</a>
						</td>
					</tr>
				</table>
			<% end if %>
			Vedi / Modifica documento
		</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<% if Session("DOC_PRA_ID")<>"" then
			CALL SelezionaPratica(conn, rsd, "DOC", rs("doc_pratica_id"), false) 
		else
			CALL SelezionaPratica(conn, rsd, "DOC", rs("doc_pratica_id"), disabled="") 
		end if
		%>
		<tr>
			<td class="label">nome:</td>
			<td class="content" colspan="3">
				<input <%= disabled %> type="text" class="text" name="tft_doc_nome" value="<%= rs("doc_nome") %>" maxlength="255" size="75">
				<span id="nome">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">files:</td>
			<%if disabled = "" then
				'modifica dei files associati%>
				<td class="content" colspan="3">
					<%CALL GestioneFilesAssociati(conn, rsd, request("ID"))%>
				</td>
			<% else %>
				<td colspan="3">
					<%CALL ElencoFileAssociati(conn, rsd, request("ID"))%>
				</td>
			<%end if%>
		</tr>
		<%if request.serverVariables("request_method") = "POST" then
			tipologia = request.form("tfn_doc_tipologia_id")
		else
			tipologia = rs("doc_tipologia_id")
		end if %>
		<tr><th colspan="4">DESCRIZIONE DOCUMENTO</th></tr>
		<tr>
			<td class="label">tipologia:</td>
			<td class="content" colspan="3">
				<% dropDown conn, "SELECT * FROM tb_tipologie", "tipo_id", "tipo_nome", "tfn_doc_tipologia_id", tipologia, false, IIF(disabled = "", "onchange='form1.submit()'", disabled), LINGUA_ITALIANO %>
				<span id="tipologia">(*)</span>
			</td>
		</tr>
		<tr><th class="L2" colspan="4">DESCRITTORI</th></tr>
		<% sql = " SELECT * FROM (rel_tipologie_descrittori t INNER JOIN tb_descrittori d ON t.rtd_descrittore_id = d.descr_id) " & _
				 " LEFT JOIN rel_documenti_descrittori r ON (d.descr_id = r.rdd_descrittore_id AND r.rdd_documento_id="& cIntero(request("ID")) &") "& _
				 " WHERE rtd_tipologia_id="& tipologia & _
				 " ORDER BY descr_ordine, descr_nome"
		CALL DesElenco(conn, sql, "tb_descrittori", "descr_id", "descr_nome", "descr_tipo", "", "rdd_valore", false, 4)
		If disabled = "" then %>
			<tr><th colspan="4">PERMESSI DEL DOCUMENTO</th></tr>
			<tr>
				<td class="content" colspan="4">
					<table width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td width="30%">
								<a href="javascript:void(0);" class="button_form" onclick="OpenAutoPositionedScrollWindow('AccessList.asp?ctrl=si&ID=<%= request("ID") %>&TIPO=DOCUMENTI', 'AL', 700, 240, true);">
									MODIFICA PERMESSI DOCUMENTO
								</a>
							</td>
							<td class="content notes">
								ATTENZIONE: Se il documento &egrave; allegato ad una o pi&ugrave; pratiche, cambiandone i permessi alcuni utenti potrebbero non essere pi&ugrave; abilitati ad utilizzarlo.
							</td>
						</tr>
					</table>
				</td>
			</tr>
		<% 	End If %>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea <%= disabled %> style="width:100%;" rows="7" name="tft_doc_note"><%= rs("doc_note")%></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				<% If disabled = "" then %>
					(*) Campi obbligatori.
					<input type="submit" class="button" name="mod" value="SALVA">
				<% Else %>
					<a href="Documenti.asp" class="button">
						INDIETRO
					</a>
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
rs.close
conn.close
set rs = nothing
set rsd = nothing
set conn = nothing
%>
