<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_immaginiFormati_accesso, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoImmaginiFormatiSalva.asp")
end if

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(1)
dicitura.sezione = "Gestione siti - formati immagini - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoImmaginiFormati.asp"
dicitura.sottosezioni(1) = "ELENCO FILES"
dicitura.links(1) = "SitoFileManager.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, rsr, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")
set rsr = server.CreateObject("ADODB.recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("WEB_IMMAGINIFORMATI_SQL"), "imf_id", "SitoImmaginiFormatiMod.asp")
end if

sql = "SELECT * FROM tb_immaginiformati WHERE imf_id="& cIntero(request("ID"))
set rs = conn.Execute(sql)
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati del formato immagine</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="plugin precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="plugin successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">DATI DEL FORMATO</th></tr>
		<tr>
			<td class="label">nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_imf_nome" value="<%= rs("imf_nome") %>" maxlength="255" size="80">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">directory:</td>
			<td class="content" colspan="3">
				<% CALL WriteFileSystemPicker_Input(Session("AZ_ID"), FILE_SYSTEM_DIRECTORY, "", "", "form1", "tft_imf_dir", rs("imf_dir"), "", false, false) %>
			</td>
		</tr>
		<tr>
			<td class="label">suffisso:</td>
			<td class="content">
				<input type="text" class="text" name="tft_imf_suffisso" value="<%= rs("imf_suffisso") %>" maxlength="50" size="10">
			</td>
			<td class="label">suffisso formato:</td>
			<td class="content">
				<input type="checkbox" name="chk_imf_suffissoFormato" value="1" class="checkbox" <%= Chk(rs("imf_suffissoFormato")) %>>
			</td>
		</tr>
		<tr>
			<td class="label">salva file originale:</td>
			<td class="content">
				<input type="checkbox" name="chk_imf_salvaOriginale" value="1" class="checkbox" <%= Chk(rs("imf_salvaOriginale") <> "") %>>
			</td>
			<td class="note" colspan="2">
				Salva una copia del file caricato dall'utente con lo stesso nome del file generato aggiunto di "_originale".
			</td>
		</tr>
		<tr><th colspan="4">DIMENSIONI</th></tr>
		<tr>
			<td class="label">larghezza:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_imf_width" value="<%= rs("imf_width") %>" maxlength="10" size="4">
			</td>
			<td class="label">altezza:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_imf_height" value="<%= rs("imf_height") %>" maxlength="10" size="4">
			</td>
		</tr>
		<tr>
			<td class="label" style="white-space: nowrap;">dimensioni massime:</td>
			<td class="content">
				<input type="checkbox" name="chk_imf_dimensioniMax" value="1" class="checkbox" <%= Chk(rs("imf_dimensioniMax")) %>>
			</td>
			<td class="note" colspan="2">
				Deselezionando il flag le immagini avranno sempre le dimensioni impostate,
				selezionandolo le immagini verrano ridimensionate solamente se superano le dimensioni impostate.
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="mod" value="SALVA">
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
set rsr = nothing
conn.Close
set conn = nothing
%>