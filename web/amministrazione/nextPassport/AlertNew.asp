<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
'controllo accesso
CALL CheckAutentication(session("PASS_ADMIN") <> "")


if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("AlertSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="Tools_Passport.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione alert - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Alert.asp"
dicitura.scrivi_con_sottosez()

dim i, conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo alert</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label">nome:</td>
			<td class="content" colspan="4">
				<input type="text" class="text" name="tft_sev_nome_IT" value="<%= Server.HtmlEncode(cString(request.form("tft_sev_nome_IT"))) %>" maxlength="250" size="75">
			</td>
		</tr>
		<tr>
			<td class="label">codice:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_sev_codice" value="<%= request("tft_sev_codice") %>" maxlength="50" size="50">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">applicazione:</td>
			<td class="content" colspan="3">
				<%	sql = "SELECT * FROM tb_siti WHERE " & SQL_IsTrue(conn, "sito_Amministrazione") & " ORDER BY sito_nome"
					CALL DropDown(conn, sql, "id_sito", "sito_nome", "tfn_sev_sito_id", request.form("tfn_sev_sito_id"), true, " style=""width:100%;""", LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">abilitato:</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="chk_sev_abilitato" value="1" <%= Chk(request.form("chk_sev_abilitato") <> "") %>>
			</td>
			<td class="label">multisito:</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="chk_sev_multisito" value="1" <%= Chk(request.form("chk_sev_multisito") <> "") %>>
			</td>
		</tr>
		<tr><th colspan="4">CONFIGURAZIONI</th></tr>
		<tr>
			<td colspan="4">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<td class="label_no_width">
							L'inserimento delle configurazioni &egrave; possibile dopo aver salvato.
						</td>
					</tr>
				</table>
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
<% conn.close
set conn = nothing
%>