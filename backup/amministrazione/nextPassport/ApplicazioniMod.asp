<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
CALL CheckAutentication(session("PASS_ADMIN") <> "")

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ApplicazioniSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim i, conn, rs, rsr, sql, lock
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("SQL_APPLICAZIONI"), "id_sito", "ApplicazioniMod.asp")
end if

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione applicazioni - modifica"
dicitura.puls_new = "INDIETRO;ACCESSI;PARAMETRI;TABELLE DATI"
dicitura.link_new = "Applicazioni.asp;ApplicazioniAccessi.asp?ID=" & request("ID") & ";ApplicazioniParamsModifica.asp?ID=" & request("ID") & ";ApplicazioniTabelle.asp?ID=" & request("ID")
dicitura.scrivi_con_sottosez()


sql = "SELECT * FROM tb_siti WHERE id_sito=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<% if not rs("sito_amministrazione") then %>
		<input type="hidden" name="tfn_sito_rubrica_area_riservata" value="<%= rs("sito_rubrica_area_riservata") %>">
	<% end if %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica applicazione</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="applicazione precedente">
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="applicazione successiva">
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="2">DEFINIZIONE DELL'APPLICAZIONE</th></tr>
		<tr>
			<td class="label" style="width:20%;">ID Applicazione:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_id_sito" value="<%= rs("id_sito") %>" maxlength="3" size="3">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">Nome:</td>
			<td class="content">
				<img src="../grafica/flag_mini_it.jpg" alt="italiano" border="0">
				<input type="text" class="text" name="tft_sito_nome" value="<%= rs("sito_nome") %>" maxlength="250" size="75">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">Nome inglese:</td>
			<td class="content">
				<img src="../grafica/flag_mini_en.jpg" alt="inglese" border="0">
				<input type="text" class="text" name="tft_sito_nome_en" value="<%= rs("sito_nome_en") %>" maxlength="250" size="75">
			</td>
		</tr>

		<% if not rs("sito_amministrazione") then %>
			<tr>
				<td class="label" style="width:20%;">Tipo applicazione:</td>
				<td class="label" style="width:80%;">pubblica</td>
			</tr>			
			<input type="hidden" name="sito_amministrazione" value="">
		<% else %>
			<tr>
				<td class="label" style="width:20%;">Tipo applicazione:</td>
				<td class="label" style="width:80%;">amministrativa</td>
			</tr>		
			<input type="hidden" name="sito_amministrazione" value="1">
			<tr>
				<td class="label">Percorso applicazione:</td>
				<td class="content">
					<input type="text" class="text" name="tft_sito_dir" value="<%= rs("sito_dir") %>" maxlength="150" size="75">
					<span id="path"></span>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label" style="width:20%;">Protetta:</td>
			<td class="content">
				<input type="checkbox" class="noBorder" name="chk_sito_protetto" value="1" <%= chk(cBoolean(rs("sito_protetto"),false))%>>
			</td>
		</tr>
		<tr><th colspan="2">DEFINIZIONE PROFILI UTENTE</th></tr>
		<% for i=1 to 9 %>
			<tr>
				<td class="label">Permesso <%=i%>:</td>
				<td class="content">
					<input type="hidden" name="old_value_sito_p<%=i%>" value="<%= rs("sito_p" & i) %>">
					<input type="text" class="text" name="tft_sito_p<%=i%>" value="<%= rs("sito_p" & i) %>" maxlength="50" size="40">
					<%if i=1 then%>(*)<%end if%>
				</td>
			</tr>
		<% next %>
		<tr><th colspan="2">GESTIONE PERMESSI ESTERNI AGGIUNTIVI</th></tr>
		<tr>
			<td class="label" nowrap>da scheda utente:</td>
			<td class="content">
				<input type="text" class="text" name="tft_sito_prmEsterni_admin" value="<%= rs("sito_prmEsterni_admin") %>" maxlength="250" size="75">
			</td>
		</tr>
		<tr>
			<td class="label">da scheda applicazione</td>
			<td class="content">
				<input type="text" class="text" name="tft_sito_prmEsterni_sito" value="<%= rs("sito_prmEsterni_sito") %>" maxlength="250" size="75">
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

<% rs.close
conn.close
set rs = nothing
set rsr = nothing
set conn = nothing%>