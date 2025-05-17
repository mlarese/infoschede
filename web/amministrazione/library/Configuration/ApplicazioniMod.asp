<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../../nextPassport/ToolsApplicazioni.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ApplicazioniSalva.asp")
end if
%>
<%
dim i, conn, rs, rsr, sql, lock
set conn = Server.CreateObject("ADODB.Connection")
conn.open GetConfigurationConnectionstring()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("SQL_APPLICAZIONI"), "id_sito", "ApplicazioniMod.asp")
end if

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione applicazioni - modifica"
dicitura.puls_new = "INDIETRO;PARAMETRI;TABELLE DATI"
dicitura.link_new = "Applicazioni.asp;ApplicazioniParamsModifica.asp?ID=" & request("ID") & ";ApplicazioniTabelle.asp?ID=" & request("ID")
dicitura.scrivi_con_sottosez()


sql = "SELECT * FROM tb_siti WHERE id_sito=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
%>
<script language="JavaScript" type="text/javascript">
	
	function show_mandatory(){
		var isAmministrazione = document.form1.sito_amministrazione;
		var span_path = document.getElementById('path')
		var input_path = document.form1.tft_sito_dir;

		if (isAmministrazione.type=="hidden")
			var check = true;
		else
			var check = isAmministrazione.checked
		
		if (check){
			span_path.innerHTML='(*)';
			input_path.disabled=false;
		}
		else{
			span_path.innerHTML='';
			input_path.disabled=true;
		}
		
	}
</script>
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
				<input type="text" class="text" name="tft_sito_nome" value="<%= rs("sito_nome") %>" maxlength="250" size="75">
				(*)
			</td>
		</tr>
		<%
		'if cInteger(rs("sito_rubrica_area_riservata"))>0 then
		'	sql = "SELECT COUNT(*) FROM rel_rub_ind WHERE id_rubrica=" & cInteger(rs("sito_rubrica_area_riservata"))
		'	if cInteger(GetValueList(conn, rsr, sql))=0 then 
		'		lock ="" 
		'	else
				lock = " disabled "
		'	end if
		'else
		'	lock = ""
		'end if
		%>
		<% if cInteger(Application("NextPassport_GruppoLavoroAreaRiservata"))>0 OR cInteger(Session("GruppoLavoroAreaRiservata"))>0 then %>
			<tr>
				<td class="label">Area amministrativa:</td>
				<td class="content">
					<input <%= lock %> type="checkbox" class="noBorder" name="sito_amministrazione" value="1" <% if rs("sito_amministrazione") then %> checked <% end if %> onClick="show_mandatory()">
					<% if lock<>"" then %>
						La caratteristica non &egrave; modificabile perch&egrave; sono presenti utenti che hanno accesso all'applicazione.
					<% end if %>
				</td>
			</tr>
		<% else %>
			<input type="hidden" name="sito_amministrazione" value="1">
		<% end if %>
		<tr>
			<td class="label">Percorso applicazione:</td>
			<td class="content">
				<input type="text" class="text" name="tft_sito_dir" value="<%= rs("sito_dir") %>" maxlength="150" size="75">
				<span id="path"></span>
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
<script language="JavaScript" type="text/javascript">
	show_mandatory()
</script>
</html>

<% rs.close
conn.close
set rs = nothing
set rsr = nothing
set conn = nothing%>