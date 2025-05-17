<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ApplicazioniSalva.asp")
end if
%>
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione applicazioni - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Applicazioni.asp"
dicitura.scrivi_con_sottosez() 

dim i
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
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova applicazione</caption>
		<tr><th colspan="2">DEFINIZIONE DELL'APPLICAZIONE</th></tr>
		<tr>
			<td class="label" style="width:20%;">ID Applicazione:</td>
			<td class="content">
				<input type="text" class="text" name="tfn_id_sito" value="<%= request("tfn_id_sito") %>" maxlength="3" size="3">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">Nome:</td>
			<td class="content">
				<input type="text" class="text" name="tft_sito_nome" value="<%= request("tft_sito_nome") %>" maxlength="250" size="75">
				(*)
			</td>
		</tr>
		<% if cInteger(Application("NextPassport_GruppoLavoroAreaRiservata"))>0 OR cInteger(Session("GruppoLavoroAreaRiservata"))>0 then %>
			<tr>
				<td class="label">Area amministrativa:</td>
				<td class="content">
					<input checked type="checkbox" class="noBorder" name="sito_amministrazione" value="1" <% if request("sito_amministrazione")<>"" then %> checked <% end if %> onClick="show_mandatory()">
				</td>
			</tr>
		<% else %>
			<input type="hidden" name="sito_amministrazione" value="1">
		<% end if %>
		<tr>
			<td class="label">Percorso applicazione:</td>
			<td class="content">
				<input type="text" class="text" name="tft_sito_dir" value="<%= request("tft_sito_dir") %>" maxlength="150" size="75">
				<span id="path"></span>
			</td>
		</tr>
		<tr><th colspan="2">DEFINIZIONE PROFILI UTENTE</th></tr>
		<% for i=1 to 9 %>
			<tr>
				<td class="label">Permesso <%=i%>:</td>
				<td class="content">
					<input type="text" class="text" name="tft_sito_p<%=i%>" value="<%= request("tft_sito_p" & i) %>" maxlength="50" size="40">
					<%if i=1 then%>(*)<%end if%>
				</td>
			</tr>
		<% next %>
		<tr><th colspan="2">GESTIONE PERMESSI ESTERNI AGGIUNTIVI</th></tr>
		<tr>
			<td class="label" nowrap>da scheda utente:</td>
			<td class="content">
				<input type="text" class="text" name="tft_sito_prmEsterni_admin" value="<%= request("tft_sito_prmEsterni_admin") %>" maxlength="250" size="75">
			</td>
		</tr>
		<tr>
			<td class="label">da scheda applicazione</td>
			<td class="content">
				<input type="text" class="text" name="tft_sito_prmEsterni_sito" value="<%= request("tft_sito_prmEsterni_sito") %>" maxlength="250" size="75">
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
