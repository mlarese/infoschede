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
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione applicazioni - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Applicazioni.asp"
dicitura.scrivi_con_sottosez() 

dim i, sql, conn
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
%>
<script language="JavaScript" type="text/javascript">
	
	function setState(){
		var sito_amministrazione = document.getElementById("sito_amministrazione_true");
		EnableIfChecked(sito_amministrazione, document.getElementById("tft_sito_dir"));
		DisableIfChecked(sito_amministrazione, document.getElementById("id_gruppo"));
		
		if (sito_amministrazione.checked){
			document.getElementById('group').innerHTML='';
			document.getElementById('path').innerHTML='(*)';
		}
		else{
			document.getElementById('group').innerHTML='(*)';
			document.getElementById('path').innerHTML='';
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
				<img src="../grafica/flag_mini_it.jpg" alt="italiano" border="0">
				<input type="text" class="text" name="tft_sito_nome" value="<%= request("tft_sito_nome") %>" maxlength="250" size="75">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">Nome inglese:</td>
			<td class="content">
				<img src="../grafica/flag_mini_en.jpg" alt="inglese" border="0">
				<input type="text" class="text" name="tft_sito_nome_en" value="<%= request("tft_sito_nome_en") %>" maxlength="250" size="75">
			</td>
		</tr>

		<tr>
			<td class="label" rowspan="4" style="width:20%;">Tipo applicazione:</td>
			<td class="label" style="">
				<input type="radio" class="checkbox" value="1" name="sito_amministrazione" id="sito_amministrazione_true" value="1" <%= chk(request("sito_amministrazione")="1" OR request("sito_amministrazione")="") %> onClick="setState()">
				amministrativa
			</td>
		</tr>
		<tr>
			<td class="label"  style="width:80%;">
				percorso applicazione:
				<input type="text" class="text" name="tft_sito_dir" id="tft_sito_dir" value="<%= request("tft_sito_dir") %>" maxlength="150" size="75">
				<span id="path"></span>
			</td>
		</tr>
		<tr>
			<td class="label" style="width:15%;">
				<input type="radio" class="checkbox" value="" name="sito_amministrazione" value="" <%= chk(request("sito_amministrazione")="") %> onClick="setState()">
				pubblica
			</td>
		</tr>
		<tr>
			<td class="label">
				gruppo di lavoro:
				<% sql = "SELECT * FROM tb_gruppi"
				CALL dropDown(conn, sql, "id_Gruppo", "nome_Gruppo", "id_gruppo", request("id_gruppo"), true, " style=""width=auto""", LINGUA_ITALIANO) %>
				<span id="group"></span>
			</td>
		</tr>
		<tr>
			<td class="label" style="width:20%;">Protetta:</td>
			<td class="content">
				<input type="checkbox" class="noBorder" name="chk_sito_protetto" value="1" <%= chk(cBoolean(request("chk_sito_protetto")<>"",false))%>>
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
	setState()
</script>
</html>
