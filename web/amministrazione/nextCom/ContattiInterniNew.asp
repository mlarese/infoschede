<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ContattiSalva.asp")
end if

'--------------------------------------------------------
sezione_testata = "Inserimento nuovo contatto interno / sede alternativa" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim conn, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
%>

<script language="JavaScript" type="text/javascript">
	function set_modo_registra(){
		var isSocieta = document.getElementById('chk_issocieta_true');
		if (isSocieta.checked)
			form1.tft_modoregistra.value = form1.tft_nomeorganizzazioneelencoindirizzi.value;
		else
			form1.tft_modoregistra.value = form1.tft_cognomeelencoindirizzi.value;
		return true;
	}
	
	function show_mandatory(){
		var isSocieta = document.getElementById('chk_issocieta_true');
		var span_nome = document.getElementById('nome');
		var span_cognome = document.getElementById('cognome');
		var span_sede = document.getElementById('sede');
		var CntSede = document.getElementById('tfn_cntsede');
		
		DisableIfChecked(isSocieta, CntSede); 

		if (isSocieta.checked){
			span_sede.innerHTML='(*)';
			span_cognome.innerHTML='';
			span_nome.innerHTML='';
		}
		else{
			span_sede.innerHTML='';
			span_cognome.innerHTML='(*)';
			span_nome.innerHTML='(*)';
		}
		
	}
</script>

<div id="content_ridotto">
<form action="" method="post" id="form1" name="form1" onsubmit="set_modo_registra();">
	<input type="hidden" name="tfn_CntRel" value="<%= request("CNT") %>">
	<input type="hidden" name="tft_modoregistra" value="">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo contatto interno / sede alternativa</caption>
		<tr><th colspan="4">ANAGRAFICA</th></tr>
		<tr>
			<td class="label">salva come:</td>
			<td class="content">
				<table border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td><input class="noBorder" type="radio" name="chk_isSocieta" id="chk_issocieta_false" value="" <%= chk(request("chk_isSocieta")<>"1")%> onClick="show_mandatory()"></td>
						<td>contatto interno</td>
						<td style="padding-left:5px;"><input class="noBorder" type="radio" name="chk_isSocieta" id="chk_issocieta_true" value="1" <%= chk(request("chk_isSocieta")="1")%> onClick="show_mandatory()"></td>
						<td>sede alternativa</td>
					</tr>
				</table>
			</td>
			<td class="label" style="width:9%;">lingua:</td>
			<td class="content">
				<%CALL DropLingue(conn, NULL, "tft_lingua", request("tft_lingua"), true, false, "") %>
			</td>
		</tr>
		<tr>
			<td class="label">sede:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_nomeorganizzazioneelencoindirizzi" value="<%= request("tft_nomeorganizzazioneelencoindirizzi") %>" maxlength="100" size="65">
				<span id="sede">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label" nowrap>nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_nomeelencoindirizzi" value="<%= request("tft_NomeElencoIndirizzi") %>" maxlength="100" size="55">
				<span id="nome">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">cognome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_cognomeelencoindirizzi" value="<%= request("tft_CognomeElencoIndirizzi") %>" maxlength="100" size="55">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label" nowrap>ruolo / qualifica:</td>
			<td class="content" colspan="3"><input type="text" class="text" name="tft_qualificaelencoindirizzi" value="<%= request("tft_qualificaelencoindirizzi") %>" maxlength="250" size="45"></td>
		</tr>
		<tr><th colspan="4">INDIRIZZO</th></tr>
		<tr>
			<td class="label">sede:</td>
			<td class="content" colspan="3">
				<% 
				sql = " SELECT IDElencoIndirizzi, " + _
					  " (NomeOrganizzazioneElencoIndirizzi " + SQL_concat(conn) + "'  ('" + SQL_concat(conn) + "IndirizzoElencoindirizzi" + SQL_concat(conn) + " " + SQL_concat(conn) + "CapElencoIndirizzi" + SQL_concat(conn) + " " + SQL_concat(conn) + "CittaElencoIndirizzi" + SQL_concat(conn) + "')') AS NOME " + _
					  " FROM tb_indirizzario WHERE " & SQL_IsTrue(conn, "IsSocieta") & " AND CntRel=" & ParseSQL(request("CNT"), adChar)
				CALL dropDown(conn, sql, "IDElencoIndirizzi", "NOME", "tfn_cntsede", request("tfn_CntSede"), false, "style=""width:100%;""", LINGUA_ITALIANO)
				%>
			</td>
		</tr>
		<tr>
			<td class="label">indirizzo:</td>
			<td class="content" colspan="3"><input type="text" class="text" name="tft_IndirizzoElencoIndirizzi" value="<%= request("tft_IndirizzoElencoIndirizzi") %>" maxlength="250" size="55"></td>
		</tr>
		<tr>
			<td class="label">localit&agrave;:</td>
			<td class="content"><input type="text" class="text" name="tft_localitaElencoIndirizzi" value="<%= request("tft_localitaElencoIndirizzi") %>" maxlength="50" size="35"></td>
			<td class="label" style="width:9%;">cap:</td>
			<td class="content"><input type="text" class="text" name="tft_CAPElencoIndirizzi" value="<%= request("tft_CAPElencoIndirizzi") %>" maxlength="20" size="8"></td>
		</tr>
		<tr>
			<td class="label">citt&agrave;:</td>
			<td class="content"><input type="text" class="text" name="tft_cittaElencoIndirizzi" value="<%= request("tft_cittaElencoIndirizzi") %>" maxlength="50" size="35"></td>
			<td class="label" style="width:9%;">provincia:</td>
			<td class="content"><input type="text" class="text" name="tft_StatoProvElencoIndirizzi" value="<%= request("tft_StatoProvElencoIndirizzi") %>" maxlength="50" size="8"></td>
		</tr>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="3" name="tft_NoteElencoIndirizzi"><%=request("tft_NoteElencoIndirizzi")%></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
				<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
			</td>
		</tr>
	</table>
</form>
<script language="JavaScript" type="text/javascript">
	show_mandatory();
	FitWindowSize(this)
</script>
</div>
</body>
</html>
<% 
conn.close
set conn = nothing
%>
