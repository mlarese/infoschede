<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ContattiSalva.asp")
end if


'--------------------------------------------------------
sezione_testata = "Modifica dati del contatto interno" 
'testata_elenco_pulsanti = "RECAPITI"
'testata_elenco_href = "ContattiInterniRecapiti.asp?ID=" & request("ID")

%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT * FROM tb_indirizzario WHERE IDElencoIndirizzi=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdtext
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

<div id="content_ridotto"> <!-- ID usato anche su javascript alla fine della pagina -->
<form action="" method="post" id="form1" name="form1" onsubmit="set_modo_registra();">
	<input type="hidden" name="tfn_CntRel" value="<%= rs("CntRel") %>">
	<input type="hidden" name="tft_modoregistra" value="">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Modifica dati del contatto interno / sede alternativa</caption>
		<tr>
			<td style="width:50%; border-right:1px solid #999999;">
				<table cellspacing="1" cellpadding="0" style="" style="width:100%;">
					<tr><th colspan="4">ANAGRAFICA</th></tr>
					<tr>
						<td class="label">salva come:</td>
						<td class="content">
							<table border="0" cellspacing="0" cellpadding="0" align="left">
								<tr>
									<td><input class="noBorder" type="radio" name="chk_isSocieta" id="chk_issocieta_false" value="" <%= chk(not rs("isSocieta"))%> onClick="show_mandatory()"></td>
									<td>contatto interno</td>
									<td style="padding-left:5px;"><input class="noBorder" type="radio" name="chk_isSocieta" id="chk_issocieta_true" value="1" <%= chk(rs("isSocieta"))%> onClick="show_mandatory()"></td>
									<td>sede alternativa</td>
								</tr>
							</table>
						</td>
						<td class="label" style="width:9%;">lingua:</td>
						<td class="content">
							<%CALL DropLingue(conn, NULL, "tft_lingua", rs("lingua"), true, false, "") %>
						</td>
					</tr>
					<tr>
						<td class="label">sede:</td>
						<td class="content" colspan="3">
							<input type="text" class="text" name="tft_nomeorganizzazioneelencoindirizzi" value="<%= rs("nomeorganizzazioneelencoindirizzi") %>" maxlength="100" style="width:95%;">
							<span id="sede">(*)</span>
						</td>
					</tr>
					<tr>
						<td class="label" nowrap>nome:</td>
						<td class="content" colspan="3">
							<input type="text" class="text" name="tft_nomeelencoindirizzi" value="<%= rs("NomeElencoIndirizzi") %>" maxlength="100" style="width:95%;">
							<span id="nome">(*)</span>
						</td>
					</tr>
					<tr>
						<td class="label">cognome:</td>
						<td class="content" colspan="3">
							<input type="text" class="text" name="tft_cognomeelencoindirizzi" value="<%= rs("CognomeElencoIndirizzi") %>" maxlength="100" style="width:95%;">
							<span id="cognome">(*)</span>
						</td>
					</tr>
					<tr>
						<td class="label" nowrap>ruolo / qualifica:</td>
						<td class="content" colspan="3"><input type="text" class="text" name="tft_qualificaelencoindirizzi" value="<%= rs("qualificaelencoindirizzi") %>" maxlength="250" style="width:50%;"></td>
					</tr>
					<tr><th colspan="4">INDIRIZZO</th></tr>
					<tr>
						<td class="label">sede:</td>
						<td class="content" colspan="3">
							<% 
							sql = " SELECT IDElencoIndirizzi, " + _
								  " (NomeOrganizzazioneElencoIndirizzi " + SQL_concat(conn) + "'  ('" + SQL_concat(conn) + "IndirizzoElencoindirizzi" + SQL_concat(conn) + " " + SQL_concat(conn) + "CapElencoIndirizzi" + SQL_concat(conn) + " " + SQL_concat(conn) + "CittaElencoIndirizzi" + SQL_concat(conn) + "')') AS NOME " + _
								  " FROM tb_indirizzario WHERE " & SQL_IsTrue(conn, "IsSocieta") & " AND CntRel=" & rs("CntRel") & " AND IdElencoIndirizzi<>" & cIntero(request("ID"))
							CALL dropDown(conn, sql, "IDElencoIndirizzi", "NOME", "tfn_cntsede", rs("CntSede"), false, "style=""width:100%;""", LINGUA_ITALIANO)
							%>
						</td>
					</tr>
					<tr>
						<td class="label">indirizzo:</td>
						<td class="content" colspan="3"><input type="text" class="text" name="tft_IndirizzoElencoIndirizzi" value="<%= rs("IndirizzoElencoIndirizzi") %>" maxlength="250" style="width:100%;"></td>
					</tr>
					<tr>
						<td class="label">localit&agrave;:</td>
						<td class="content"><input type="text" class="text" name="tft_localitaElencoIndirizzi" value="<%= rs("localitaElencoIndirizzi") %>" maxlength="50" style="width:100%;"></td>
						<td class="label" style="width:9%;">cap:</td>
						<td class="content"><input type="text" class="text" name="tft_CAPElencoIndirizzi" value="<%= rs("CAPElencoIndirizzi") %>" maxlength="20" style="width:100%;"></td>
					</tr>
					<tr>
						<td class="label">citt&agrave;:</td>
						<td class="content"><input type="text" class="text" name="tft_cittaElencoIndirizzi" value="<%= rs("cittaElencoIndirizzi") %>" maxlength="50" style="width:100%;"></td>
						<td class="label" style="width:9%;">provincia:</td>
						<td class="content"><input type="text" class="text" name="tft_StatoProvElencoIndirizzi" value="<%= rs("StatoProvElencoIndirizzi") %>" maxlength="50" style="width:100%;"></td>
					</tr>
					<tr><th colspan="4">NOTE</th></tr>
					<tr>
						<td class="content" colspan="4">
							<textarea style="width:100%;" rows="3" name="tft_NoteElencoIndirizzi"><%=rs("NoteElencoIndirizzi")%></textarea>
						</td>
					</tr>
				</table>
			</td>
			<td style="vertical-align:top;">
				<%
				dim iframe_url
				iframe_url = GetSiteUrl(conn, 0, 0) & "/amministrazione/nextCom/ContattiRecapiti_iFrame.asp?MODE=iframe&ID=" & request("ID")
				%>
				<iframe id="IFrameRecapiti" name="" style="border:0px; width:100%;" src="<%= iframe_url %>"  frameborder="0" scrolling="no">
				</iframe>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input style="width:14%;" type="button" class="button" name="salva" value="SALVA" onclick="window.frames[0].document.forms[0].submit();">
				<!--<input style="width:20%;" type="button" class="button" name="salva_elenco" value="SALVA & TORNA AL CONTATTO" onclick="window.frames[0].document.forms[0].submit();">-->
				<input style="width:14%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
			</td>
		</tr>
	</form>
	</table>
</div>
</body>
</html>
<% 
conn.close
set conn = nothing
%>
<script language="JavaScript" type="text/javascript">
	show_mandatory();
	//alert(document);
	FitWindowSize(this)
</script>