<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->
<% '*******************************************************************************************************************************
ParentFrameName = "IFrameMacchine"
bodyClass ="contattimacchine" %>
<!--#INCLUDE FILE="../library/Intestazione_iframe.asp" -->
<% '*******************************************************************************************************************************

dim conn, rs, rsc, sql, BlankRowsCount, i

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

if Request.ServerVariables("REQUEST_METHOD")="POST" then
'call listRequest()
'response.end
	'modifica esistenti
	dim imaIds
	imaIds = split(replace(request("ima_id"), " ", ""), ",")
	for each i in imaIds
		sql = "UPDATE tb_indirizzario_macchine SET ima_esito_trattativa = NULL WHERE ima_id = " & i
		conn.execute(sql)
		'modifica la riga
		CALL SalvaCampiEsterniUltra(conn, rs, _
									"SELECT * FROM tb_indirizzario_macchine", _
									"ima_id", i, "ima_contatto_id", request("id"), "mod_"&i&"_C_ima_stato_trattativa", _
									"ima_", request.Form, "mod_" & i & "_")
	next
	
	'inserimento nuovi
	BlankRowsCount = cIntero(request("BlankRowsCount"))
	for i = 1 to BlankRowsCount
		if Trim(cstring(request("new_" & i & "_t_ima_marchio") & request("new_" & i & "_t_ima_modello") & request("new_" & i & "_t_ima_numero")))<>"" then
			CALL SalvaCampiEsterniUltra(conn, rs, _
										"SELECT * FROM tb_indirizzario_macchine", _
										"ima_id", 0, "ima_contatto_id", request("id"), "new_"&i&"_C_ima_stato_trattativa", _
										"ima_", request.Form, "new_" & i & "_")
		end if
	next
elseif cIntero(request("cancella"))>0 then
	
	'cancella la macchina indicata
	sql = "DELETE FROM tb_indirizzario_macchine WHERE ima_id=" & request("cancella")
	CALL Conn.execute(sql)
	
	response.redirect "ContattiMacchine.asp?ID=" & request("ID")
	
end if

%>
<script language="JavaScript" type="text/javascript">
	function CheckCancella(ima_id){
		if (window.confirm('Cancellare la registrazione?')){
			document.location = "ContattiMacchine.asp?ID=<%=request("ID")%>&cancella=" + ima_id;
		}
	}
</script>

<form action="" method="post" id="form1" name="form1">
	<script language="JavaScript" type="text/javascript">
		function SetRadioButton(button, prefix){
			var esito1 = document.getElementById(prefix + 'esito_in_corso');
			EnableIfChecked(button, esito1);
			var esito2 = document.getElementById(prefix + 'esito_vinta');
			EnableIfChecked(button, esito2);
			var esito3 = document.getElementById(prefix + 'esito_persa');
			EnableIfChecked(button, esito3);
			
			//var data_scadenza = document.getElementById(prefix + 'd_ima_scadenza_data');
			//DisableIfChecked(button, data_scadenza);
			
			var esitoHidden = document.getElementById(prefix + '_esito_hidden');
			if (!button.checked){
				esito1.checked = false;
				esito2.checked = false;
				esito3.checked = false;
			}
		}
		
		function SetDate(clicked, prefix){
			var dateField = document.getElementById(prefix + 'd_ima_chiusura_trattativa_data');
			if (clicked.value > 0){
				if (dateField.value == ''){
					dateField.value = '<%=DateIta(Now())%>';
				}
			}
			else {
				dateField.value = '';
			}
		}
	</script>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-right:0px; border-left:0px;" >
		<% if session("ERRORE")<> "" then %>
			<tr> <td class="errore" colspan="6"><%= Session("ERRORE")%></th> </tr>
			<% Session("ERRORE")=""
		end if
		%>
		<tr>
			<th colspan="14">PARCO MACCHINE</th>
		</tr>
		<tr>
			<td colspan="14" class="content note">Registrazione del parco macchine multifunzioni/stampanti del contatto</td>
		</tr>
		<tr>
			<th class="L2">Marca</th>
			<th class="L2">Modello</th>
			<th class="L2" style="width:4%;">N&ordm;</th>
			<th class="L2" style="width:9%;">B/N - Colore</th>
			<th class="L2" style="width:7%;">Contratto</th>
			<th class="L2" style="width:7%;">installazione</th>
			<th class="L2" style="width:7%;">note scadenza</th>
			<th class="L2" style="width:7%;">data scadenza</th>
			<th class="L2">Fornitore</th>
			<th class="L2" style="width:7%;">Matricola</th>
			<th class="L2" style="width:5%;">Trattativa</th>
			<th class="L2" style="width:6%;">Stato trattativa</th>
			<th class="L2">Chiusura trattativa il</th>
			<th class="L2" style="width:6%;">cancella</th>
		</tr>
		<%
		sql = "SELECT * FROM tb_indirizzario_macchine where ima_contatto_id=" & cIntero(request("ID"))
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		BlankRowsCount = IIF(rs.recordcount<2, 4, 2)
		if request("salva_aggiungi_righe")<>"" then
			BlankRowsCount = BlankRowsCount + 2
		end if
		%>
		<input type="hidden" name="BlankRowsCount" value="<%=BlankRowsCount%>">
		
		<% while not rs.eof %>
			<input type="hidden" name="ima_id" value="<%=rs("ima_id")%>">
			<tr <%=IIF(rs("ima_stato_trattativa"), "class=""trattative""", "")%>>
				<td class="content">
					<input id="mod_<%=rs("ima_id")%>_t_ima_marchio" name="mod_<%=rs("ima_id")%>_t_ima_marchio" value="<%=rs("ima_marchio")%>" type="text" class="text" style="width:100%;" maxlength="250">
				</td>
				<td class="content">
					<input id="mod_<%=rs("ima_id")%>_t_ima_modello" name="mod_<%=rs("ima_id")%>_t_ima_modello" value="<%=rs("ima_modello")%>" type="text" class="text" style="width:100%;" maxlength="250">
				</td>
				<td class="content">
					<input id="mod_<%=rs("ima_id")%>_t_ima_numero" name="mod_<%=rs("ima_id")%>_t_ima_numero" value="<%=rs("ima_numero")%>" type="text" class="number" style="width:100%;" maxlength="50">
				</td>
				<td class="content">
					<table cellspacing="0" cellpadding="0" width="100%">
						<tr>
							<td class="content">
								<input type="radio" name="mod_<%=rs("ima_id")%>_t_ima_tipocolore" id="mod_<%=rs("ima_id")%>_t_ima_tipocolore_bn" value="Bianco/Nero" class="noborder" <%=chk(rs("ima_tipocolore") = "Bianco/Nero")%>>
								Bianco e nero
							</td>
						</tr>
						<tr>
							<td class="content">
								<input type="radio" name="mod_<%=rs("ima_id")%>_t_ima_tipocolore" id="mod_<%=rs("ima_id")%>_t_ima_tipocolore_colore" value="Colore" class="noborder" <%=chk(rs("ima_tipocolore") = "Colore")%>>
								Colore
							</td>
						</tr>
					</table>
				</td>
				<td class="content">
					<table cellspacing="0" cellpadding="0" width="100%">
						<tr>
							<td class="content">
								<input type="radio" name="mod_<%=rs("ima_id")%>_t_ima_contratto" id="mod_<%=rs("ima_id")%>_t_ima_contratto_acq" value="Acquisto" class="noborder" <%=chk(rs("ima_contratto") = "Acquisto")%>>
								Acquisto
							</td>
						</tr>
						<tr>
							<td class="content">
								<input type="radio" name="mod_<%=rs("ima_id")%>_t_ima_contratto" id="mod_<%=rs("ima_id")%>_t_ima_contratto_noleggio" value="Noleggio" class="noborder" <%=chk(rs("ima_contratto") = "Noleggio")%>>
								Noleggio
							</td>
						</tr>
					</table>
				</td>
				<td class="content">
					<input id="mod_<%=rs("ima_id")%>_t_ima_installazione" name="mod_<%=rs("ima_id")%>_t_ima_installazione" value="<%=rs("ima_installazione")%>" type="text" class="text" style="width:100%;" maxlength="250">
				</td>
				<td class="content">
					<input id="mod_<%=rs("ima_id")%>_t_ima_scadenza" name="mod_<%=rs("ima_id")%>_t_ima_scadenza" value="<%=rs("ima_scadenza")%>" type="text" class="text" style="width:100%;" maxlength="250">
				</td>
				<td class="content">
					<% CALL WriteDataPicker_Input("form1", "mod_"&rs("ima_id")&"_d_ima_scadenza_data", rs("ima_scadenza_data"), "", "/", true, true, LINGUA_ITALIANO) %>
				</td>
				<td class="content">
					<input id="mod_<%=rs("ima_id")%>_t_ima_fornitore" name="mod_<%=rs("ima_id")%>_t_ima_fornitore" value="<%=rs("ima_fornitore")%>" type="text" class="text" style="width:100%;" maxlength="250">
				</td>
				<td class="content">
					<input id="mod_<%=rs("ima_id")%>_t_ima_matricola" name="mod_<%=rs("ima_id")%>_t_ima_matricola" value="<%=rs("ima_matricola")%>" type="text" class="text" style="width:100%;" maxlength="250">
				</td>
				<td class="content">
					<input type="checkbox" class="noborder" name="mod_<%=rs("ima_id")%>_C_ima_stato_trattativa" onclick="SetRadioButton(this, 'mod_<%=rs("ima_id")%>_')" <%=chk(rs("ima_stato_trattativa"))%>>
					trattativa
				</td>	
				<td class="content">
					<table cellspacing="0" cellpadding="0" width="100%">
						<tr>
							<td class="content">
								<input type="radio" name="mod_<%=rs("ima_id")%>_n_ima_esito_trattativa" value="0" id="mod_<%=rs("ima_id")%>_esito_in_corso" onclick="SetDate(this, 'mod_<%=rs("ima_id")%>_')" class="noborder" <%=chk(cIntero(rs("ima_esito_trattativa")) = 0 AND cString(rs("ima_esito_trattativa"))<>"")%>>
								In corso
							</td>
						</tr>
						<tr>
							<td class="content">
								<input type="radio" name="mod_<%=rs("ima_id")%>_n_ima_esito_trattativa" value="1" id="mod_<%=rs("ima_id")%>_esito_vinta" onclick="SetDate(this, 'mod_<%=rs("ima_id")%>_')" class="noborder" <%=chk(cIntero(rs("ima_esito_trattativa")) = 1)%>>
								Vinta
							</td>
						</tr>
						<tr>
							<td class="content">
								<input type="radio" name="mod_<%=rs("ima_id")%>_n_ima_esito_trattativa" value="2" id="mod_<%=rs("ima_id")%>_esito_persa" onclick="SetDate(this, 'mod_<%=rs("ima_id")%>_')" class="noborder" <%=chk(cIntero(rs("ima_esito_trattativa")) = 2)%>>
								Persa 
								<%
								if cIntero(rs("ima_esito_trattativa")) = 2 then
									%><span style="float:right"><%
									CALL WriteColoreTipo("#FF0606", "Trattativa persa")
									%></span><%
								end if
								%>
							</td>
						</tr>
					</table>
				</td>
				<td class="content">
					<% CALL WriteDataPicker_Input_Manuale("form1", "mod_"&rs("ima_id")&"_d_ima_chiusura_trattativa_data", rs("ima_chiusura_trattativa_data"), "", "/", true, true, LINGUA_ITALIANO, "", false, "width:70px;") %>
				</td>
				<td class="content_center">
					<a class="button_L2" href="javascript:void(0);" onclick="CheckCancella('<%=rs("ima_id")%>')">
						CANCELLA
					</a>
				</td>
			</tr>
			<script language="JavaScript" type="text/javascript">
				SetRadioButton(form1.mod_<%=rs("ima_id")%>_C_ima_stato_trattativa, 'mod_<%=rs("ima_id")%>_');
			</script>
			<% rs.movenext
		wend
		rs.close
		
		for i = 1 to BlankRowsCount%>
			<tr class="vuote">
				<td class="content">
					<input id="new_<%=i%>_t_ima_marchio" name="new_<%=i%>_t_ima_marchio" type="text" class="text" style="width:100%;" maxlength="250">
				</td>
				<td class="content">
					<input id="new_<%=i%>_t_ima_modello" name="new_<%=i%>_t_ima_modello" type="text" class="text" style="width:100%;" maxlength="250">
				</td>
				<td class="content">
					<input id="new_<%=i%>_t_ima_numero" name="new_<%=i%>_t_ima_numero" type="text" class="number" style="width:100%;" maxlength="50">
				</td>
				<td class="content">
					<table cellspacing="0" cellpadding="0" width="100%">
						<tr>
							<td class="content">
								<input type="radio" name="new_<%=i%>_t_ima_tipocolore" id="new_<%=i%>_t_ima_tipocolore_bn" value="Bianco/Nero" class="noborder">
								Bianco e nero
							</td>
						</tr>
						<tr>
							<td class="content">
								<input type="radio" name="new_<%=i%>_t_ima_tipocolore" id="new_<%=i%>_t_ima_tipocolore_colore" value="Colore" class="noborder">
								Colore
							</td>
						</tr>
					</table>
				</td>
				<td class="content">
					<table cellspacing="0" cellpadding="0" width="100%">
						<tr>
							<td class="content">
								<input type="radio" name="new_<%=i%>_t_ima_contratto" id="new_<%=i%>_t_ima_contratto_acq" value="Acquisto" class="noborder">
								Acquisto
							</td>
						</tr>
						<tr>
							<td class="content">
								<input type="radio" name="new_<%=i%>_t_ima_contratto" id="new_<%=i%>_t_ima_contratto_noleggio" value="Noleggio" class="noborder">
								Noleggio
							</td>
						</tr>
					</table>
				</td>
				<td class="content">
					<input id="new_<%=i%>_t_ima_installazione" name="new_<%=i%>_t_ima_installazione" type="text" class="text" style="width:100%;" maxlength="250">
				</td>
				<td class="content">
					<input id="new_<%=i%>_t_ima_scadenza" name="new_<%=i%>_t_ima_scadenza" type="text" class="text" style="width:100%;" maxlength="250">
				</td>
				<td class="content">
					<% CALL WriteDataPicker_Input("form1", "new_"&i&"_d_ima_scadenza_data", "", "", "/", true, true, LINGUA_ITALIANO) %>
				</td>
				<td class="content">
					<input id="new_<%=i%>_t_ima_fornitore" name="new_<%=i%>_t_ima_fornitore" type="text" class="text" style="width:100%;" maxlength="250">
				</td>
				<td class="content">
					<input id="new_<%=i%>_t_ima_matricola" name="new_<%=i%>_t_ima_matricola" type="text" class="text" style="width:100%;" maxlength="250">
				</td>
				<td class="content">
					<input type="checkbox" class="noborder" name="new_<%=i%>_C_ima_stato_trattativa" onclick="SetRadioButton(this, 'new_<%=i%>_')">
					trattativa
				</td>	
				<td class="content">
					<table cellspacing="0" cellpadding="0" width="100%">
						<tr>
							<td class="content">
								<input type="radio" name="new_<%=i%>_n_ima_esito_trattativa" id="new_<%=i%>_esito_in_corso" onclick="SetDate(this, 'new_<%=i%>_')" value="0" class="noborder">
								In corso
							</td>
						</tr>
						<tr>
							<td class="content">
								<input type="radio" name="new_<%=i%>_n_ima_esito_trattativa" id="new_<%=i%>_esito_vinta" onclick="SetDate(this, 'new_<%=i%>_')" value="1" class="noborder">
								Vinta
							</td>
						</tr>
						<tr>
							<td class="content">
								<input type="radio" name="new_<%=i%>_n_ima_esito_trattativa" id="new_<%=i%>_esito_persa" onclick="SetDate(this, 'new_<%=i%>_')" value="2" class="noborder">
								Persa
							</td>
						</tr>
					</table>
				</td>
				<td class="content">
					<% CALL WriteDataPicker_Input_Manuale("form1", "new_"&i&"_d_ima_chiusura_trattativa_data", request("new_"&i&"_d_ima_chiusura_trattativa_data"), "", "/", true, true, LINGUA_ITALIANO, "", false, "width:70px;") %>
				</td>
				<td class="content">&nbsp;</td>
			</tr>
			<script language="JavaScript" type="text/javascript">
				SetRadioButton(form1.new_<%=i%>_C_ima_stato_trattativa, 'new_<%=i%>_');
			</script>
		<% next %>
		<tr class="vuote">
			<td class="content_right" colspan="14">
				<input type="submit" name="salva_aggiungi_righe" id="salva_aggiungi_righe" class="button_L2" value="Aggiungi altre righe">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% 

conn.close 
set rs = nothing
set rsc = nothing
set conn = nothing

%>