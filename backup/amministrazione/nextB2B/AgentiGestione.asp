<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/classIndirizzarioLock.asp" -->
<!--#INCLUDE FILE="../library/classIndirizzarioSyncro.asp" -->
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../nextPassport/Tools_Passport.asp" -->
<!--#INCLUDE FILE="../library/ClassCryptography.asp"-->
<%
dim rs, rsr, sql, OBJ_Contatto, fields
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")

set OBJ_contatto = new IndirizzarioLock

if request("goto")<>"" then
	CALL GotoRecord(OBJ_contatto.conn, rs, session("B2B_AGENTI_SQL"), "IDElencoIndirizzi", "AgentiGestione.asp")
end if


if request.ServerVariables("REQUEST_METHOD") = "POST" then
	%>
	<!--#INCLUDE FILE="AgentiGestioneSalva_Tools.asp" -->
	<%
end if

if request("IDCNT") <> "" OR request("ID") <> "" then
	OBJ_contatto.LoadFromDB(cInteger(request("IDCNT") & request("ID")))
else
	OBJ_contatto.LoadFromForm("isSocieta")
end if


'gestione dei campi esterni a IndirizzarioLock come se fossero interni
sql = "SELECT * FROM gtb_agenti INNER JOIN tb_admin ON gtb_agenti.ag_admin_id = tb_admin.id_admin"
CaricaCampiEsterni OBJ_contatto.conn, rs, OBJ_contatto, sql, "ag_id", _
				   cInteger(GetValueList(OBJ_contatto.conn, rs, "SELECT ut_id FROM tb_utenti WHERE ut_NextCom_id=" & cInteger(request("ID"))))

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione agenti"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Agenti.asp"
dicitura.scrivi_con_sottosez()  
%>

<script language="JavaScript" type="text/javascript">

	function show_mandatory(){
		var isSocieta = document.getElementById('chk_issocieta_true');
		var span_ente = document.getElementById('ente')
		var span_cognome = document.getElementById('cognome')
		var span_nome = document.getElementById('nome')
		
		if (isSocieta.checked){
			span_ente.innerHTML='(*)'
			span_cognome.innerHTML=''
			span_nome.innerHTML=''
		}
		else{
			span_ente.innerHTML=''
			span_cognome.innerHTML='(*)'
			span_nome.innerHTML='(*)'
		}
		
	}
</script>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<!-- 
	Agente_id = <%= OBJ_contatto("ag_id") %>
	Utente_id = <%= OBJ_contatto("ut_id") %>
	Contatto_id = <%= OBJ_contatto("IDelencoindirizzi") %>
	Admin_id = <%= OBJ_contatto("ag_admin_id") %>
	 -->
	<input type="hidden" name="ag_admin_id" value="<%= OBJ_contatto("ag_admin_id") %>">
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<% 	if request("ID") = "" then %>
			<input type="hidden" name="tft_rubrica" value="<%= session("RUBRICA_AGENTI") %>">
			<caption>Inserimento nuovo agente</caption>
			<tr><th colspan="4">SELEZIONA L'AGENTE DAI CONTATTI</th></tr>
			<tr>
				<td class="label">contatto:</td>
				<td class="content" colspan="3">
                    <%
                    sql = " (( ut_id NOT IN (SELECT riv_id FROM gtb_rivenditori) AND ut_id NOT IN (SELECT ag_id FROM gtb_agenti)) OR " + SQL_IsNULL(OBJ_contatto.conn, "ut_id") + ") " + _
                          " AND (" + SQL_IsNULL(OBJ_contatto.conn, "cntRel") + " OR CntRel=0) "
                    CALL WriteContactPicker_Input(OBJ_contatto.conn, rs, sql, "", "form1", "tfn_IDElencoIndirizzi", request("IDCNT"), "LOGIN EMAIL", false, false, false, "REDIRECT")
                    %> 
				</td>
			</tr>
		<% 	else %>
			<caption>	
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td class="caption">Modifica dati dell'agente</td>
						<td align="right" style="font-size: 1px;">
							<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="agente precedente" <%= ACTIVE_STATUS %>>
								&lt;&lt; PRECEDENTE
							</a>
							&nbsp;
							<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="agente successivo" <%= ACTIVE_STATUS %>>
								SUCCESSIVO &gt;&gt;
							</a>
						</td>
					</tr>
				</table>
			</caption>
			<input type="hidden" name="tfn_IDElencoIndirizzi" value="<%= request("ID") %>">
		<% 	end if %>
		<tr><th colspan="4">DATI ANAGRAFICI</th></tr>
		<tr>
			<td class="label">salva come:</td>
			<td class="content">
				<table border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td><input class="noBorder" type="radio" name="isSocieta" id="chk_issocieta_false" value="" <%=chk(not OBJ_contatto("isSocieta"))%> onClick="show_mandatory()"></td>
						<td width="30%">persona fisica</td>
						<td><input class="noBorder" type="radio" name="isSocieta" id="chk_issocieta_true" value="1" <%=chk(OBJ_contatto("isSocieta"))%> onClick="show_mandatory()"></td>
						<td>ente / societ&agrave; / organizzazione</td>
					</tr>
				</table>
			</td>
			<td class="label">lingua:</td>
			<td class="content">
				<% CALL DropLingue(OBJ_contatto.conn, rs, "tft_lingua", OBJ_contatto("lingua"), true, false, "width:100px;") %>
			</td>
		</tr>
		<tr>
			<td class="label">ente / societ&agrave;:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_NomeOrganizzazioneElencoIndirizzi" value="<%= OBJ_contatto("NomeOrganizzazioneElencoIndirizzi") %>" maxlength="250" size="75">
				<span id="ente">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">cognome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_cognomeelencoindirizzi" value="<%= OBJ_contatto("CognomeElencoIndirizzi") %>" maxlength="100" size="75">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_nomeelencoindirizzi" value="<%= OBJ_contatto("NomeElencoIndirizzi") %>" maxlength="100" size="75">
				<span id="nome">(*)</span>
			</td>
		</tr>
		<script language="JavaScript" type="text/javascript">
			show_mandatory();
		</script>
		<tr>
			<td class="label">luogo di nascita:</td>
			<td class="content">
				<input type="text" class="text" name="tft_LuogoNascita" value="<%= OBJ_contatto("LuogoNascita") %>" maxlength="255" size="50">
			</td>
			<td class="label">data di nascita:</td>
			<td class="content">
				<input type="text" class="text" name="tfd_DTNASCElencoIndirizzi" value="<%= OBJ_contatto("DTNASCElencoIndirizzi") %>" maxlength="10" size="10">
			</td>
		</tr>
		<tr>
			<td class="label">p.iva / cod. fisc.:</td>
			<td class="content">
				<input type="text" class="text" name="tft_CF" value="<%= OBJ_contatto("CF") %>" maxlength="16" size="16">
			</td>
			<td class="label">codice aziendale:</td>
			<td class="content">
				<input type="text" class="text" name="extT_ag_codice" value="<%= OBJ_contatto("ag_codice") %>" maxlength="20" size="10">
			</td>
		</tr>
		<tr><th colspan="4">INDIRIZZO</th></tr>
		<tr>
			<td class="label">indirizzo:</td>
			<td class="content" colspan="3"><input type="text" class="text" name="tft_IndirizzoElencoIndirizzi" value="<%= OBJ_contatto("IndirizzoElencoIndirizzi") %>" maxlength="250" size="75"></td>
		</tr>
		<tr>
			<td class="label">localit&agrave;:</td>
			<td class="content"><input type="text" class="text" name="tft_LocalitaElencoIndirizzi" value="<%= OBJ_contatto("LocalitaElencoIndirizzi") %>" maxlength="50" size="50"></td>
			<td class="label">cap:</td>
			<td class="content" ><input type="text" class="text" name="tft_CAPElencoIndirizzi" value="<%= OBJ_contatto("CAPElencoIndirizzi") %>" maxlength="20" size="10"></td>
		</tr>
		<tr>
			<td class="label">citt&agrave;:</td>
			<td class="content"><input type="text" class="text" name="tft_cittaElencoIndirizzi" value="<%= OBJ_contatto("cittaElencoIndirizzi") %>" maxlength="50" size="50"></td>
			<td class="label" nowrap>provincia / stato:</td>
			<td class="content"><input type="text" class="text" name="tft_StatoProvElencoIndirizzi" value="<%= OBJ_contatto("StatoProvElencoIndirizzi") %>" maxlength="50" size="15"></td>
		</tr>
		<tr><th colspan="4">RECAPITI</th></tr>
		<tr>
			<td class="label">telefono:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_telefono" value="<%= OBJ_contatto("telefono") %>" maxlength="250" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">fax:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_fax" value="<%= OBJ_contatto("fax") %>" maxlength="20" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">cellulare:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_cellulare" value="<%= OBJ_contatto("cellulare") %>" maxlength="20" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">email:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_email" value="<%= OBJ_contatto("email") %>" maxlength="250" size="50">
				(*)
			</td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<tr><th colspan="4">PROFILO COMMERCIALE</th></tr>
		<tr>
			<td class="label">commissione:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="extN_ag_commissione" value="<%= FormatPrice(OBJ_contatto("ag_commissione"), 2, false) %>" maxlength="5" size="5">
				%
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="2">portafoglio clienti:</td>
			<td class="content" colspan="2">
				<input type="radio" class="checkbox" name="extN_ag_supervisore" value="0" <%= chk(not Obj_contatto("ag_supervisore")) %>>
				opera solo clienti associati
			</td>
		</tr>
		<tr>
			<td class="content supervisore" colspan="2">
				<input type="radio" class="checkbox" name="extN_ag_supervisore" value="1" <%= chk(Obj_contatto("ag_supervisore")) %>>
				opera su tutti i clienti
			</td>
		</tr>
		<tr>
			<td class="label">finestra di sconto:</td>
			<td class="content" style="width:30%;">
				<input type="text" class="text" name="extN_ag_range_sconto_massimo" value="<%= FormatPrice(OBJ_contatto("ag_range_sconto_massimo"), 2, false) %>" maxlength="5" size="5">
				%
			</td>
			<td class="note">
				La percentuale fissa il limite massimo di sconto applicabile dall'agente dal 
				prezzo applicato al cliente rispetto al prezzo minimo di vendita.
			</td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<tr><th colspan="4">PROFILO DI ACCESSO</th></tr>
		<input type="hidden" name="chk_abilitato" value="1">
		<tr>
			<td class="label">Login:</td>
			<td class="content">
                <input type="hidden" name="old_login" value="<%= OBJ_contatto("login") %>">
				<input type="text" class="text" name="tft_login" value="<%= OBJ_contatto("login") %>" maxlength="50" size="20">
				(*)
			</td>
			<td class="label">Scadenza:</td>
			<td class="content">
				<% CALL WriteDataPicker_Input("form1", "tfd_scadenza", OBJ_contatto("Scadenza"), "", "/", true, true, LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">Password:</td>
			<td class="content">
				<input type="password" class="text" name="tft_password" value="<%= OBJ_contatto("password") %>" maxlength="50" size="20">
				(*)
			</td>
			<td class="note" colspan="2" rowspan="2" style="width:55%;">
				Per i valori di login e password utilizzare solo caratteri alfanumerici o &quot;_&quot; 
				indifferentemente con lettere minuscole o maiuscole, ma senza spazi bianchi.
				<span style="letter-spacing:2px;">(<%= LOGIN_VALID_CHARSET %>)</span>
			</td>
		</tr>
		<tr>
			<td class="label">Conferma password:</td>
			<td class="content">
				<input type="password" class="text" name="conferma_password" value="<%= OBJ_contatto("password") %>" maxlength="50" size="20">
				(*)
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA <%= IIF(request("ID")="", "&gt;&gt;", "") %>">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% set rs = nothing %>