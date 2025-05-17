<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_Infoschede_Const.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede_Categorie.asp" -->
<!--#INCLUDE FILE="../library/classIndirizzarioLock.asp" -->
<!--#INCLUDE FILE="../library/ClassCryptography.asp"-->
<!--#INCLUDE FILE="../nextB2B/Tools_B2B.asp" -->

<%
dim OBJ_contatto, sql, rs, admin_id, applicativo_amm_agente, permesso_amm_agente, conn
set rs = Server.CreateObject("ADODB.RecordSet")
set OBJ_contatto = new IndirizzarioLock

if request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim cnt_id, ut_id, mandatory_flds, EsitoOk
	OBJ_contatto.LoadFromForm("chk_isSocieta;chk_abilitato;chk_TipoProfilo")
	
	mandatory_flds = "CognomeElencoIndirizzi;NomeElencoIndirizzi;"
	if OBJ_contatto("TipoProfilo") then
		mandatory_flds = mandatory_flds + "email;"
	end if
	if OBJ_contatto("SyncroKey")<>"" OR OBJ_contatto("SyncroTable")<>"" then
		mandatory_flds = mandatory_flds + "SyncroKey;SyncroTable;"
	end if
	
	
	'controlla altri campi obbligatori
	if OBJ_contatto.ValidateFields(mandatory_flds, OBJ_contatto("TipoProfilo")) then

		if OBJ_contatto("TipoProfilo") then
			EsitoOk = OBJ_contatto.ValidateLoginAndPassword(request("old_login"), request("conferma_password"))
		else
			EsitoOk = true
		end if
		
		if EsitoOk then
			'controllo esito positivo: dati validi
			OBJ_contatto.conn.beginTrans
			
			if request("ID") <> "" then			'sono in modifica o ho scelto dal menu
				OBJ_contatto.UpdateDB()
				cnt_id = OBJ_contatto("IDElencoIndirizzi")
			else
				cnt_id = OBJ_contatto.InsertIntoDB()
			end if
			
			if OBJ_contatto("TipoProfilo") then
				'registrazione utente
				'ut_id = OBJ_Contatto.UserFromContact(cnt_id, UTENTE_PERMESSO_SUBCLIENTE)
				ut_id = OBJ_Contatto.UserFromContact(cnt_id, PERMESSO_AR_CENTRO_ASSISTENZA &","& UTENTE_PERMESSO_SUBCLIENTE)
			else
				'CALL obj_contatto.RemoveUserFormContact(cnt_id, 0, UTENTE_PERMESSO_SUBCLIENTE)
				CALL obj_contatto.RemoveUserFormContact(cnt_id, 0, PERMESSO_AR_CENTRO_ASSISTENZA &","& UTENTE_PERMESSO_SUBCLIENTE)
			end if
			
			
			
			'gestione utente area amministrativa
			sql = "SELECT ut_admin_id FROM tb_utenti WHERE ut_id=" & ut_id
			sql = "SELECT * FROM tb_admin WHERE id_admin=" & cInteger(GetValueList(OBJ_contatto.conn, rs, sql))
			rs.open sql, OBJ_contatto.conn, adOpenKeySet, adLockOptimistic, adCmdText
			if rs.eof then
				rs.AddNew
			end if
			if OBJ_Contatto("IsSocieta") then
				rs("admin_cognome") =  OBJ_contatto("NomeOrganizzazioneElencoIndirizzi")
			else
				rs("admin_nome") = OBJ_contatto("NomeElencoIndirizzi")
				rs("admin_cognome") = OBJ_contatto("CognomeElencoIndirizzi")
			end if
			rs("admin_email") = OBJ_contatto("email")
			rs("admin_login") = OBJ_contatto("login")
			rs("admin_password") = EncryptPassword(OBJ_contatto("password"))
			if isDate(OBJ_contatto("Scandenza")) then
				rs("admin_scadenza") = NULL
			else
				rs("admin_scadenza") = OBJ_contatto("Scandenza")
			end if
			rs.update
			admin_id = rs("id_admin")
			rs.close


			'inserisce permesso per accesso area amministrativa
			if cIntero(request("chk_permesso")) = POS_PERMESSO_CENTRO_ASSISTENZA then
				permesso_amm_agente = POS_PERMESSO_CENTRO_ASSISTENZA
			else
				permesso_amm_agente = POS_PERMESSO_OFFICINA
			end if
			
			if Trim(cString(applicativo_amm_agente)) = "" then
				applicativo_amm_agente = INFOSCHEDE
			end if
			sql = " SELECT * FROM rel_admin_sito WHERE admin_id=" & admin_id & " AND sito_id=" & applicativo_amm_agente & _
				  " AND rel_as_permesso IN ("&POS_PERMESSO_CENTRO_ASSISTENZA&","&POS_PERMESSO_OFFICINA&")"
			rs.open sql, OBJ_contatto.conn, adOpenStatic, adLockOptimistic, adCmdText
			if rs.eof then
				rs.AddNew
				rs("admin_id") = admin_id
				rs("sito_id") = applicativo_amm_agente
				rs("rel_as_permesso") = permesso_amm_agente
				rs.update
			else
				rs("rel_as_permesso") = permesso_amm_agente
				rs.update
			end if  
			rs.close
			
			
			OBJ_contatto.conn.Execute("UPDATE tb_utenti SET ut_admin_id = "&admin_id&" WHERE ut_id = " & ut_id)
			
			
			'chiude transazione e conferma dati
			OBJ_contatto.conn.commitTrans %>
			<script language="JavaScript" type="text/javascript">
				opener.location.reload(true);
				window.close();
			</script>
		<%end if
	end if
end if

if request("ID") <> "" then
	'modifica del contatto
	OBJ_contatto.LoadFromDB(request("ID"))
	if OBJ_contatto.Fields.Exists("ut_id") then
		OBJ_contatto.Fields.Add "tipoprofilo", true
	end if
else
	OBJ_contatto.LoadFromForm("chk_isSocieta;chk_abilitato;chk_TipoProfilo")
end if



'--------------------------------------------------------
if request("ID")="" then
	sezione_testata = "Inserimento nuovo operatore" 
else
	sezione_testata = "Modifica dati operatore" 
end if%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'-----------------------------------------------------
%> 
<script language="JavaScript" type="text/javascript">
	function set_modo_registra(){
			form1.tft_modoregistra.value = form1.tft_cognomeelencoindirizzi.value;
	}
	
</script>


<div id="content_ridotto">
<form action="" method="post" id="form1" name="form1" onsubmit="set_modo_registra();">
	<input type="hidden" name="tft_modoregistra" value="">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<% 	if request("ID") = "" then 
			OBJ_contatto("CntRel") = cIntero(request("CNT")) %>
			<input type="hidden" name="tfn_CntRel" value="<%= OBJ_contatto("CntRel") %>">
			<caption>Inserimento operatore</caption>
		<% 	else %>
			<input type="hidden" name="tfn_IDElencoIndirizzi" value="<%= cIntero(request("ID")) %>">
			<caption>Modifica dati dell'operatore</caption>
		<% 	end if %>
		<tr><th colspan="4">ANAGRAFICA</th></tr>
		<tr>
			<input type="hidden" name="chk_isSocieta" id="chk_issocieta_false" value="" <%= chk(not OBJ_contatto("isSocieta"))%>>
			<td class="label" nowrap style="width:18%;">nome:</td>
			<td class="content">
				<input type="text" class="text" name="tft_nomeelencoindirizzi" value="<%= OBJ_contatto("NomeElencoIndirizzi") %>" maxlength="100" size="40">
				<span id="nome">(*)</span>
			</td>
			<td class="label_right">lingua:</td>
			<td class="content">
				<%CALL DropLingue(OBJ_contatto.conn, NULL, "tft_lingua", OBJ_contatto("lingua"), true, false, "") %>
			</td>
		</tr>
		<tr>
			<td class="label">cognome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_cognomeelencoindirizzi" value="<%= OBJ_contatto("CognomeElencoIndirizzi") %>" maxlength="100" size="40" onChange="set_modo_registra();">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">ruolo / qualifica:</td>
			<td class="content" colspan="3"><input type="text" class="text" name="tft_qualificaelencoindirizzi" value="<%= OBJ_contatto("qualificaelencoindirizzi") %>" maxlength="250" size="40"></td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<tr><th colspan="4">INDIRIZZO</th></tr>
		<tr>
			<td class="label" style="width:18%;">indirizzo:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_IndirizzoElencoIndirizzi" value="<%= OBJ_contatto("IndirizzoElencoIndirizzi") %>" maxlength="250" size="55">
			</td>
		</tr>
		<tr>
			<td class="label">localit&agrave;:</td>
			<td class="content"><input type="text" class="text" name="tft_localitaElencoIndirizzi" value="<%= OBJ_contatto("localitaElencoIndirizzi") %>" maxlength="50" size="35"></td>
			<td class="label">cap:</td>
			<td class="content"><input type="text" class="text" name="tft_CAPElencoIndirizzi" value="<%= OBJ_contatto("CAPElencoIndirizzi") %>" maxlength="20" size="8"></td>
		</tr>
		<tr>
			<td class="label">citt&agrave;:</td>
			<td class="content">
				<input type="text" class="text" name="tft_cittaElencoIndirizzi" value="<%= OBJ_contatto("cittaElencoIndirizzi") %>" maxlength="50" size="35">
			</td>
			<td class="label">provincia:</td>
			<td class="content"><input type="text" class="text" name="tft_StatoProvElencoIndirizzi" value="<%= OBJ_contatto("StatoProvElencoIndirizzi") %>" maxlength="50" size="8"></td>
		</tr>
		<tr><th colspan="4">RECAPITI</th></tr>
		<tr>
			<td class="label">telefono:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_telefono" value="<%= OBJ_contatto("telefono") %>" maxlength="250" size="45">
			</td>
		</tr>
		<tr>
			<td class="label">fax:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_fax" value="<%= OBJ_contatto("fax") %>" maxlength="20" size="45">
			</td>
		</tr>
		<tr>
			<td class="label">cellulare:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_cellulare" value="<%= OBJ_contatto("cellulare") %>" maxlength="20" size="45">
			</td>
		</tr>
		<tr>
			<td class="label">email:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_email" value="<%= OBJ_contatto("email") %>" maxlength="250" size="55">
				<span id="email">(*)</span>
			</td>
		</tr>
	</table>
	<input type="Hidden" name="old_login" value="<%= OBJ_contatto("login") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<tr><th colspan="4">PROFILO DI ACCESSO SUBORDINATO</th></tr>
		<%  dim permesso
			sql = " SELECT ut_admin_id FROM tb_utenti WHERE ut_id = " & cIntero(OBJ_contatto("ut_id"))
			sql = " SELECT rel_as_permesso FROM rel_admin_sito WHERE sito_id = " & INFOSCHEDE & _
			      " AND admin_id = " & cIntero(GetValueList(OBJ_contatto.conn, NULL, sql))   
			permesso = cIntero(GetValueList(OBJ_contatto.conn, NULL, sql))
			if permesso = 0 then
				permesso = POS_PERMESSO_CENTRO_ASSISTENZA
			end if
		%>
		<tr>
			<td class="label">permesso:</td>
			<td class="content">
				<input class="checkbox" type="radio" name="chk_permesso" id="chk_permesso_centro_assistenza" value="<%=POS_PERMESSO_CENTRO_ASSISTENZA%>" <%= chk(permesso = POS_PERMESSO_CENTRO_ASSISTENZA)%>>
				Responsabile centro assistenza
			</td>
			<td class="content">
				<input class="checkbox" type="radio" name="chk_permesso" id="chk_permesso_officina" value="<%=POS_PERMESSO_OFFICINA%>" <%= chk(permesso = POS_PERMESSO_OFFICINA)%>>
				Operatore officina
			</td>
		</tr>
		
		<input type="Hidden" name="chk_tipoprofilo" id="chk_tipoprofilo_true" value="1">
		<tr>
			<td class="label" style="width:22%;">stato:</td>
			<td class="content" style="width:43%;">
				<input type="checkbox" class="checkbox" name="chk_abilitato" <%= Chk(IIF(request("ID")<>"" OR request.serverVariables("REQUEST_METHOD")="POST", OBJ_contatto("abilitato"), true)) %>>
				abilitato all'accesso
			</td>
			<td class="note" rowspan="5">
				Per i valori di login e password utilizzare solo caratteri alfanumerici o &quot;_&quot; 
				indifferentemente con lettere minuscole o maiuscole, ma senza spazi bianchi.
				<span style="letter-spacing:2px;">(<%= replace(LOGIN_VALID_CHARSET, "s", "s ") %>)</span>
			</td>
		</tr>
		<tr>
			<td class="label">Login:</td>
			<td class="content">
				<input type="text" class="text" name="tft_login" value="<%= OBJ_contatto("login") %>" maxlength="50" size="20">
				<span id="login">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">Password:</td>
			<td class="content">
				<input type="password" class="text" name="tft_password" value="<%= OBJ_contatto("password") %>" maxlength="50" size="20">
				<span id="password">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">Conferma password</td>
			<td class="content">
				<input type="password" class="text" name="conferma_password" value="<%= OBJ_contatto("password") %>" maxlength="50" size="20">
				<span id="c_password">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">Scadenza accesso:</td>
			<td class="content">
				<% CALL WriteDataPicker_Input("form1", "tfd_scadenza", OBJ_contatto("Scadenza"), "", "/", true, true, LINGUA_ITALIANO) %>
			</td>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="3" name="tft_NoteElencoIndirizzi"><%=OBJ_contatto("NoteElencoIndirizzi")%></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
				<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
			</td>
		</tr>
</form>
	</table>
</div>
</body>
</html>
<% 
set OBJ_contatto = nothing
%>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>