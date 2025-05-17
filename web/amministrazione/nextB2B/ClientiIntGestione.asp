<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/classIndirizzarioLock.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<%
dim OBJ_contatto, sql
set OBJ_contatto = new IndirizzarioLock

if request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim cnt_id, ut_id, mandatory_flds, EsitoOk
	OBJ_contatto.LoadFromForm("chk_isSocieta;chk_abilitato;chk_TipoProfilo")
	
	if OBJ_contatto("isSocieta") then
		mandatory_flds = "NomeOrganizzazioneElencoIndirizzi;IndirizzoElencoIndirizzi;CittaElencoIndirizzi;"
		if OBJ_contatto("TipoProfilo") then
			mandatory_flds = mandatory_flds + "CognomeElencoIndirizzi;NomeElencoIndirizzi;"
		end if
	else
		mandatory_flds = "CognomeElencoIndirizzi;NomeElencoIndirizzi;"
	end if
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
			
			OBJ_Contatto("PraticaPrefisso") = OBJ_Contatto("SyncroKey")
			
			if request("ID") <> "" then			'sono in modifica o ho scelto dal menu
				OBJ_contatto.UpdateDB()
				cnt_id = OBJ_contatto("IDElencoIndirizzi")
			else
				cnt_id = OBJ_contatto.InsertIntoDB()
			end if
			
			if OBJ_contatto("TipoProfilo") then
				'registrazione utente
				ut_id = OBJ_Contatto.UserFromContact(cnt_id, UTENTE_PERMESSO_SUBCLIENTE)
			else
				CALL obj_contatto.RemoveUserFormContact(cnt_id, 0, UTENTE_PERMESSO_SUBCLIENTE)
			end if
			
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
		if cIntero(OBJ_contatto("ut_id"))>0 then
			OBJ_contatto.Fields.Add "tipoprofilo", true
		end if
	end if
else
	OBJ_contatto.LoadFromForm("chk_isSocieta;chk_abilitato;chk_TipoProfilo")
end if



'--------------------------------------------------------
if request("ID")="" then
	sezione_testata = "Inserimento nuovo contatto interno / sede alternativa" 
else
	sezione_testata = "Modifica dati contatto interno / sede alternativa" 
end if%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'-----------------------------------------------------
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
		var span_nome = document.getElementById('nome')
		var span_cognome = document.getElementById('cognome')
		var span_sede = document.getElementById('sede')
		var span_indirizzo = document.getElementById('indirizzo')
		var span_citta = document.getElementById('citta')
		
		var CntSede = document.getElementById('tfn_cntsede');
		
		DisableIfChecked(isSocieta, CntSede); 
		
		var profilo = document.getElementById('chk_tipoprofilo_true')
		var span_login = document.getElementById('login')
		var span_password = document.getElementById('password')
		var span_c_password = document.getElementById('c_password')
		var span_email = document.getElementById('email')
		
		var input_tipo_sede = document.getElementById('tft_syncrotable')
		var input_codice_sede = document.getElementById('tft_syncrokey')
		var span_tipo_sede = document.getElementById('tipo_sede')
		var span_codice_sede = document.getElementById('codice_sede')
		
		if (isSocieta.checked){
			span_sede.innerHTML='(*)'
			span_indirizzo.innerHTML='(*)'
			span_citta.innerHTML='(*)'
			
			if (profilo.checked){
				span_cognome.innerHTML='(*)'
				span_nome.innerHTML='(*)'
			} else {
				span_cognome.innerHTML=''
				span_nome.innerHTML=''
			}
			
		}
		else{
			span_sede.innerHTML=''
			span_indirizzo.innerHTML=''
			span_citta.innerHTML=''
			
			span_cognome.innerHTML='(*)'
			span_nome.innerHTML='(*)'
		}
		
		
		if (profilo.checked){
			span_email.innerHTML='(*)'
			span_login.innerHTML='(*)'
			span_password.innerHTML='(*)'
			span_c_password.innerHTML='(*)'
		} else {
			span_email.innerHTML=''
			span_login.innerHTML=''
			span_password.innerHTML=''
			span_c_password.innerHTML=''
		}
		
		EnableIfChecked(profilo, form1.chk_abilitato);
		EnableIfChecked(profilo, form1.tft_login);
		EnableIfChecked(profilo, form1.old_login);
		EnableIfChecked(profilo, form1.tft_password);
		EnableIfChecked(profilo, form1.conferma_password);
		EnableIfChecked(profilo, form1.tfd_scadenza);
		
		if (input_tipo_sede.value!='' || input_codice_sede.value!=''){
			span_tipo_sede.innerHTML='(*)'
			span_codice_sede.innerHTML='(*)'
		} else {
			span_tipo_sede.innerHTML=''
			span_codice_sede.innerHTML=''
		}
	}
</script>


<div id="content_ridotto">
<form action="" method="post" id="form1" name="form1" onsubmit="set_modo_registra();">
	<input type="hidden" name="tft_modoregistra" value="">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<% 	if request("ID") = "" then 
			OBJ_contatto("CntRel") = request("CNT") %>
			<input type="hidden" name="tfn_CntRel" value="<%= OBJ_contatto("CntRel") %>">
			<caption>Inserimento nuovo contatto interno / sede alternativa</caption>
		<% 	else %>
			<input type="hidden" name="tfn_IDElencoIndirizzi" value="<%= request("ID") %>">
			<caption>Modifica dati del contatto interno / sede alternativa</caption>
		<% 	end if %>
		<tr><th colspan="4">ANAGRAFICA</th></tr>
		<tr>
			<td class="label" style="width:19%;">salva come:</td>
			<td class="content">
				<table border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td><input class="noBorder" type="radio" name="chk_isSocieta" id="chk_issocieta_false" value="" <%= chk(not OBJ_contatto("isSocieta"))%> onClick="show_mandatory()"></td>
						<td>contatto interno</td>
						<td style="padding-left:5px;"><input class="noBorder" type="radio" name="chk_isSocieta" id="chk_issocieta_true" value="1" <%= chk(OBJ_contatto("isSocieta"))%> onClick="show_mandatory()"></td>
						<td>sede alternativa</td>
					</tr>
				</table>
			</td>
			<td class="label">lingua:</td>
			<td class="content">
				<%CALL DropLingue(OBJ_contatto.conn, NULL, "tft_lingua", OBJ_contatto("lingua"), true, false, "") %>
			</td>
		</tr>
		<tr>
			<td class="label">sede:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_nomeorganizzazioneelencoindirizzi" value="<%= OBJ_contatto("nomeorganizzazioneelencoindirizzi") %>" maxlength="100" size="65">
				<span id="sede">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label" nowrap>nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_nomeelencoindirizzi" value="<%= OBJ_contatto("NomeElencoIndirizzi") %>" maxlength="100" size="55">
				<span id="nome">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">cognome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_cognomeelencoindirizzi" value="<%= OBJ_contatto("CognomeElencoIndirizzi") %>" maxlength="100" size="55">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">ruolo / qualifica:</td>
			<td class="content" colspan="3"><input type="text" class="text" name="tft_qualificaelencoindirizzi" value="<%= OBJ_contatto("qualificaelencoindirizzi") %>" maxlength="250" size="35"></td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<tr><th colspan="4">CODIFICA AZIENDALE</th></tr>
		<tr>
			<td class="label" style="width:19%;">tipo:</td>
			<% dim tmp
			set tmp = Server.CreateObject("Scripting.Dictionary") 
			tmp.CompareMode = vbTextCompare
			tmp.add "CfSede", "Sede"
			tmp.add "CfDest", "Destinazione diversa"%>
			<td class="content" style="width:35%;">
				<% CALL DropDownDictionary(Tmp, "tft_syncrotable", OBJ_contatto("SyncroTable"), false, " onChange=""show_mandatory()"" ", Session("LINGUA")) %>
				<span id="tipo_sede">(*)</span>
			</td>
			<td class="label">codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_syncrokey" value="<%= OBJ_contatto("SyncroKey") %>" maxlength="20" size="8" onChange="show_mandatory()" onkeyup="show_mandatory()">
				<span id="codice_sede">(*)</span>
			</td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<tr><th colspan="4">INDIRIZZO</th></tr>
		<tr>
			<td class="label">sede:</td>
			<td class="content" colspan="3">
				<% 
				sql = " SELECT IDElencoIndirizzi, " + _
					  " (NomeOrganizzazioneElencoIndirizzi " + SQL_concat(OBJ_contatto.conn) + "'  ('" + SQL_concat(OBJ_contatto.conn) + "IndirizzoElencoindirizzi" + SQL_concat(OBJ_contatto.conn) + " " + SQL_concat(OBJ_contatto.conn) + "CapElencoIndirizzi" + SQL_concat(OBJ_contatto.conn) + " " + SQL_concat(OBJ_contatto.conn) + "CittaElencoIndirizzi" + SQL_concat(OBJ_contatto.conn) + "')') AS NOME " + _
					  " FROM tb_indirizzario WHERE " & SQL_IsTrue(OBJ_contatto.conn, "IsSocieta") & " AND CntRel=" & OBJ_contatto("CntRel")
				CALL dropDown(OBJ_contatto.conn, sql, "IDElencoIndirizzi", "NOME", "tfn_cntsede", OBJ_contatto("CntSede"), false, "style=""width:100%;""", LINGUA_ITALIANO)
				%>
			</td>
		</tr>
		<tr>
			<td class="label" style="width:19%;">indirizzo:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_IndirizzoElencoIndirizzi" value="<%= OBJ_contatto("IndirizzoElencoIndirizzi") %>" maxlength="250" size="55">
				<span id="indirizzo">(*)</span>
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
				<span id="citta">(*)</span>
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
		<% if OBJ_contatto("tipoprofilo") AND cIntero(Obj_contatto("ut_id")) > 0 then
			sql = "SELECT COUNT(*) FROM rel_utenti_sito WHERE rel_ut_id=" & Obj_contatto("ut_id")
			EsitoOk = cInteger(GetValueList(Obj_contatto.conn, NULL, sql)) = 1
		else
			EsitoOk = true
		end if
		if EsitoOk then%>
			<tr>
				<td class="label" rowspan="2">tipo profilo:</td>
				<td class="content" colspan="2">
					<input class="checkbox" type="radio" name="chk_tipoprofilo" id="chk_tipoprofilo_false" value="" <%= chk(not OBJ_contatto("tipoprofilo") OR cString(OBJ_contatto("tipoprofilo"))="")%> onClick="show_mandatory()">
					senza profilo di accesso
				</td>
			</tr>
			<tr>
				<td class="content" colspan="2">
					<input class="checkbox" type="radio" name="chk_tipoprofilo" id="chk_tipoprofilo_true" value="1" <%= chk(OBJ_contatto("tipoprofilo"))%> onClick="show_mandatory()">
					con profilo di accesso subordinato
				</td>
			</tr>
		<% else %>
			<input type="Hidden" name="chk_tipoprofilo" id="chk_tipoprofilo_true" value="1">
			<tr>
				<td class="note" colspan="3">
					ATTENZIONE: Non &egrave; possibile rimuovere il profilo di accesso in quanto il contatto ha accesso anche ad altre applicazioni del sistema.					
				</td>
			</tr>
			<script language="JavaScript" type="text/javascript">
				form1.chk_tipoprofilo.checked = true;
			</script>
		<% end if %>
		<tr>
			<td class="label" style="width:23%;">stato:</td>
			<td class="content" style="width:32%;">
				<input type="checkbox" class="checkbox" name="chk_abilitato" <%= Chk(IIF(request("ID")<>"" OR request.serverVariables("REQUEST_METHOD")="POST", OBJ_contatto("abilitato"), true)) %>>
				abilitato all'accesso
			</td>
			<td class="note" rowspan="4">
				Per i valori di login e password utilizzare solo caratteri alfanumerici o &quot;_&quot; 
				indifferentemente con lettere minuscole o maiuscole, ma senza spazi bianchi.
				<span style="letter-spacing:2px;">(<%= replace(LOGIN_VALID_CHARSET, "s", "s ") %>)</span>
			</td>
		</tr>
		<tr>
			<td class="label">Login:</td>
			<td class="content">
				<input type="text" class="text" name="tft_login" value="<%= OBJ_contatto("login") %>" maxlength="50" size="15">
				<span id="login">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">Password:</td>
			<td class="content">
				<input type="password" class="text" name="tft_password" value="<%= OBJ_contatto("password") %>" maxlength="50" size="15">
				<span id="password">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">Conferma password</td>
			<td class="content">
				<input type="password" class="text" name="conferma_password" value="<%= OBJ_contatto("password") %>" maxlength="50" size="15">
				<span id="c_password">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">Scadenza accesso:</td>
			<td class="content" colspan="2">
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
	show_mandatory();
	FitWindowSize(this);
</script>