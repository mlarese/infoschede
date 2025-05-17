<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
dim rs, sql, OBJ_Contatto
set rs = Server.CreateObject("ADODB.RecordSet")
set OBJ_contatto = new IndirizzarioLock

if request.form("salva")<>"" then
	dim cnt_id, ut_id, mandatory_flds
	OBJ_contatto.LoadFromForm("isSocieta;abilitato")
	
	'controllo per login e password
	if OBJ_contatto.ValidateLoginAndPassword(request("old_login"), request("conferma_password")) then
		'controlla altri campi obbligatori
		if Obj_contatto("isSocieta") then
			mandatory_flds = "NomeOrganizzazioneElencoIndirizzi;"
		else
			mandatory_flds = "CognomeElencoIndirizzi;NomeElencoIndirizzi;"
		end if
		if OBJ_contatto.ValidateFields(mandatory_flds, true)	then
			'controllo esito positivo: dati validi
			OBJ_contatto.conn.beginTrans
			
			if cInteger(request("tfn_IDElencoindirizzi"))>0 then			'sono in modifica o ho scelto dal menu
				OBJ_contatto.UpdateDB()
				cnt_id = OBJ_contatto("IDElencoIndirizzi")
			else
				cnt_id = OBJ_contatto.InsertIntoDB()
			end if
			
			'registrazione utente
			ut_id = OBJ_Contatto.UserFromContact(cnt_id, 0)
			
			'blocca il contatto
			CALL OBJ_Contatto.LockContact(cnt_id, NEXTPASSPORT)
			
			CALL save_permessi(Obj_contatto.conn, rs, false, ut_id)
			
			'chiude transazione e conferma dati
			OBJ_contatto.conn.commitTrans
			
			response.redirect "Utenti.asp"
		end if
	end if
end if

if cInteger(request("tfn_IDElencoindirizzi"))>0 then
	OBJ_contatto.LoadFromDB(cInteger(request("tfn_IDElencoindirizzi")))
else
	OBJ_contatto.LoadFromForm("isSocieta;abilitato")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="Tools_Passport.asp" -->
<!--#INCLUDE FILE="../library/classIndirizzarioLock.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione utenti area riservata - nuovo utente"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Utenti.asp"
dicitura.scrivi_con_sottosez() 
%>
<script language="JavaScript" type="text/javascript">

	function show_mandatory(){
		var isSocieta = document.getElementById('chk_issocieta_true');
		var span_nome = document.getElementById('nome')
		var span_cognome = document.getElementById('cognome')
		var span_ente = document.getElementById('ente')

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
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo utente dell'area riservata</caption>
		
		<tr><th colspan="5">SELEZIONA L'UTENTE DAI CONTATTI</th></tr>
		<tr>
			<td class="label">Contatto:</td>
			<td class="content" colspan="4">
				<%
                sql = " IDElencoIndirizzi NOT IN (SELECT ut_nextCom_ID FROM tb_Utenti) AND (" + SQL_IsNULL(OBJ_contatto.conn, "cntRel") + " OR CntRel=0) "
                CALL WriteContactPicker_Input(OBJ_contatto.conn, rs, sql, "", "form1", "tfn_IDElencoIndirizzi", request("tfn_IDElencoindirizzi"), "EMAIL", false, false, false, "SUBMIT")
                %>
			</td>
		</tr>
		<tr><th colspan="5">DATI ANAGRAFICI</th></tr>
		<tr>
			<td class="label">salva come:</td>
			<td class="content" colspan="2">
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
			<td class="label">Organizzazone:</td>
			<td class="content" colspan="4">
				<input type="text" class="text" name="tft_NomeOrganizzazioneElencoIndirizzi" value="<%= OBJ_contatto("NomeOrganizzazioneElencoIndirizzi") %>" maxlength="250" size="100">
				<span id="ente">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">Cognome:</td>
			<td class="content" colspan="4">
				<input type="text" class="text" name="tft_cognomeelencoindirizzi" value="<%= OBJ_contatto("CognomeElencoIndirizzi") %>" maxlength="250" size="75">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label" style="width:18%;">Nome:</td>
			<td class="content" colspan="4">
				<input type="text" class="text" name="tft_nomeelencoindirizzi" value="<%= OBJ_contatto("NomeElencoIndirizzi") %>" maxlength="250" size="75">
				<span id="nome">(*)</span>
			</td>
		</tr>
		<script language="JavaScript" type="text/javascript">
			show_mandatory();
		</script>
		<tr><th colspan="5">INDIRIZZO</th></tr>
		<tr>
			<td class="label">indirizzo:</td>
			<td class="content" colspan="4"><input type="text" class="text" name="tft_IndirizzoElencoIndirizzi" value="<%= OBJ_contatto("IndirizzoElencoIndirizzi") %>" maxlength="250" size="100"></td>
		</tr>
		<tr>
			<td class="label">localit&agrave;:</td>
			<td class="content" colspan="2"><input type="text" class="text" name="tft_LocalitaElencoIndirizzi" value="<%= OBJ_contatto("LocalitaElencoIndirizzi") %>" maxlength="50" size="50"></td>
			<td class="label">cap:</td>
			<td class="content"><input type="text" class="text" name="tft_CAPElencoIndirizzi" value="<%= OBJ_contatto("CAPElencoIndirizzi") %>" maxlength="20" size="10"></td>
		</tr>
		<tr>
			<td class="label">citt&agrave;:</td>
			<td class="content" colspan="2"><input type="text" class="text" name="tft_cittaElencoIndirizzi" value="<%= OBJ_contatto("cittaElencoIndirizzi") %>" maxlength="50" size="35"></td>
			<td class="label">provincia / stato:</td>
			<td class="content"><input type="text" class="text" name="tft_StatoProvElencoIndirizzi" value="<%= OBJ_contatto("StatoProvElencoIndirizzi") %>" maxlength="50" size="20"></td>
		</tr>
		<tr><th colspan="5">RECAPITI</th></tr>
		<tr>
			<td class="label">telefono:</td>
			<td class="content" colspan="4">
				<input type="text" class="text" name="tft_telefono" value="<%= OBJ_contatto("telefono") %>" maxlength="250" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">fax:</td>
			<td class="content" colspan="4">
				<input type="text" class="text" name="tft_fax" value="<%= OBJ_contatto("fax") %>" maxlength="20" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">cellulare:</td>
			<td class="content" colspan="4">
				<input type="text" class="text" name="tft_cellulare" value="<%= OBJ_contatto("cellulare") %>" maxlength="20" size="50">
			</td>
		</tr>
		<tr>
			<td class="label">email:</td>
			<td class="content" colspan="4">
				<input type="text" class="text" name="tft_email" value="<%= OBJ_contatto("email") %>" maxlength="250" size="75">
				(*)
			</td>
		</tr>
		<tr><th colspan="5">ACCOUNT DI ACCESSO</th></tr>
		<tr>
			<td class="label">stato:</td>
			<td class="content" colspan="4">
				<input type="checkbox" class="checkbox" name="abilitato" <%= chk(OBJ_contatto("abilitato")) %>>
				abilitato all'accesso
			</td>
		</tr>
		<tr>
			<td class="label">Login:</td>
			<td class="content">
				<input type="text" class="text" name="tft_login" value="<%= OBJ_contatto("login") %>" maxlength="50" size="20">
				(*)
			</td>
			<td class="label">Scadenza:</td>
			<td class="content" colspan="2">
				<% CALL WriteDataPicker_Input("form1", "tfd_scadenza", OBJ_contatto("Scadenza"), "", "/", true, true, LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">Password:</td>
			<td class="content">
				<input type="password" class="text" name="tft_password" value="<%= OBJ_contatto("password") %>" maxlength="50" size="20">
				(*)
			</td>
			<td class="note" colspan="3" rowspan="2" style="width:60%;">
				Per i valori di login e password utilizzare solo caratteri alfanumerici o &quot;_&quot; 
				indifferentemente con lettere minuscole o maiuscole, ma senza spazi bianchi.
				<span style="letter-spacing:2px;">(<%= LOGIN_VALID_CHARSET %>)</span>
			</td>
		</tr>
		<tr>
			<td class="label">Conferma password:</td>
			<td class="content">
				<input type="password" class="text" name="conferma_password" value="<%= request("conferma_password") %>" maxlength="50" size="20">
				(*)
			</td>
		</tr>
		<tr><th colspan="5">PROFILO DI ACCESSO</th></tr>
		<tr>
			<td colspan="5">
				<% CALL write_permessi(OBJ_contatto.conn, rs, false, 0, 3) %>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="5">
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
<% set rs = nothing
set OBJ_contatto = nothing
%>