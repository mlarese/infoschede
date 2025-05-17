<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
dim rs, sql, OBJ_Contatto, mandatory_flds
set rs = Server.CreateObject("ADODB.RecordSet")
set OBJ_contatto = new IndirizzarioLock

if request("goto")<>"" then
	CALL GotoRecord(OBJ_contatto.conn, rs, Session("SQL_UTENTI"), "IDElencoIndirizzi", "UtentiMod.asp")
end if

if request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim ut_id, val
	
	OBJ_contatto.LoadFromForm("isSocieta;abilitato")
	sql = "SELECT ut_password FROM tb_utenti WHERE ut_NextCom_id=" & cIntero(request("ID"))
	OBJ_contatto("password") = GetValueList(OBJ_contatto.conn, rs, sql)
	
	'controllo per login e password
	if OBJ_contatto.ValidateLoginAndPassword(request("old_login"), obj_contatto("password")) then
		'controlla altri campi obbligatori
		if Obj_contatto("isSocieta") then
			mandatory_flds = "NomeOrganizzazioneElencoIndirizzi;"
		else
			mandatory_flds = "CognomeElencoIndirizzi;NomeElencoIndirizzi;"
		end if
		if OBJ_contatto.ValidateFields(mandatory_flds, true)	then
			'controllo esito positivo: dati validi
			OBJ_contatto.conn.beginTrans
			
			OBJ_contatto.UpdateDB()
			
			'registrazione utente
			if Session("PERMESSO_AREA_RISERVATA") <> "" then
				ut_id = OBJ_Contatto.UserFromContact(request("ID"), Session("PERMESSO_AREA_RISERVATA"))
			else
				ut_id = OBJ_Contatto.UserFromContact(request("ID"), 0)
			end if

			'inserimento relazioni tra profili e utente
			sql = "DELETE FROM mrel_profili_utenti WHERE rpu_utenti_id = " & ut_id
			Obj_contatto.conn.Execute(sql)
			for each val in Split(request.form("profili_associati"), ",")
				if CIntero(val) > 0 then
					sql = " INSERT INTO mrel_profili_utenti(rpu_utenti_id, rpu_profilo_id)"& _
						  " VALUES (" & ut_id & ", " & val & ")"
					Obj_contatto.conn.Execute(sql)
				end if
			next
			val = ""
			
			'chiude transazione e conferma dati
			OBJ_contatto.conn.commitTrans
			
			if request("salva")<>"" then
				response.redirect "UtentiMod.asp?ID="&request("ID")
			else
				response.redirect "Utenti.asp"
			end if
		end if
	end if
end if

OBJ_contatto.LoadFromDB(request("ID"))
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../nextPassport/Tools_Passport.asp" -->
<!--#INCLUDE FILE="../library/classIndirizzarioLock.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione utenti area riservata - modifica utente"
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
	<input type="hidden" name="tfn_IDElencoindirizzi" value="<%= request("ID") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati dell'utente</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="utente precedente">
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="utente successiva">
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
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
			<td class="content" colspan="2">
				<input type="checkbox" class="checkbox" name="abilitato" <%= chk(OBJ_contatto("abilitato")) %>>
				abilitato all'accesso
			</td>
			<% if cIntero(Session("PAGINA_AVVISO_ABILITAZIONE_UTENTE"))>0 then %>
				<td class="content" colspan="2">
					<a <%=IIF(cBoolean(OBJ_contatto("abilitato"),0),"","disabled")%> class="button_L2<%=IIF(cBoolean(OBJ_contatto("abilitato"),0),"","_disabled")%>"
					   style="width:100%;text-align:center;"
					   href="../nextPassport/UtentiAvvisoAttivazione.asp?ID=<%=request("ID")%>"
					   onclick="OpenAutoPositionedScrollWindow('', 'attivazione_<%=request("ID")%>', 840, 500, true)"
					   target="attivazione_<%=request("ID")%>" 
					   title="<%=IIF(cBoolean(OBJ_contatto("abilitato"),0),"Spedisci avviso attivazione all'utente", "Potrai spedire l'avviso dopo aver abilitato l'utente all'accesso.") %>"
					   <%= ACTIVE_STATUS %>>
						Spedisci avviso di attivazione all'utente
					</a>
				</td>
			<% else %>
				<td class="content" colspan="2">&nbsp;</td>
			<% end if %>
		</tr>
		<tr>
			<td class="label">Login:</td>
			<td class="content">
				<input type="text" class="text" name="tft_login" value="<%= OBJ_contatto("login") %>" maxlength="50" size="20">
				(*)
				<input type="hidden" name="old_login" value="<%= OBJ_contatto("login") %>">
			</td>
			<td class="label">Scadenza:</td>
			<td class="content" colspan="2">
				<% CALL WriteDataPicker_Input("form1", "tfd_scadenza", OBJ_contatto("Scadenza"), "", "/", true, true, LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">Password:</td>
			<td class="content" colspan="4" style="padding:2px;">
				<a href="javascript:void(0);" class="button_form" onclick="OpenAutoPositionedWindow('UtentiPassword.asp?ID=<%= Obj_contatto("ut_id") %>', 'pws', 402, 240);">
					MODIFICA LA PASSWORD
				</a>
			</td>
		</tr>
		<% sql = "SELECT pro_id FROM mtb_profili" %>
		<% if (cBoolean(Session("CONDIVISIONE_INTERNA"), false) OR cBoolean(Session("CONDIVISIONE_PUBBLICA"), false)) _ 
				AND cString(GetValueList(OBJ_contatto.conn, NULL, sql)) <> "" then %>
			<tr><th colspan="5">PROFILI PER LA VISUALIZZAZIONE DEI DOCUMENTI / CIRCOLARI</th></tr>
			<tr>
				<td class="content" colspan="5">
					<% sql = " SELECT * FROM mtb_profili LEFT JOIN mrel_profili_utenti " + _
							 " ON (mtb_profili.pro_id = mrel_profili_utenti.rpu_profilo_id AND " + _
							 "					mrel_profili_utenti.rpu_utenti_id = " & Obj_contatto("ut_id") & ")" + _
							 " ORDER BY pro_nome_it"
					   CALL Write_Relations_Checker(OBJ_contatto.conn, rs, sql, 4, "pro_id", "pro_nome_it", "rpu_id", "profili_associati")%>
				</td>
			</tr>
		<% end if %>
		
		<tr>
			<td class="footer" colspan="5">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
				<input type="submit" class="button" name="salva_elenco" value="SALVA & TORNA ALL'ELENCO">
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