<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = true %>
<!--#INCLUDE FILE="../library/classIndirizzarioLock.asp" -->
<%

dim standAlone
if request("STANDALONE")="true" then
	standAlone = true
else
	standAlone = false
end if


if cString(request("PROFILO"))<>"" then
	Session("PROFILO") = request("PROFILO")
end if

if cString(request("IDPROFILO"))<>"" then
	Session("IDPROFILO") = request("IDPROFILO")
end if
id_profilo = cIntero(Session("IDPROFILO"))

dim conn, rs, rsr, rsa, sql, OBJ_Contatto, id_ext, profilo, profiloS, colore, id_profilo
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")
set OBJ_contatto = new IndirizzarioLock


'se arrivo dalla scelta del cliente per una scheda, dopo aver inserito l'anagrafica chiudo 
'la finestra e scrivo sul form della scheda i dati del cliente
if request("WRITE_DATA_AND_CLOSE") = "true" then
	ut_id = cIntero(request("ID"))
	sql = "SELECt * FROM gv_rivenditori WHERE ut_id = " & ut_id
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	%>
	<script language="JavaScript" type="text/javascript">
		opener.document.form1.<%=request("field_id")%>.value = <%=ut_id%>;
		opener.document.form1.cliente.value = '<%=ContactFullName(rs)%>';
		opener.document.form1.submit();
		window.close();
	</script>
	<%
	rs.close
	response.end
end if


if request("IDCNT") <> "" OR request("ID") <> "" then
	OBJ_contatto.LoadFromDB(cInteger(request("IDCNT") & request("ID")))
	if request("ID") <> "" then
		id_ext = GetValueList(conn, rs, "SELECT riv_id FROM gv_rivenditori WHERE IDElencoIndirizzi="& cIntero(request("ID")))
	end if
else
	OBJ_contatto.LoadFromForm("isSocieta;chk_abilitato")
end if

'gestione dei campi esterni a IndirizzarioLock come se fossero interni
CaricaCampiEsterni conn, rs, OBJ_contatto, "SELECT * FROM gtb_rivenditori", "riv_id", id_ext

if request("ID") <> "" then
	id_profilo = cIntero(OBJ_contatto("riv_profilo_id"))
end if

profilo = Session("PROFILO")
if profilo = "anagrafiche_clienti" then
	sql = "SELECT pro_nome_it FROM gtb_profili WHERE pro_id = " & id_profilo
	profilo = GetValueList(conn, NULL, sql)
end if
sql = "SELECT pro_codice FROM gtb_profili WHERE pro_id = " & id_profilo
profiloS = GetValueList(conn, NULL, sql)


'salvataggio
if request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim pagina_redirect
	pagina_redirect = "Clienti.asp?PROFILO=" & Session("PROFILO")
	if StandAlone then
		pagina_redirect = "ClientiGestione.asp?WRITE_DATA_AND_CLOSE=true&PROFILO="&ParseSQL(request("PROFILO"), adChar)&"&IDPROFILO="&ParseSQL(request("IDPROFILO"), adChar)&"&field_id="&request("field_id")
	end if
	sql = "SELECT pro_rubrica_id FROM gtb_profili WHERE pro_id = " & cIntero(request("id_profilo"))
	id_rubrica = GetValueList(conn, NULL, sql)
	
	sql = Replace(sql,"pro_rubrica_id","pro_nome_it")
	log_desc = GetValueList(conn, NULL, sql)
	
	permesso_area_riservata = GetPermessoUtente(id_ext)

	write_log = true
	log_desc = "Infoschede - " & log_desc
	%>
	<!--#INCLUDE FILE="../nextB2B/ClientiGestioneSalva_Tools.asp" -->
	<%	
end if

'cambio record (precedente o successivo)
if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_"&Session("PROFILO")&"_SQL"), "IDElencoIndirizzi", "ClientiGestione.asp?PROFILO="&Session("PROFILO"))
end if


%>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione " & LCase(Replace(profilo,"_", " "))
if not standAlone then
	dicitura.puls_new = "INDIETRO"
	dicitura.link_new = "Clienti.asp?PROFILO=" & Session("PROFILO")
else
	dicitura.puls_new = ""
	dicitura.link_new = ""
end if
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
	<input type="Hidden" name="old_login" value="<%= OBJ_contatto("login") %>">
	<input type="hidden" name="id_profilo" value="<%= id_profilo %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
	<%
	colore = ""
	if request("PROFILO")="anagrafiche_clienti" then
		select case id_profilo
			case CLIENTI_PRIVATI
				colore = "style=""background:" & COLOR_CLIENTI_PRIVATI & """"
			case CLIENTI_PROFESSIONALI
				colore = "style=""background:" & COLOR_CLIENTI_PROFESSIONALI & """"
			case RIVENDITORI
				colore = "style=""background:" & COLOR_RIVENDITORI & """"
			case SUPERVISORI_NEGOZI
				colore = "style=""background:" & COLOR_SUPERVISORI_NEGOZI & """"
			case else
				colore = ""
		end select
	end if
	%>
	
	<% 	if request("ID") = "" then %>
		<caption <%=colore%>>Inserimento nuovo <%=profiloS%></caption>
		<input type="hidden" name="tfn_cnt_insAdmin_id" value="<%= Session("ID_ADMIN") %>">
		<input type="hidden" name="tfd_cnt_insData" value="<%= Now() %>">
		<input type="hidden" name="tfn_cnt_modAdmin_id" value="<%= Session("ID_ADMIN") %>">
		<input type="hidden" name="tfd_cnt_modData" value="<%= Now() %>">
	<% 	else %>
		<input type="hidden" name="tfn_cnt_modAdmin_id" value="<%= Session("ID_ADMIN") %>">
		<input type="hidden" name="tfd_cnt_modData" value="<%= Now() %>">
		<caption <%=colore%>>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati del <%=profiloS%></td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= cIntero(request("ID")) %>&goto=PREVIOUS" title="cliente precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= cIntero(request("ID")) %>&goto=NEXT" title="cliente successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
	<% 	end if
        if request("ID") = "" then %>
			<% if (id_profilo=CLIENTI_PRIVATI OR id_profilo=RIVENDITORI) AND Session("INFOSCHEDE_ADMIN")<>"" then %>
				<tr><th colspan="4">SELEZIONA IL <%=UCase(profiloS)%> DAI CONTATTI</th></tr>
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
			<% end if %>
		<% 	else %>
			<input type="hidden" name="tfn_IDElencoIndirizzi" value="<%= request("ID") %>">
		<% 	end if %>
		<tr><th colspan="4">DATI ANAGRAFICI</th></tr>
		<tr>
			<td class="label">salva come:</td>
			<td class="content">
				<table border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<% dim start_societa
						start_societa = false
						if id_profilo = COSTRUTTORI OR id_profilo = TRASPORTATORI then start_societa = true
						%>
						<td><input class="noBorder" type="radio" name="isSocieta" id="chk_issocieta_false" value="" <%=chk((not OBJ_contatto("isSocieta")) AND not start_societa)%> onClick="show_mandatory()"></td>
						<td width="30%">persona fisica</td>
						<td><input class="noBorder" type="radio" name="isSocieta" id="chk_issocieta_true" value="1" <%=chk((OBJ_contatto("isSocieta")) OR start_societa)%> onClick="show_mandatory()"></td>
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
			<td class="label" style="width:19%;">ente / societ&agrave;:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_NomeOrganizzazioneElencoIndirizzi" value="<%= OBJ_contatto("NomeOrganizzazioneElencoIndirizzi") %>" maxlength="250" size="90">
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
			<td class="label">codice fiscale:</td>
			<td class="content">
				<input type="text" class="text" name="tft_CF" value="<%= OBJ_contatto("CF") %>" maxlength="16" size="16">
			</td>
			<td class="label">partita i.v.a.:</td>
			<td class="content"><input type="text" class="text" name="tft_partita_iva" value="<%= OBJ_contatto("partita_iva") %>" maxlength="11" size="14"></td>
		</tr>
		<tr>
			<td class="label">codice aziendale:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="extT_riv_codice" value="<%= OBJ_contatto("riv_codice") %>" maxlength="20" size="10">
			</td>
		</tr>
		<tr><th colspan="4">INDIRIZZO PRINCIPALE</th></tr>
		<tr>
			<td class="label">indirizzo:</td>
			<td class="content" colspan="3"><input type="text" class="text" name="tft_IndirizzoElencoIndirizzi" value="<%= OBJ_contatto("IndirizzoElencoIndirizzi") %>" maxlength="250" size="90"></td>
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
			<td class="label">provincia / stato:</td>
			<td class="content"><input type="text" class="text" name="tft_StatoProvElencoIndirizzi" value="<%= OBJ_contatto("StatoProvElencoIndirizzi") %>" maxlength="50" size="20"></td>
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
				<input type="text" class="text" name="tft_email" value="<%= OBJ_contatto("email") %>" maxlength="250" size="75">
				(*)
			</td>
		</tr>
		<% if request("ID") <> "" then %>
			<% if id_profilo = RIVENDITORI then
				sql = " ut_id IN (SELECT riv_id FROM gtb_rivenditori WHERE riv_profilo_id = "&SUPERVISORI_NEGOZI&") "
				%>
				<tr>
					<th colspan="4">SUPERVISORE</th>
				</tr>
				<tr>
					<td class="label" style="width:19%;">supervisore negozio:</td>
					<td class="content" colspan="3">
						<%
						CALL WriteContactPicker_Input(OBJ_contatto.conn, rs, sql, "", "form1", "extN_riv_azienda_capogruppo_id", OBJ_contatto("riv_azienda_capogruppo_id"), "LOGIN LOGINID EMAIL", false, false, false, "")
						%>
					</td>
				</tr>
			<% elseif id_profilo = SUPERVISORI_NEGOZI then 
				sql = "SELECT * FROM gv_rivenditori WHERE riv_profilo_id = "&RIVENDITORI&" AND riv_azienda_capogruppo_id = "&OBJ_contatto("riv_id") 
				rsr.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
				
				if not rsr.eof then %>
					<tr>
						<th colspan="4">RIVENDITORI SUPERVISIONATI</th>
					</tr>
					<tr>
						<td colspan="4">
							<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
								<tr>
									<th class="L2">rivenditore</th>
									<th class="L2" width="40%">indirizzo</th>
								</tr>
								<% while not rsr.eof %>
									<tr>
										<td class="content"><%= ContactFullName(rsr) %></td>
										<td class="content"><%= ContactAddress(rsr) %></td>
									</tr>
									<% rsr.movenext
								wend %>
							</table>
						</td>
					</tr>
				<% end if %>
				<% rsr.close %>
			<% end if %>
		<% end if %>
		
		<% if request("ID") <> "" then 'gestione contatti interni
			sql = " SELECT *, (SELECT COUNT(*) FROM gtb_dettagli_ord WHERE det_ind_id = tb_indirizzario.IDElencoIndirizzi) AS N_ORD, " + _
				  " (SELECT ut_id FROM tb_utenti WHERE ut_nextCom_id = tb_indirizzario.IDElencoIndirizzi) AS UTENTE, " + _
				  " (SELECT COUNT(*) FROM rel_utenti_sito INNER JOIN tb_utenti ON rel_utenti_sito.rel_ut_id = tb_utenti.ut_id " + _
				  " WHERE tb_utenti.ut_nextCom_id = tb_indirizzario.IDElencoIndirizzi) AS N_PERMESSI " + _
				  " FROM tb_Indirizzario INNER JOIN tb_cnt_lingue ON tb_Indirizzario.lingua = tb_cnt_lingue.lingua_codice " + _
				  " WHERE CntRel=" & cIntero(request("ID")) & " ORDER BY ModoRegistra "
			rsr.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
			%>
			<tr>
				<th colspan="4">CONTATTI INTERNI / SEDI DIVERSE</th>
			</tr>
			<tr>
				<td colspan="4">
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr>
							<td class="label" style="width:30%">
								<% if rsr.eof then %>
									Nessun contatto / sede inseriti.
								<% else %>
									Trovati n&ordm; <%= rsr.recordcount %> record
								<% end if %>
							</td>
							<td colspan="6" class="content_right" style="padding-right:0px;">
								<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ClientiIntGestione.asp?CNT=<%= cIntero(request("ID")) %>', 'cntInt', 550, 450, true)">
									NUOVO CONTATTO INTERNO
								</a>
							</td>
						</tr>
						<% if not rsr.eof then %>
							<tr>
								<th class="L2">contatto</th>
								<th class="L2" width="25%">e-mail</th>
								<th class="L2" width="35%">indirizzo</th>
								<th class="L2" width="6%">&nbsp;</th>
								<th class="l2_center" width="14%" colspan="2">operazioni</th>
							</tr>
							<% while not rsr.eof %>
								<tr>
									<td class="content" title="ruolo / qualifica: <%= rsr("QualificaElencoIndirizzi") %>"><%= ContactFullName(rsr) %></td>
									<td class="content">
										<% sql = " SELECT ValoreNumero FROM tb_ValoriNumeri WHERE email_default = 1 AND id_TipoNumero = 6 " & _
												 " AND id_Indirizzario = " & rsr("IDElencoIndirizzi") %>
										<%= GetValueList(conn, NULL, sql) %>
									</td>
									<td class="content">
										<%= ContactAddress(rsr) %>
									</td>
									<td class="content_center">
										<% if rsr("lingua")<>"" then %>
											<img src="../grafica/flag_mini_<%= rsr("lingua") %>.jpg" alt="Lingua: <%= rsr("lingua_nome_it") %>">
										<% end if %>
									</td>
									<td class="content_center">
										<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ClientiIntGestione.asp?ID=<%= rsr("IdElencoIndirizzi") %>', 'cntInt', 550, 450, true)">
											MODIFICA
										</a>
									</td>
									<td class="content_center">
										<% if cInteger(rsr("N_ORD"))>0 then %>
											<a class="button_L2_disabled" title="impossibile cancellare l'indirizzo altenativo perch&egrave; &egrave; associato ad almeno una riga d'ordine">
												CANCELLA
											</a>
										<% elseif cInteger(rsr("N_PERMESSI"))>1 then %>
											<a class="button_L2_disabled" title="impossibile cancellare il contatto associato all'indirizzo in quanto ha accesso anche ad altre applicazioni dell'area riservata.">
												CANCELLA
											</a>
										<% else %>
											<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('CLIENTI_INDIRIZZI','<%= rsr("IDElencoIndirizzi") %>');">
												CANCELLA
											</a>
										<% end if %>
									</td>
								</tr>
								<%rsr.movenext
							wend
						end if%>
					</table>
				</td>
			</tr>
			<% rsr.close 
		end if %>

	</table>
	
	<input type="hidden" name="listino_associato" id="listino_associato_esistente" value="<%=GetValueList(conn,NULL,"SELECT TOP 1 listino_id FROM gtb_listini")%>">
	<input type="hidden" name="extn_riv_valuta_id" value="1">
	<input type="hidden" name="extn_riv_modopagamento_id" value="0">
	<input type="hidden" name="extn_riv_profilo_id" value="<%=id_profilo%>">
		
	<% if id_profilo = COSTRUTTORI then %>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">	
			<tr><th colspan="4">PROFILO COMMERCIALE</th></tr>
			<tr>
				<td class="label" style="width:19%;">sconto:</td>
				<td class="content">
					<input type="text" class="text" name="extn_riv_sconto" value="<%= FormatPrice(OBJ_contatto("riv_sconto"), 2, false) %>" maxlength="5" size="5">
					%
				</td>
				<td class="note">
					Percentuale di sconto applicata ai pezzi di ricambio di questo costrutture.
				</td>
			</tr>
		</table>
	<% end if %>
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<% if Session("PROFILO") = "trasportatori" then %>
			<input type="hidden" name="chk_abilitato" value="">
			<input type="hidden" name="tft_login" value="<%=RANDOM_LOGIN_E_PASSWORD%>" >
			<input type="hidden" name="tft_password" value="<%=RANDOM_LOGIN_E_PASSWORD%>">
			<input type="hidden" name="conferma_password" value="<%=RANDOM_LOGIN_E_PASSWORD%>">
		<% else %>
			<tr><th colspan="4">PROFILO DI ACCESSO</th></tr>
			<tr>
				<td class="label" style="width:19%; line-height:20px;">profilo:</td>
				<td class="content_b" <%=colore%>  style="line-height:20px;">
					<%= profilo %>
				</td>
				<td class="content_right" style="padding-right:0px;" colspan="2">
					<% if request("ID") <> "" then %>
						<a class="button" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ClientiCambiaProfilo.asp?RIV_ID=<%=OBJ_contatto("riv_id")%>&PRO_ID=<%=OBJ_contatto("riv_profilo_id")%>', 'cntCambiaProfilo', 300, 350, true)">
							CAMBIA PROFILO
						</a>
					<% else %>
						&nbsp;
					<% end if %>
				</td>
			</tr>
			<tr>
				<td class="label" style="width:19%;">stato:</td>
				<td class="content" colspan="3">
					<input type="checkbox" class="checkbox" name="chk_abilitato" <%= Chk(IIF(request("ID")<>"" OR request.serverVariables("REQUEST_METHOD")="POST", OBJ_contatto("abilitato"), true)) %>>
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
				<td class="content">
					<% CALL WriteDataPicker_Input("form1", "tfd_scadenza", OBJ_contatto("Scadenza"), "", "/", true, true, LINGUA_ITALIANO) %>
				</td>
			</tr>
			<tr>
				<td class="label">Password:</td>
				<td class="content">
					<input type="text" class="text" name="tft_password" value="<%= OBJ_contatto("password") %>" maxlength="50" size="20">
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
					<input type="text" class="text" name="conferma_password" value="<%= OBJ_contatto("password") %>" maxlength="50" size="20">
					(*)
				</td>
			</tr>
		<% end if %>
		<tr>
			<% if request("ID") <> "" then %>
				<% rs.open "SELECT * FROM tb_indirizzario WHERE IDElencoIndirizzi="&cIntero(request("ID")), OBJ_contatto.conn, adOpenStatic, adLockReadOnly, adCmdText %>
				<% CALL Form_DatiModifica(OBJ_contatto.conn, rs, "cnt_") %>	
				<% rs.close %>
			<% end if %>
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

<% if standAlone then %>
	<script language="JavaScript" type="text/javascript">
	<!--
		FitWindowSize(this);
	//-->
	</script>
<% end if %>


<% conn.close
set rs = nothing
set rsa = nothing
set rsr = nothing
set conn = nothing
%>