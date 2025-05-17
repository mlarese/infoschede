<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if (request("salva")<>"" OR request("salva_continua")<>"") AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SchedeSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/classIndirizzarioLock.asp" -->

<% 	
dim dicitura, data
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione richieste di assistenza - nuovo"
dicitura.puls_new = "INDIETRO"
if request("ID_CENTRO_ASSISTENZA")<>"" then
	dicitura.link_new = "Schede.asp?ASSEGNATA=true"
else
	dicitura.link_new = "Schede.asp?ASSEGNATA=false"
end if
dicitura.scrivi_con_sottosez() 

dim conn, rs, rsa, sql, i, ordine, is_centro_ass
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")


if cString(Session("INFOSCHEDE_CENTRO_ASSISTENZA"))<>"" then
	is_centro_ass = true
else
	is_centro_ass = false
end if

%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="nuovo_inserimento" value="true">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova scheda</caption>
		<tr><th colspan="6">DATI PRINCIPALI</th></tr>
		<% sql = "SELECT * FROM sgtb_stati_schede ORDER BY sts_ordine, sts_nome_it" %>
		<tr>
			<td class="label">stato scheda:</td>
			<td class="content" colspan="5">
				<% CALL dropDown(conn, sql, "sts_id", "sts_nome_it", "tfn_sc_stato_id", request("tfn_sc_stato_id"), true, "", Session("LINGUA")) %>
				(*)
			</td>
		</tr>
		<tr>
			<td class="label" style="width:20%;">numero:</td>
			<td class="content" colspan="5">
				------
			</td>
		</tr>
		<tr>
			<td class="label">data ricevimento:</td>
			<td class="content" colspan="5">
				<% CALL WriteDataPicker_Input_Manuale("form1", "tfd_sc_data_ricevimento", IIF(cString(request("tfd_sc_data_ricevimento"))="",Date(),request("tfd_sc_data_ricevimento")), "", "/", false, true, LINGUA_ITALIANO, "", true, "") 
				%>
			</td>
		</tr>
		<% if request("ID_CENTRO_ASSISTENZA")<>"" then %>
			<tr>
				<td class="label">centro assistenza:</td>
				<td class="content_b" colspan="5" style="width:65%;">
					<% sql = " SELECT * FROM gtb_agenti INNER JOIN tb_Utenti ON gtb_agenti.ag_id = tb_Utenti.ut_ID INNER JOIN " & _
							 "               tb_Indirizzario ON tb_Utenti.ut_NextCom_ID = tb_Indirizzario.IDElencoIndirizzi" & _
							 " WHERE ag_id = " & cIntero(request("ID_CENTRO_ASSISTENZA"))
					rsa.open sql, conn
					response.write ContactFullName(rsa) 
					%>
				</td>
			</tr>
			<input type="hidden" name="tfn_sc_centro_assistenza_id" value="<%=rsa("ag_id")%>">
			<% rsa.close %>
		<% else %>
			<tr>
				<td class="label">centro assistenza:</td>
				<td class="content" colspan="3">
					<%
					dim id_centro_ass_default
					if cIntero(request("tfn_sc_centro_assistenza_id")) > 0 then
						id_centro_ass_default = request("tfn_sc_centro_assistenza_id")
					else
						id_centro_ass_default = cIntero(GetValueList(conn, NULL, "SELECT TOP 1 ag_id FROM gtb_agenti WHERE ag_supervisore = 1 ORDER BY ag_id "))
					end if
					sql = " ut_id IN (SELECT ag_id FROM gv_agenti) "
					CALL WriteContactPicker_Input(conn, rs, sql, "", "form1", "tfn_sc_centro_assistenza_id", id_centro_ass_default, "LOGIN LOGINID EMAIL", false, true, false, "")
					%>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label" <%=IIF(cIntero(request.form("tfn_sc_cliente_id")) > 0,"rowspan=""2""","")%>>cliente:</td>
			<td class="content" colspan="5">
				<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
						<td>
							<input type="hidden" name="tfn_sc_cliente_id" id="tfn_sc_cliente_id" value="<%= request.form("tfn_sc_cliente_id") %>">
							<input READONLY type="text" name="cliente" id="cliente" style="padding-left:3px; width:100%" value="<%= request.form("cliente") %>" 
								   onclick="OpenAutoPositionedScrollWindow('ClientiSelezione.asp?field_nome=cliente&field_id=tfn_sc_cliente_id&selected=' + tfn_sc_cliente_id.value + '&filtro_profilo=<%=TRASPORTATORI&","&COSTRUTTORI%>&filtro_exclude=true&BUTTONS_ADD=true&CENTRO_ASSISTENZA_ID=<%=request("ID_CENTRO_ASSISTENZA")%>&AFTER=submit', 'SelezioneCliente', 620, 480, true)" 
										title="Click per aprire la finestra per la selezione del cliente">
						</td>
						<td width="28%">
							<a class="button_input" href="javascript:void(0)" onclick="form1.cliente.onclick();" 
								 title="Apre la filnestra per la selezione del cliente" <%= ACTIVE_STATUS %> style="display:inline;">
								SCEGLI
							</a>
							<% if cIntero(request.form("tfn_sc_cliente_id")) > 0 then
								sql = "SELECT * FROM gv_rivenditori WHERE riv_id = " & cIntero(request.form("tfn_sc_cliente_id"))
								rsa.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
								%>
								<a class="button_input" onclick="OpenAutoPositionedScrollWindow('ClientiGestione.asp?ID=<%=rsa("IDElencoIndirizzi")%>&PROFILO=anagrafiche_clienti&STANDALONE=true', 'DatiCliente', 500, 500, true)" 
									 href="javascript:void(0)"  title="Apre la finestra per la visualizzazione o la modifica dei dati del cliente" <%= ACTIVE_STATUS %>
									 style="display:inline;">
									VISUALIZZA DATI
								</a>
							<% end if %>
							&nbsp;(*)
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<% if cIntero(request.form("tfn_sc_cliente_id")) > 0 then %>
			<tr>
				<td colspan="5">
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr><th class="L2" colspan="2">DATI ANAGRAFICA</th></tr>
						<tr>
							<td class="label">indirizzo:</td>
							<td class="content"><%=ContactAddress(rsa)%></td>
						</tr>
						<tr>
							<td class="label">codice fiscale:</td>
							<td class="content"><%=rsa("CF")%></td>
						</tr>
						<tr>
							<td class="label">p. iva:</td>
							<td class="content"><%=rsa("partita_iva")%></td>
						</tr>
						<% dim Obj_Cnt 
						set Obj_Cnt = new IndirizzarioLock 
						Obj_Cnt.LoadFromDB(rsa("IDElencoIndirizzi"))
						%>
						<tr>
							<td class="label">telefono:</td>
							<td class="content"><%=Obj_Cnt("telefono")%></td>
						</tr>
						<tr>
							<td class="label">fax:</td>
							<td class="content"><%=Obj_Cnt("fax")%></td>
						</tr>
						<tr>
							<td class="label">cellulare:</td>
							<td class="content"><%=Obj_Cnt("cellulare")%></td>
						</tr>
						<tr>
							<td class="label">e-mail:</td>
							<td class="content"><a href="mailto:<%= Obj_Cnt("email") %>"><%= Obj_Cnt("email") %></a></td>
						</tr>
					</table>
				</td>
			</tr>
			<% rsa.close %>
		<% end if %>
		
		<tr>
			<td class="label">note del cliente:</td>
			<td class="content" colspan="5">
				<textarea style="width:100%;" rows="5" name="tft_sc_note_cliente"><%= request("tft_sc_note_cliente") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="label">riferimento cliente:</td>
			<td class="content" colspan="5">
				<input type="text" class="text" name="tft_sc_rif_cliente" value="<%= request("tft_sc_rif_cliente") %>" maxlength="255" style="width:100%;">
			</td>
		</tr>
		<tr><th colspan="6">DATI DEL MODELLO</th></tr>
		<tr>
			<td class="label" <% if cIntero(request("tfn_sc_modello_id")) = MODELLO_DEFAULT then %>rowspan="2"<%end if%> >modello:</td>
			<td class="content" colspan="5">
				<% CALL WritePicker_ArticoloVariante(conn, rs, "form1", "tfn_sc_modello_id", request("tfn_sc_modello_id"), 89, true, "Infoschede/ArticoliSeleziona.asp?SUBMIT_AFTER=true&TYPE=M&") %>
			</td>
		</tr>
		<% if cIntero(request("tfn_sc_modello_id")) > 0 then
			sql = "SELECT rel_cod_int FROM grel_art_valori WHERE rel_id = " & cIntero(request("tfn_sc_modello_id"))
			dim codice
			codice = cString(GetValueList(conn, NULL, sql))
			%>
			<tr>
				<td class="label">codice:</td>
				<td class="content" colspan="3">
					<input type="text" READONLY class="text disabled" name="sc_codice" value="<%= codice %>" maxlength="255" size="52">
				</td>
			</tr>
		<% end if %>
		<% if cIntero(request("tfn_sc_modello_id")) = MODELLO_DEFAULT then %>
			<tr>
				<td class="content" colspan="5">
					nome modello:&nbsp;
					<input type="text" class="text" name="tft_sc_modello_altro" value="<%= request("tft_sc_modello_altro") %>" maxlength="500" size="91">
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label">matricola:</td>
			<td class="content" colspan="5">
				<input type="text" class="text" name="tft_sc_matricola" value="<%= request("tft_sc_matricola") %>" maxlength="255" size="52">
			</td>
		</tr>
		
		<tr><th colspan="6">DATI DI ACQUISTO</th></tr>
		<tr>
			<td class="label">negozio di acquisto:</td>
			<td class="content" colspan="5">
				<input type="text" class="text" name="tft_sc_negozio_acquisto" value="<%= request("tft_sc_negozio_acquisto") %>" maxlength="255" size="52">
			</td>
		</tr>
		<tr>
			<td class="label">data acquisto:</td>
			<td class="content" colspan="5">
				<% CALL WriteDataPicker_Input_Manuale("form1", "tfd_sc_data_acquisto", request("tfd_sc_data_acquisto"), "", "/", true, true, LINGUA_ITALIANO, "", true, "")
				%>
			</td>
		</tr>
		<tr>
			<td class="label">numero scontrino:</td>
			<td class="content" colspan="5">
				<input type="text" class="text" name="tft_sc_numero_scontrino" value="<%= request("tft_sc_numero_scontrino") %>" maxlength="100" size="52">
			</td>
		</tr>
		<% if is_centro_ass then %>
			<tr>
				<td class="label">richiesta di garanzia:</td>
				<td class="content" colspan="5">
					<input type="checkbox" class="noBorder" name="chk_sc_richiesta_garanzia" value="1" <%= chk(cBoolean(request("chk_sc_richiesta_garanzia")<>"",false))%>>
					<input type="hidden" name="chk_sc_in_garanzia" value="">
				</td>
			</tr>
		<% else %>
			<tr>
				<td class="label">garanzia:</td>
				<td class="content" colspan="5">
					<input type="checkbox" class="noBorder" name="chk_sc_in_garanzia" value="1" <%= chk(cBoolean(request("chk_sc_in_garanzia")<>"",false))%>>
					<input type="hidden" name="chk_sc_richiesta_garanzia" value="">
				</td>
			</tr>
		<% end if %>
				
		<% sql = "SELECT * FROM sgtb_accessori ORDER BY acc_nome_it" %>
		<tr>
			<td class="label" rowspan="2">accessori presenti:</td>
			<td class="content" colspan="5">
				<% CALL dropDown(conn, sql, "acc_id", "acc_nome_it", "tfn_sc_accessori_presenti_id", request("tfn_sc_accessori_presenti_id"), false, "style=""width:50%;""", Session("LINGUA")) %>
			</td>
		</tr>
		<tr>
			<td class="content" colspan="5">
				altro&nbsp;
				<input type="text" class="text" name="tft_sc_accessori_presenti_altro" value="<%= request("tft_sc_accessori_presenti_altro") %>" maxlength="500" style="width:100%;">
			</td>
		</tr>
		
		<tr><th colspan="6">DATI DELLA RIPARAZIONE</th></tr>
		<% if cIntero(request("tfn_sc_modello_id")) > 0 then
			dim marca_id, tipologia_id, rel_art_id
			rel_art_id = cIntero(request("tfn_sc_modello_id"))
			marca_id = CIntero(GetValueList(conn, NULL, "SELECT art_marca_id FROM gv_articoli WHERE rel_id = "&rel_art_id))
			tipologia_id = CIntero(GetValueList(conn, NULL, "SELECT art_tipologia_id FROM gv_articoli WHERE rel_id = "&rel_art_id))
			sql = "SELECT DISTINCT prb_id, prb_nome_it FROM sgtb_problemi " & _
                 " LEFT JOIN srel_problemi_articoli ON prb_id = rpa_problema_id" & _
                 " LEFT JOIN srel_problemi_mar_tip ON prb_id = rpm_problema_id" & _
                 " LEFT JOIN grel_art_valori ON rel_id = rpa_articolo_rel_id" & _
                     " WHERE ISNULL(prb_riscontrato, 0)=0 " & _
                       " AND prb_visibile = 1" & _
                       " AND ((rpa_problema_id = prb_id AND rel_art_id =" & rel_art_id & ")" & _
                         " OR (rpm_problema_id = prb_id" & _
                       " AND (rpm_tipologia_id = " & tipologia_id &" OR rpm_tipologia_id = 0)" & _
                       " AND (rpm_marchio_id = " & marca_id & " OR rpm_marchio_id = 0)))"
		  else
			sql = "SELECT * FROM sgtb_problemi WHERE ISNULL(prb_riscontrato, 0)=0 ORDER BY prb_nome_it" 
		  end if		
		
		%>
		<tr>
			<td class="label" rowspan="2">guasto segnalato:</td>
			<td class="content" colspan="5">
				<% CALL dropDown(conn, sql, "prb_id", "prb_nome_it", "tfn_sc_guasto_segnalato_id", request("tfn_sc_guasto_segnalato_id"), false, "style=""width:50%;""", Session("LINGUA")) %>
			</td>
		</tr>
		<tr>
			<td class="content" colspan="5">
				altro&nbsp;
				<input type="text" class="text" name="tft_sc_guasto_segnalato_altro" value="<%= request("tft_sc_guasto_segnalato_altro") %>" maxlength="500" style="width:100%;">
			</td>
		</tr>

		<tr><th colspan="6">DATI DEL TRASPORTO</th></tr>
		<tr><th class="L2" colspan="6">presa</th></tr>
		<tr>
			<td class="label">numero DDT di carico:</td>
			<td class="content" colspan="5">
				<input type="text" class="text" name="tft_sc_numero_DDT_di_carico" value="<%= request("tft_sc_numero_DDT_di_carico") %>" maxlength="255" size="52">
			</td>
		</tr>
				<tr>
			<td class="label">data DDT di carico:</td>
			<td class="content" colspan="5">
				<% CALL WriteDataPicker_Input_Manuale("form1", "tfd_sc_data_DDT_di_carico", request("tfd_sc_data_DDT_di_carico"), "", "/", false, true, LINGUA_ITALIANO, "", true, "") 
				%>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="6">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
				<input type="submit" class="button" name="salva_continua" value="SALVA & TORNA ALL'ELENCO">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>

<script language="JavaScript" type="text/javascript">

</script>


