<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ContattiSalva.asp")
end if

dim conn, rs, sql, rubriche_visibili, value

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

'recupera rubriche visibili all'utente
rubriche_visibili = GetList_Rubriche(conn, rs)

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = "Anagrafica contatti - nuovo"
'Indirizzo pagina per link su sezione 
		HREF = "Contatti.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
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

	function ShowDatiAggiuntivi(state){
		if (document.getElementById("Agg1").style.visibility == "visible" || state == "hide"){
			document.getElementById("Agg1").style.visibility = 'hidden';
			document.getElementById("Agg1").style.display = 'none';
			document.getElementById("Agg2").style.visibility = 'hidden';
			document.getElementById("Agg2").style.display = 'none';
			document.getElementById("Agg3").style.visibility = 'hidden';
			document.getElementById("Agg3").style.display = 'none';
			document.getElementById("PulsanteAgg").innerHTML = 'Mostra dati aggiuntivi';
		}
		else
		{
			document.getElementById("Agg1").style.visibility = 'visible';
			document.getElementById("Agg1").style.display = '';
			document.getElementById("Agg2").style.visibility = 'visible';
			document.getElementById("Agg2").style.display = '';
			document.getElementById("Agg3").style.visibility = 'visible';
			document.getElementById("Agg3").style.display = '';
			document.getElementById("PulsanteAgg").innerHTML = 'Nascondi dati aggiuntivi';
		}
	}
	</script>
<div id="content">
	<form action="" method="post" id="form1" name="form1" onsubmit="set_modo_registra();">
	<input type="hidden" name="tft_modoregistra" value="">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Inserimento nuovo contatto</td>
					<td align="right" style="font-size:10px;">
						<a href="javascript:void(0)" class="button_L2" onclick="ShowDatiAggiuntivi('');" id="PulsanteAgg">Mostra dati aggiuntivi</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">ANAGRAFICA</th></tr>
		<tr>
			<td class="label_no_width" style="width:20%;">salva come:</td>
			<td class="content" style="width:42%;">
				<table border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td><input class="noBorder" type="radio" name="chk_isSocieta" id="chk_issocieta_false" value="" <%= chk(request("chk_isSocieta")<>"1")%> onClick="show_mandatory()"></td>
						<td width="30%">persona fisica</td>
						<td><input class="noBorder" type="radio" name="chk_isSocieta" id="chk_issocieta_true" value="1" <%= chk(request("chk_isSocieta")="1")%> onClick="show_mandatory()"></td>
						<td>ente / societ&agrave; / organizzazione</td>
					</tr>
				</table>
			</td>
			<td class="label_no_width" style="width:18%;">lingua comunicazioni:</td>
			<td class="content">
				<% CALL DropLingue(conn, rs, "tft_lingua", request("tft_lingua"), true, false, "width:100%;") %>
			</td>
		</tr>
		<tr>
			<td class="label_no_width">ente:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_nomeorganizzazioneelencoindirizzi" value="<%= request("tft_NomeOrganizzazioneElencoIndirizzi") %>" maxlength="250" style="width:95%;">
				<span id="ente">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label_no_width">titolo:</td>
			<td class="content" colspan="3"><input type="text" class="text" name="tft_TitoloElencoIndirizzi" value="<%= request("tft_TitoloElencoIndirizzi") %>" maxlength="50" style="width:22%;"></td>
		</tr>
		<tr>
			<td class="label_no_width">nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_nomeelencoindirizzi" value="<%= request("tft_NomeElencoIndirizzi") %>" maxlength="100" style="width:70%;">
				<span id="nome">(*)</span>
			</td>
		</tr>
		<tr id="Agg1">
			<td class="label_no_width">secondo nome:</td>
			<td class="content" colspan="3"><input type="text" class="text" name="tft_secondonomeelencoindirizzi" value="<%= request("tft_SecondoNomeElencoIndirizzi") %>" maxlength="100" style="width:70%;"></td>
		</tr>
		<tr>
			<td class="label_no_width">cognome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_cognomeelencoindirizzi" value="<%= request("tft_CognomeElencoIndirizzi") %>" maxlength="100" style="width:70%;">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<tr id="Agg2">
			<td class="label_no_width">ruolo / qualifica:</td>
			<td class="content" colspan="3"><input type="text" class="text" name="tft_qualificaelencoindirizzi" value="<%= request("tft_qualificaelencoindirizzi") %>" maxlength="250" style="width:52%;"></td>
		</tr>
		<tr>	
			<td class="label_no_width">codice fiscale:</td>
			<td class="content"><input type="text" class="text" name="tft_CF" value="<%= request("tft_CF") %>" maxlength="16" style="width:60%;"></td>
			<td class="label_no_width">partita i.v.a.:</td>
			<td class="content"><input type="text" class="text" name="tft_partita_iva" value="<%= request("tft_partita_iva") %>" maxlength="11" style="width:100%;"></td>
		</tr>
		<tr id="Agg3">
			<td class="label_no_width">luogo di nascita:</td>
			<td class="content"><input type="text" class="text" name="tft_luogonascita" value="<%= request("tft_luogonascita") %>" maxlength="255" style="width:100%;"></td>
			<td class="label_no_width">data di nascita:</td>
			<td class="content"><input type="text" class="text" name="tfd_dtnascelencoindirizzi" value="<%= request("tfd_DTNASCElencoIndirizzi") %>" maxlength="10" style="width:100%;"></td>
		</tr>
		<% if Session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE") then %>
			<tr>
				<td class="label">categoria:</td>
				<td class="content" colspan="3">
					<%CALL dropDown(conn, CatContatti.QueryElenco(true, ""), "icat_id", "NAME", "tfn_cnt_categoria_id", cInteger(request("tfn_cnt_categoria_id")), false, " onchange='form1.submit()'", LINGUA_ITALIANO)%>
				</td>
			</tr>
		<% end if %>
		<script language="JavaScript" type="text/javascript">
			show_mandatory();
		</script>
		<tr><th colspan="4">INDIRIZZO</th></tr>
		<tr>
			<td class="label_no_width">indirizzo:</td>
			<td class="content" colspan="3"><input type="text" class="text" name="tft_IndirizzoElencoIndirizzi" value="<%= request("tft_IndirizzoElencoIndirizzi") %>" maxlength="250" style="width:100%;"></td>
		</tr>
		<tr>
			<td class="label_no_width">localit&agrave;:</td>
			<td class="content"><input type="text" class="text" name="tft_LocalitaElencoIndirizzi" value="<%= request("tft_LocalitaElencoIndirizzi") %>" maxlength="50" style="width:100%;"></td>
			<td class="label_no_width">cap:</td>
			<td class="content"><input type="text" class="text" name="tft_CAPElencoIndirizzi" value="<%= request("tft_CAPElencoIndirizzi") %>" maxlength="20" style="width:100%;"></td>
		</tr>
		<tr>
			<td class="label_no_width">citt&agrave;:</td>
			<td class="content"><input type="text" class="text" name="tft_cittaElencoIndirizzi" value="<%= request("tft_cittaElencoIndirizzi") %>" maxlength="50" style="width:100%;"></td>
			<td class="label_no_width" nowrap>provincia / stato:</td>
			<td class="content"><input type="text" class="text" name="tft_StatoProvElencoIndirizzi" value="<%= request("tft_StatoProvElencoIndirizzi") %>" maxlength="50" style="width:100%;"></td>
		</tr>
		<tr>
			<td class="label_no_width">zona:</td>
			<td class="content"><input type="text" class="text" name="tft_ZonaElencoIndirizzi" value="<%= request("tft_ZonaElencoIndirizzi") %>" maxlength="50" style="width:100%;"></td>
			<td class="label_no_width">nazione:</td>
			<td class="content"><input type="text" class="text" name="tft_CountryElencoIndirizzi" value="<%= request("tft_CountryElencoIndirizzi") %>" maxlength="50" style="width:100%;"></td>
		</tr>
		<tr><th colspan="4">RUBRICHE (*)</th></tr>
		<tr>
			<td colspan="4">
				<% sql = "SELECT *, (NULL) AS id_rub_ind FROM tb_rubriche WHERE tb_rubriche.id_rubrica " &_
						 " IN (" & GetList_Rubriche(conn, rs) & ") " &_
						 " AND NOT(" & SQL_IsTrue(conn, "tb_rubriche.rubrica_esterna") & ") " &_
		  				 " ORDER BY nome_rubrica"
				CALL Write_Relations_Checker(conn, rs, sql, 3, "id_rubrica", "nome_rubrica", "id_rub_ind", "rubriche") %>
			</td>
		</tr>
		<% If InStr(Application("NextCom_codice"), "<PREFISSOCLIENTE>") > 0 AND Application("NextCrm") then %>
			<tr><th colspan="4">PARAMETRI GESTIONE PRATICHE</th></tr>
			<tr>
				<td class="label_no_width">prefisso:</td>
				<td class="content"><input type="text" class="text" name="tft_PraticaPrefisso" value="<%= request("tft_PraticaPrefisso") %>" maxlength="5" style="width:100%;"></td>
				<td class="note" colspan="2">
					Prefisso per la generazione del codice di ogni pratica.
				</td>
			</tr>
		<% End If
		
		dim newsletter_on
		
		sql = "SELECT * FROM tb_TipNumeri" 
		rs.open sql, conn, adOpenForwardOnly, adLockOptimistic
		if not rs.eof then%>
			<tr><th colspan="4">RECAPITI</th></tr>
			<% while not rs.eof 
				Select case rs("id_TipoNumero")
					case VAL_URL
						value = 75
					case VAL_EMAIL
						value = 75
					case else
						value = 50
				end select
				
				if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) AND rs("id_TipoNumero") = VAL_EMAIL then
					newsletter_on = true
				else
					newsletter_on = false
				end if
				%>
				<tr>
					<td class="label_no_width"><%= rs("nome_tiponumero") %>:</td>
					<td class="content" <%=IIF(newsletter_on,"colspan=""2""","colspan=""3""")%>><input type="text" class="text" name="recapito_<%= rs("id_TipoNumero") %>" value="<%= request("recapito_" & rs("id_TipoNumero")) %>" maxlength="250" style="width:<%= value %>%;"></td>
					<% if newsletter_on then %>
						<td class="Content">
							<input type="checkbox" class="noBorder" name="email_newsletter" value="true" />
							Usa per le newsletter
						</td>
					<% end if %>
				</tr>
				<% rs.movenext
			wend
		end if
		rs.close %>
		
		<% if Session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE") then %>
			<tr><th colspan="4">CARATTERISTICHE</th></tr>
		<%	sql = " SELECT * FROM (tb_indirizzario_carattech " + _
				  "	INNER JOIN rel_categ_ctech ON (tb_indirizzario_carattech.ict_id = rel_categ_ctech.rcc_ctech_id " & _
				  "						AND rel_categ_ctech.rcc_categoria_id=" & cInteger(request("tfn_cnt_categoria_id")) & ") )" + _
				  " LEFT JOIN rel_cnt_ctech ON (tb_indirizzario_carattech.ict_id=rel_cnt_ctech.ric_ctech_id AND rel_cnt_ctech.ric_cnt_id="& CInteger(request("ID")) &") " + _
				  " LEFT JOIN tb_indirizzario_carattech_raggruppamenti ON tb_indirizzario_carattech.ict_raggruppamento_id = tb_indirizzario_carattech_raggruppamenti.icr_id " + _
				  " ORDER BY tb_indirizzario_carattech_raggruppamenti.icr_ordine, rel_categ_ctech.rcc_ordine "
			CALL DesForm  (conn, sql, "tb_indirizzario_carattech", "ict_id", "ict_nome_it", "ict_tipo", "ict_unita_it", "", "ric_valore_", "ric_valore_", "icr_titolo_it", cIntero(request("ID")) = 0, 4)
		end if%>

		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="6" name="tft_NoteElencoIndirizzi"><%=request("tft_NoteElencoIndirizzi")%></textarea>
			</td>
		</tr>

		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input style="width:25%;" type="submit" class="button" name="salva_elenco" value="SALVA & TORNA ALL'ELENCO">
				<input style="width:14%;" type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
	<script language="JavaScript" type="text/javascript">
		ShowDatiAggiuntivi('hide');
	</script>
</div>
</body>
</html>
<% 
conn.close 
set rs = nothing
set conn = nothing%>
