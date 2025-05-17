<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("OrdiniSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione ordini - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Ordini.asp"
dicitura.scrivi_con_sottosez()

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.recordset")
conn.open Application("DATA_ConnectionString")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfd_ord_data_ins" value="NOW">
	<input type="hidden" name="tfd_ord_data_ultima_mod" value="NOW">
	<input type="hidden" name="tfn_ord_movimenta" value="0">
	<input type="hidden" name="tfn_ord_impegna" value="0">
	<input type="hidden" name="tfn_ord_archiviato" value="0">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>Inserimento nuovo ordine</caption>
		<tr><th colspan="4">DATI DEL'ORDINE</th></tr>
		<tr>
			<td class="label">data:</td>
			<td class="content" style="width:25%;"><% CALL WriteDataPicker_Input("form1", "tfd_ord_data", Date(), "", "/", false, true, LINGUA_ITALIANO) %></td>
			<td class="label">riferimento:</td>
			<td class="content">
				<input type="text" class="text" name="tft_ord_cod" value="<%= request("tft_ord_cod") %>" maxlength="50" size="20">
			</td>
		</tr>
		<tr>
			<td class="label">cliente:</td>
			<td class="content" colspan="3">
				<table cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td>
						<input type="hidden" name="tfn_ord_riv_id" value="<%= request.form("tfn_ord_riv_id") %>">
						<input READONLY type="text" name="cliente" style="padding-left:3px; width:100%" value="<%= request.form("cliente") %>" 
							   onclick="OpenAutoPositionedScrollWindow('ClientiSelezione.asp?field_nome=cliente&field_id=tfn_ord_riv_id&selected=' + tfn_ord_riv_id.value, 'SelezioneCliente', 450, 480, true)" title="Click per aprire la finestra per la selezione del cliente">
					</td>
					<td width="25%">
						<a class="button_input" href="javascript:void(0)" onclick="form1.cliente.onclick();" 
							 title="Apre la filnestra per la selezione del cliente" <%= ACTIVE_STATUS %>>
							SELEZIONA CLIENTE
						</a>
						&nbsp;(*)
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr><th colspan="4">GESTIONE DEL'ORDINE</th></tr>
		<tr>
			<td class="label">magazzino:</td>
			<td class="content" colspan="3">
				<% sql = "SELECT * FROM gtb_magazzini ORDER BY mag_nome"
				CALL dropDown(conn, sql, "mag_id", "mag_nome", "tfn_ord_magazzino_id", request("tfn_ord_magazzino_id"), true, "", LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">stato dell'ordine:</td>
			<td colspan="3">
				<table cellpadding="0" cellspacing="1" width="100%">
					<tr>
						<td width="22%" class="content<%= STILI_STATI_ORDINE(ORDINE_NON_CONFERMATO) %>">
							<input type="radio" class="checkbox" name="stato_ordine" value="<%= ORDINE_NON_CONFERMATO %>" <%= chk(cInteger(request("stato_ordine"))=ORDINE_NON_CONFERMATO) %>>
							<%= STATI_ORDINE(ORDINE_NON_CONFERMATO) %>
						</td>
						<td class="content<%= STILI_STATI_ORDINE(ORDINE_NON_CONFERMATO) %>" colspan="2">
							<% sql = "SELECT * FROM gtb_stati_ordine WHERE so_stato_ordini=" & ORDINE_NON_CONFERMATO & " ORDER BY so_ordine"
							CALL dropDown(conn, sql, "so_id", "so_nome_it", "tfn_ord_stato_id", request("tfn_ord_stato_id"), true, " ", LINGUA_ITALIANO) %>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr><th colspan="4">IMPORTO SPESE</th></tr>
		<tr>
			<td class="label">spese di spedizione:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tfn_ord_spesespedizione" value="<%= request("tfn_ord_spesespedizione") %>" maxlength="255" size="5">
				&euro;
			</td>
		</tr>
		<tr>
			<td class="label">spese fisse:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tfn_ord_spesefisse" value="<%= request("tfn_ord_spesefisse") %>" maxlength="255" size="5">
				&euro;
			</td>
		</tr>
		<tr>
			<td class="label">spese incasso:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tfn_ord_speseincasso" value="<%= request("tfn_ord_speseincasso") %>" maxlength="255" size="5">
				&euro;
			</td>
		</tr>
		<tr>
			<td class="label">altre spese:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tfn_ord_spesealtre" value="<%= request("tfn_ord_spesealtre") %>" maxlength="255" size="5">
				&euro;
			</td>
		</tr>
		<%
		dim n_porti, n_tipo_consegna, n_trasportatori
		n_porti = cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM gtb_porti"))
		n_tipo_consegna = cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM gtb_tipo_consegna"))
		n_trasportatori = cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM gtb_trasportatori"))
		%>
		<% if n_porti > 0 OR n_tipo_consegna > 0 OR n_trasportatori > 0 then %>
			<tr><th colspan="4">TRASPORTO E CONSEGNA</th></tr>
		<% end if %>
		<% if n_porti > 0 then %>
			<tr>
				<td class="label">porto:</td>
				<td class="content" colspan="3">
					<% 	sql = " SELECT * FROM gtb_porti ORDER BY prt_nome_it"
					dropDown conn, sql, "prt_id", "prt_nome_it", "tfn_ord_porto_id", request("tfn_ord_porto_id"), false, "", LINGUA_ITALIANO %>
				</td>
			</tr>
		<% end if %>
		<% if n_tipo_consegna > 0 then %>
			<tr>
				<td class="label">modalit&agrave; consegna:</td>
				<td class="content" colspan="3">
					<% 	sql = " SELECT * FROM gtb_tipo_consegna ORDER BY tco_ordine, tco_nome_it"
					dropDown conn, sql, "tco_id", "tco_nome_it", "tfn_ord_tipo_consegna_id", request("ord_tipo_consegna_id"), false, "", LINGUA_ITALIANO %>
				</td>
			</tr>
		<% end if %>
		<% if n_trasportatori > 0 then %>
			<tr>
				<td class="label">trasportatore:</td>
				<td class="content" colspan="3">
					<% 	sql = " SELECT * FROM gtb_trasportatori ORDER BY tra_nome_it"
					dropDown conn, sql, "tra_id", "tra_nome_it", "tfn_ord_trasportatore_id", request("tfn_ord_trasportatore_id"), false, "", LINGUA_ITALIANO %>
				</td>
			</tr>
		<% end if %>
	</table>
	<% 
	'......................................................................................................
	'FUNZIONE DICHIARATA NEL FILE DI CONFIGURAZIONE DEL CLIENTE
	CALL ADDON__ORDINI__form_insert(conn, rs)
	'......................................................................................................
	%>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<tr><th colspan="4">ARTICOLI ORDINATI</th></tr>
		<tr><td class="note" colspan="4">L'immissione degli articoli dell'ordine sar&agrave; disponibile dopo aver salvato.</td></tr>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="3" name="tft_ord_note"><%= request("tft_ord_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="AVANTI &gt;&gt;">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
