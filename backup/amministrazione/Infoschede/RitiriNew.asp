<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
dim post
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	post = true
else
	post = false
end if

if post AND (request("salva")<>"" OR request("salva_elenco")<>"") then
	Server.Execute("RitiriSalva.asp")
end if

dim conn, rs, rsc, sql, label, id_cliente
set conn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.recordset")
set rsc = Server.CreateObject("ADODB.recordset")
conn.open Application("DATA_ConnectionString")

sql = "SELECT cat_nome_it FROM sgtb_ddt_categorie WHERE cat_id = " & cIntero(request("CAT_ID"))
label = GetValueList(conn, NULL, sql)


id_cliente = cIntero(request("ID_CLIENTE"))


%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione " & lCase(label) & " - inserimento "
if id_cliente>0 then
	dicitura.puls_new = ""
	dicitura.link_new = ""
else
	dicitura.puls_new = "INDIETRO"
	dicitura.link_new = "Ritiri.asp"
end if
dicitura.scrivi_con_sottosez()

%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_ddt_categoria_id" value="<%= request("CAT_ID") %>">
	<input type="hidden" name="nuovo_inserimento" value="true">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>Inserimento <%=lCase(label)%></caption>
		<tr><th colspan="4">DATI <%=UCase(label)%></th></tr>
		<tr>
			<td class="label" style="width:18%;">numero:</td>
			<td class="content" colspan="3">
				-----
			</td>
		</tr>
		<tr>
			<td class="label">data:</td>
			<td class="content" colspan="3">
				<% CALL WriteDataPicker_Input("form1", "tfd_ddt_data", IIF(Request.ServerVariables("REQUEST_METHOD")="POST",request("tfd_ddt_data"),Date()), "", "/", false, true, LINGUA_ITALIANO) 
				%>
			</td>
		</tr>
		<tr>
			<% if id_cliente > 0 then %>
				<% sql = " SELECT * FROM tb_Indirizzario INNER JOIN tb_Utenti " & _
						 " ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID " & _
						 " WHERE ut_ID = " & id_cliente
				rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				%>
				<input type="hidden" name="reload" value="true">
				<td class="label">cliente:</td>
				<td class="content" colspan="3"><%= ContactFullName(rsc)%></td>
				<input type="hidden" name="tfn_ddt_cliente_id" value="<%=id_cliente%>">
			<% else %>
				<td class="label">cliente:</td>
				<td class="content" colspan="3">
					<table cellpadding="0" cellspacing="0" width="100%">
						<tr>
							<td>
								<% dim filtro_profili
								filtro_profili = "&filtro_profilo="&TRASPORTATORI&","&COSTRUTTORI&"&filtro_exclude=true"
								%>
								<input type="hidden" name="tfn_ddt_cliente_id" value="<%= request.form("tfn_ddt_cliente_id") %>">
								<input READONLY type="text" name="cliente" style="padding-left:3px; width:100%" value="<%= request.form("cliente") %>" 
									   onclick="OpenAutoPositionedScrollWindow('ClientiSelezione.asp?field_nome=cliente&field_id=tfn_ddt_cliente_id&selected=' + tfn_ddt_cliente_id.value + '&field_destinazione_id=tfn_ddt_destinazione_id&field_destinazione_nome=destinazione<%=filtro_profili%>', 'SelezioneCliente', 620, 480, true)" title="Click per aprire la finestra per la selezione del cliente">
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
			<% end if %>
		</tr>
		<tr>
			<td class="label">destinazione:</td>
			<td class="content" colspan="3">
				<table cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td>
						<% dim destinazione_id, destinazione
						if id_cliente > 0 then
							destinazione_id = rsc("IDElencoIndirizzi")
							destinazione = ContactAddress(rsc)
						else
							destinazione_id = request.form("tfn_ddt_destinazione_id")
							destinazione = request.form("destinazione")
						end if
						%>
						<input type="hidden" name="tfn_ddt_destinazione_id" value="<%= destinazione_id %>">
						<input READONLY type="text" name="destinazione" style="padding-left:3px; width:100%" value="<%= destinazione %>" 
							   onclick="OpenAutoPositionedScrollWindow('ClientiSelezione.asp?field_nome=destinazione&field_id=tfn_ddt_destinazione_id&selected=' + tfn_ddt_destinazione_id.value, 'SelezioneDestinazione', 620, 520, true)" title="Click per aprire la finestra per la selezione della destinazione">
					</td>
					<td width="30%" nowrap>
						<a class="button_input" href="javascript:void(0)" onclick="form1.destinazione.onclick();" 
							 title="Apre la filnestra per la selezione del destinazione" <%= ACTIVE_STATUS %>>
							SELEZIONA DESTINAZIONE
						</a>
						&nbsp;(*)
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td class="label">trasportatore:</td>
			<td class="content" colspan="3">
				<table cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td>
						<input type="hidden" name="tfn_ddt_trasportatore_id" value="<%= request.form("tfn_ddt_trasportatore_id") %>">
						<input READONLY type="text" name="trasportatore" style="padding-left:3px; width:100%" value="<%= request.form("trasportatore") %>" 
							   onclick="OpenAutoPositionedScrollWindow('ClientiSelezione.asp?field_nome=trasportatore&field_id=tfn_ddt_trasportatore_id&selected=' + tfn_ddt_trasportatore_id.value + '&filtro_profilo=<%=TRASPORTATORI%>', 'SelezioneTrasportatore', 620, 480, true)" title="Click per aprire la finestra per la selezione del trasportatore">
					</td>
					<td width="32%" nowrap>
						<a class="button_input" href="javascript:void(0)" onclick="form1.trasportatore.onclick();" 
							 title="Apre la filnestra per la selezione del trasportatore" <%= ACTIVE_STATUS %>>
							SELEZIONA TRASPORTATORE
						</a>
						&nbsp;(*)
					</td>
				</tr>
				</table>
			</td>	
		</tr>
		<tr>
			<td class="label">peso:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_ddt_peso" value="<%=request("tft_ddt_peso")%>" maxlength="255" size="22">
			</td>
		</tr>
		<tr>
			<td class="label">volume:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_ddt_volume" value="<%=request("tft_ddt_volume")%>" maxlength="255" size="22">
			</td>
		</tr>
		<tr>
			<td class="label">numero colli:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_ddt_numero_colli" value="<%=request("tft_ddt_numero_colli")%>" maxlength="255" size="22">
			</td>
		</tr>
		<% if (cIntero(request.form("tfn_ddt_cliente_id"))>0 AND post) OR id_cliente > 0 then %>
			<% sql = " SELECT * FROM sgtb_schede INNER JOIN gv_articoli ON sgtb_schede.sc_modello_id = gv_articoli.rel_id " & _
					 " INNER JOIN sgtb_stati_schede ON sgtb_schede.sc_stato_id = sgtb_stati_schede.sts_id " & _
					 " WHERE ISNULL(sts_elenco_ddt_da_ritirare,0)=1 AND ISNULL(sc_documento_ritiro_id, 0)=0 AND sc_cliente_id = "
					 
			if id_cliente > 0 AND not post then
				sql = sql & id_cliente
			else
				sql = sql & cIntero(request.form("tfn_ddt_cliente_id"))
			end if
			
			dim id_scheda
			id_scheda = cIntero(request("ID_SCHEDA"))			
 
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
				<tr><th colspan="5">SCHEDE ASSOCIATE ALLA RICHIESTA DI RITIRO</th></tr>
				<% if rs.eof then %>
					<tr><td colspan="4" class="note">Nessuna scheda da ritirare per il cliente selezionato</td></tr>
				<% else %>
					<tr>
						<th class="l2_center" style="width:6%">associa</th>
						<th class="l2_center" style="width:18%">numero scheda e data</th>
						<th class="l2_center" style="width:18%">stato</th>
						<th class="l2_center" style="width:14%">costo ritiro</th>
						<th class="L2">modello</th>
					</tr>
				<% end if %>
				<% while not rs.eof %>
					<tr>
						<td class="content_center">
							<input type="checkbox" class="noBorder<%=IIF(rs("sc_id")=id_scheda," checked","")%>" name="id_schede" value="<%=rs("sc_id")%>" 
									<%= chk(cBoolean(inStr(", "&request("id_schede")&",",", "&rs("sc_id")&",")>0 OR rs("sc_id")=id_scheda,false))%>>
						</td>
						<td class="content_center"><% CALL SchedaLink(rs("sc_id"), rs("sc_numero") & " del " & rs("sc_data_ricevimento"))%></td>
						<td class="content"><%=rs("sts_nome_it")%></td>
						<td class="content_center">
							<input type="text" class="number" name="costo_ritiro_scheda_<%=rs("sc_id")%>" value="<%= FormatPrice(cReal(request("costo_ritiro_scheda_"&rs("sc_id"))), 2, false) %>" size="7"> &euro;
						</td>
						<td class="content">
							<% CALL ArticoloLink(rs("art_id"), rs("art_nome_it"), rs("art_cod_int")) %>
							<% if rs("art_varianti") then %>
								<%= ListValoriVarianti(conn, rsd, rs("rel_id")) %>
							<% else %>
								&nbsp;
							<% end if %>
						</td>
					</tr>
				<% rs.moveNext %>
			<% wend %>
			</table>
		<% end if %>

		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<tr><th colspan="4">NOTE</th></tr>
			<tr>
				<td class="content" colspan="4">
					<textarea style="width:100%;" rows="4" name="tft_ddt_note"><%= request("tft_ddt_note") %></textarea>
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="4">
					(*) Campi obbligatori.
					<input type="submit" class="button" name="salva" value="SALVA">
					<% if not id_cliente > 0 then %>
						<input type="submit" class="button" name="salva_elenco" value="SALVA & TORNA A ELENCO">
					<% end if %>
				</td>
			</tr>
		</table>
		&nbsp;
	</form>
</div>
</body>
</html>


<% if id_cliente>0 then 
	rsc.close() %>
	<script language="JavaScript" type="text/javascript">
		FitWindowSize(this);
	</script>
<% end if %>

<%
set rs = nothing
set rsc = nothing
conn.Close
set conn = nothing
%>
