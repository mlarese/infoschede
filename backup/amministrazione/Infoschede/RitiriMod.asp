<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" AND (request("salva")<>"" OR request("salva_elenco")<>"") then
	Server.Execute("RitiriSalva.asp")
end if

dim conn, rs, rsd, sql, rsi, label, standalone
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.Recordset")
set rsi = Server.CreateObject("ADODB.Recordset")

if request("STANDALONE") = "true" then
	standalone = true
else
	standalone = false
end if

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("INFOSCHEDE_RITIRI_SQL"), "ddt_ID", "RitiriMod.asp")
end if

%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	

sql = " SELECT *, (SELECT COUNT(sc_id) FROM sgtb_schede WHERE sc_rif_DDT_di_resa_id = sgtb_ddt.ddt_id) AS N_IND_DIV " + _
	  " FROM sgtb_ddt INNER JOIN gv_rivenditori ON sgtb_ddt.ddt_cliente_id = gv_rivenditori.riv_id " + _
	  " WHERE ddt_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 

label = GetValueList(conn, NULL, "SELECT cat_nome_it FROM sgtb_ddt_categorie WHERE cat_id = "& rs("ddt_categoria_id"))


dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione "&label&" - modifica "
if standalone then
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
	<input type="hidden" name="tfn_ddt_categoria_id" value="<%= rs("ddt_categoria_id") %>">
	<% if standalone then %><input type="hidden" name="reload" value="true"><% end if %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati <%=label%></td>
					<td align="right" style="font-size: 1px;">
						<% if standalone then %>
							<a class="button" href="javascript:window.close();" title="chiudi la finestra" <%= ACTIVE_STATUS %>>
								CHIUDI</a>
						<% else %>
							<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="<%=Lcase(label)%> precedente" <%= ACTIVE_STATUS %>>
								&lt;&lt; PRECEDENTE
							</a>
							&nbsp;
							<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="<%=Lcase(label)%> successivo" <%= ACTIVE_STATUS %>>
								SUCCESSIVO &gt;&gt;
							</a>
						<% end if %>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">DATI <%=Ucase(label)%></th></tr>
		<tr>
			<td class="label" style="width:18%;">numero:</td>
			<td class="content" colspan="3"><%=rs("ddt_numero")%></td>
			<input type="hidden" name="tfn_ddt_numero" value="<%=rs("ddt_numero")%>">
		</tr>
		<tr>
			<td class="label">data:</td>
			<td class="content" colspan="3">
				<%= DateIta(rs("ddt_data"))%>
			</td>
			<input type="hidden" name="tfd_ddt_data" value="<%=rs("ddt_data")%>">
		</tr>
		<tr>
			<td class="label">cliente:</td>
			<td class="content" colspan="3"><%= ContactFullName(rs)%></td>
			<input type="hidden" name="tfn_ddt_cliente_id" value="<%=rs("ddt_cliente_id")%>">
		</tr>
		<tr>
			<td class="label">destinazione:</td>
			<td class="content" colspan="3">
				<table cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<% dim nome_destinazione
					if cIntero(rs("ddt_destinazione_id"))>0 then
						sql = "SELECT * FROM tb_indirizzario WHERE IDElencoIndirizzi = " & rs("ddt_destinazione_id")
						rsd.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
						nome_destinazione = ContactAddress(rsd)
						rsd.close 
					else
						nome_destinazione = ""
					end if					
					%>
					<td>
						<input type="hidden" name="tfn_ddt_destinazione_id" value="<%= rs("ddt_destinazione_id") %>">
						<input READONLY type="text" name="destinazione" style="padding-left:3px; width:100%" value="<%= nome_destinazione %>" 
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
					<% dim nome_trasportatore
					if cIntero(rs("ddt_trasportatore_id"))>0 then
						sql = "SELECT * FROM gv_rivenditori WHERE riv_id = " & rs("ddt_trasportatore_id")
						rsd.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
						nome_trasportatore = ContactFullName(rsd)
						rsd.close 
					else
						nome_trasportatore = ""
					end if					
					%>
					<td>
						<input type="hidden" name="tfn_ddt_trasportatore_id" value="<%= rs("ddt_trasportatore_id") %>">
						<input READONLY type="text" name="trasportatore" style="padding-left:3px; width:100%" value="<%= nome_trasportatore %>" 
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
				<input type="text" class="text" name="tft_ddt_peso" value="<%=rs("ddt_peso")%>" maxlength="255" size="22">
			</td>
		</tr>
		<tr>
			<td class="label">volume:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_ddt_volume" value="<%=rs("ddt_volume")%>" maxlength="255" size="22">
			</td>
		</tr>
		<tr>
			<td class="label">numero colli:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_ddt_numero_colli" value="<%=rs("ddt_numero_colli")%>" maxlength="255" size="22">
			</td>
		</tr>
		<script language="JavaScript" type="text/javascript">
			function rimuovi(idScheda){
				var campo_hidden = document.getElementById('scheda_' + idScheda)
				var button = document.getElementById('button_' + idScheda)
				campo_hidden.value = '';
				button.disabled = 'disabled';
				button.className = 'button_L2_disabled';
			}
		</script>
		<% if cIntero(rs("ddt_cliente_id"))>0 then %>
			<% sql = " SELECT * FROM (sgtb_schede INNER JOIN gv_articoli ON sgtb_schede.sc_modello_id = gv_articoli.rel_id) " & _
					 " LEFT JOIN sgtb_ddt ON sgtb_schede.sc_documento_ritiro_id = sgtb_ddt.ddt_id " & _
					 " INNER JOIN sgtb_stati_schede ON sgtb_schede.sc_stato_id = sgtb_stati_schede.sts_id " & _
					 " WHERE sc_documento_ritiro_id="&cIntero(request("ID")) & _
					 " ORDER BY sc_numero "
			rsd.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
				<tr><th colspan="5">SCHEDE ASSOCIATE</th></tr>
				<% if rsd.eof then %>
					<tr><td colspan="4" class="note">Nessuna scheda da ritirare per il cliente selezionato</td></tr>
				<% else %>
					<tr>
						<th class="l2_center" style="width:8%">&nbsp;</th>
						<th class="l2_center" style="width:18%">numero scheda e data</th>
						<th class="l2_center" style="width:18%">stato</th>
						<th class="l2_center" style="width:14%">costo ritiro</th>
						<th class="L2">modello</th>
					</tr>
				<% end if %>
				<% while not rsd.eof %>
					<tr>
						<td class="content_center">
							<input type="hidden" id="scheda_<%=rsd("sc_id")%>" name="id_schede" value="<%=rsd("sc_id")%>">
							<% if cIntero(rsd.recordcount)>1 then %>
								<a class="button_L2" href="javascript:void(0)" id="button_<%=rsd("sc_id")%>" onclick="rimuovi(<%=rsd("sc_id")%>);" 
									title="Rimuovi associazione con questa scheda" <%= ACTIVE_STATUS %>>
									rimuovi
								</a>
							<% else %>
								<a class="button_L2_disabled" disabled href="javascript:void(0)" title="Impossibile rimuovere associazione con questa scheda">
									rimuovi
								</a>
							<% end if %>
						</td>
						<td class="content_center"><% CALL SchedaLink(rsd("sc_id"), rsd("sc_numero") & " del " & rsd("sc_data_ricevimento"))%></td>
						<td class="content"><%=rsd("sts_nome_it")%></td>
						<td class="content_center">
							<input type="text" class="number" name="costo_ritiro_scheda_<%=rsd("sc_id")%>" value="<%= FormatPrice(cReal(rsd("sc_costo_presa")), 2, false) %>" size="7"> &euro;
						</td>
						<td class="content">
							<% CALL ArticoloLink(rsd("art_id"), rsd("art_nome_it"), rsd("art_cod_int")) %>
							<% if rsd("art_varianti") then %>
								<%= ListValoriVarianti(conn, rsi, rsd("rel_id")) %>
							<% else %>
								&nbsp;
							<% end if %>
						</td>
					</tr>
				<% rsd.moveNext %>
			<% wend %>
			</table>
			<% rsd.close %>
		<% end if %>
		
		<% sql = " SELECT * FROM (sgtb_schede INNER JOIN gv_articoli ON sgtb_schede.sc_modello_id = gv_articoli.rel_id) " & _
				 " LEFT JOIN sgtb_ddt ON sgtb_schede.sc_documento_ritiro_id = sgtb_ddt.ddt_id " & _
				 " INNER JOIN sgtb_stati_schede ON sgtb_schede.sc_stato_id = sgtb_stati_schede.sts_id " & _
				 " WHERE ISNULL(sts_elenco_ddt_da_ritirare,0)=1 AND ISNULL(sc_documento_ritiro_id, 0)=0 " & _
				 " AND sc_cliente_id = " & cIntero(rs("ddt_cliente_id"))
		rsd.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<tr><th colspan="5">SCHEDE NON ASSOCIATE</th></tr>
			<% if rsd.eof then %>
				<tr><td colspan="4" class="note">Nessuna scheda da ritirare per il cliente selezionato</td></tr>
			<% else %>
				<tr>
					<th class="l2_center" style="width:8%">associa</th>
					<th class="l2_center" style="width:18%">numero scheda e data</th>
						<th class="l2_center" style="width:18%">stato</th>
					<th class="l2_center" style="width:14%">costo ritiro</th>
					<th class="L2">modello</th>
				</tr>
			<% end if %>
			<% while not rsd.eof %>
				<tr>
					<td class="content_center">
						<input type="checkbox" class="noBorder <%=IIF(cIntero(rsd("sc_documento_ritiro_id"))=cIntero(request("ID")),"checked","")%>" name="id_schede" value="<%=rsd("sc_id")%>" 
								<%= chk(cIntero(rsd("sc_documento_ritiro_id"))=cIntero(request("ID")))%>>
					</td>
					<td class="content_center"><% CALL SchedaLink(rsd("sc_id"), rsd("sc_numero") & " del " & rsd("sc_data_ricevimento"))%></td>
					<td class="content"><%=rsd("sts_nome_it")%></td>
					<td class="content_center">
						<input type="text" class="number" name="costo_ritiro_scheda_<%=rsd("sc_id")%>" value="<%= FormatPrice(cReal(rsd("sc_costo_presa")), 2, false) %>" size="7"> &euro;
					</td>
					<td class="content">
						<% CALL ArticoloLink(rsd("art_id"), rsd("art_nome_it"), rsd("art_cod_int")) %>
						<% if rsd("art_varianti") then %>
							<%= ListValoriVarianti(conn, rsi, rsd("rel_id")) %>
						<% else %>
							&nbsp;
						<% end if %>
					</td>
				</tr>
			<% rsd.moveNext %>
		<% wend %>
		</table>

		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<tr><th colspan="4">NOTE</th></tr>
			<tr>
				<td class="content" colspan="4">
					<textarea style="width:100%;" rows="3" name="tft_ddt_note"><%= request("tft_ddt_note") %></textarea>
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="4">
					(*) Campi obbligatori.
					<input type="submit" class="button" name="salva" value="SALVA">
					<% if not standalone then %>
						<input type="submit" class="button" name="salva_elenco" value="SALVA & TORNA A ELENCO">
					<% end if %>
				</td>
			</tr>
		</table>
		&nbsp;
	</form>
</div>

<% if Session("INFOSCHEDE_ADMIN")<>"" then %>
	<div id="pulsanti" style="position:absolute; top:91px; left:760px; width:180px;">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Operazioni</caption>
			<tr><th>RITIRI</th></tr>
			<tr>
				<td class="content_center" style="font-size:1px;">
					&nbsp;<br>
					<%
					dim IDCNT, KEY

					sql = "SELECT IDElencoIndirizzi FROM gv_rivenditori WHERE riv_id = " & rs("ddt_cliente_id")
					IDCNT = GetValueList(conn, NULL, sql)
					sql = "SELECT codiceInserimento FROM gv_rivenditori WHERE riv_id = " & rs("ddt_cliente_id")
					KEY = GetValueList(conn, NULL, sql)
					%>
					<a class="button_L2" style="width:160px;text-align:center;" href="<%= GetPageSiteUrl(conn, Session("INFOSCHEDE_ID_PAG_RICH_RITIRO"), "it") & "&ID_ADMIN="&Session("ID_ADMIN")&"&DDTID="&cIntero(rs("ddt_id"))&"&CLIENTEID="&rs("ddt_cliente_id")&"&IDCNT="&IDCNT&"&KEY="&KEY %>" target="richiesta_ritiro_<%=rs("ddt_id")%>"
					onclick="OpenAutoPositionedScrollWindow('', 'richiesta_ritiro_<%=rs("ddt_id")%>', 800, 800, true)" title="Click per aprire la finestra per visualizzare il documento di ritiro" <%= ACTIVE_STATUS %>>
						VISUALIZZA RICHIESTA DI RITIRO
					</a>
					<br>&nbsp;<br>
					<a class="button_L2" style="width:160px;text-align:center;" href="<%= GetPageSiteUrl(conn, Session("INFOSCHEDE_ID_PAG_INVIO_RICH_RIT"), "it") & "&ID_ADMIN="&Session("ID_ADMIN")&"&ID_DDT="&cIntero(rs("ddt_id"))%>" target="richiesta_ritiro_<%=rs("ddt_id")%>_invia"
					onclick="OpenAutoPositionedScrollWindow('', 'richiesta_ritiro_<%=rs("ddt_id")%>_invia', 520, 300, true)" title="Click per inviare il documento di ritiro" <%= ACTIVE_STATUS %>>
						INVIA RICHIESTA DI RITIRO
					</a>
					<br>&nbsp;
				</td>
			</tr>
		</table>
	</div>
<% end if %>


</body>
</html>

<% if standalone then %>
	<script language="JavaScript" type="text/javascript">
		FitWindowSize(this);
	</script>
<% end if %>

<%
rs.close
rsd.close
set rs = nothing
set rsd = nothing
set rsi = nothing
conn.Close
set conn = nothing
%>