<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if (request("salva")<>"" OR request("salva_elenco")<>"") AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ArticoliSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../nextB2B/Tools4Save_B2B.asp" -->

<% 	
dim conn, rs, rsv, rsp, sql, i, aux, txt, rs_spe
dim categoria

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.Recordset")
set rsv = Server.CreateObject("ADODB.Recordset")
set rsp = Server.CreateObject("ADODB.Recordset")
set rs_spe = Server.CreateObject("ADODB.Recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_ARTICOLI_SQL"), "art_id", "ArticoliMod.asp")
end if

sql = " SELECT * FROM (gtb_articoli INNER JOIN gtb_iva ON gtb_Articoli.art_iva_id = gtb_iva.iva_id) " + _
	  " LEFT JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_Valori.rel_art_id " + _
	  " WHERE art_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText



if not cBoolean(rs("art_varianti"),false) then
	rs.close
	sql = " SELECT *, (SELECT COUNT(*) FROM gtb_dettagli_ord WHERE det_art_var_id = gv_articoli.rel_id) AS N_ORDINI " + _
		  " FROM gv_articoli WHERE art_id=" & cIntero(request("ID"))
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
end if


dim dicitura

set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione modelli - modifica "
dicitura.puls_new = "INDIETRO;"
dicitura.link_new = "Articoli.asp;"

'CALL dicitura.InitializeIndex(Index, "gtb_articoli", request("ID"))
dicitura.scrivi_con_sottosez()

%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_art_varianti" value="<%= IIF(rs("art_varianti"), "1", "0") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati modello con codice &quot;<%= rs("art_cod_int") %>&quot;</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" name="<%= Server.HTMLEncode(rs("art_cod_int")) %>" href="?ID=<%= cIntero(request("ID")) %>&goto=PREVIOUS" title="articolo precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= cIntero(request("ID")) %>&goto=NEXT" title="articolo successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="7">DATI PRINCIPALI</th></tr>
		<% if not rs("art_varianti") then 
			sql = " SELECT * FROM gtb_lista_codici LEFT JOIN gtb_codici ON " + _
				  " ( gtb_lista_codici.lstCod_id = gtb_codici.Cod_lista_id AND gtb_codici.cod_variante_id=" & rs("rel_id") & " )" + _
				  " WHERE lstCod_sistema=1 ORDER BY lstCod_nome" 
			rsv.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
			<tr>
				<td class="label" style="width:16%;"<% if rsv.recordcount>0 then %>rowspan="<%= 1+rsv.recordcount %>"<% end if %>>codici:</td>
				<td class="label" style="width:8%;">interno:</td>
				<td class="content" style="width:25%;">
					<input type="text" class="text" name="tft_art_cod_int" value="<%= rs("art_cod_int") %>" maxlength="50" size="15">
					(*)
				</td>
				<td class="label" style="width:8%;">alternativo:</td>
				<td class="content" colspan="3">
					<input type="text" class="text" name="tft_art_cod_alt" value="<%= rs("art_cod_alt") %>" maxlength="50" size="15">
				</td>
			</tr>
			<% while not rsv.eof %>
				<tr>
					<td class="label_no_width" colspan="2">
						<%= rsv("lstCod_nome") %>
					</td>
					<td class="content" colspan="4">
						<input type="text" class="text" name="codice_articolo_<%= rsv("lstCod_id") %>" value="<%= rsv("cod_codice") %>" maxlength="50" size="23">
					</td>
				</tr>
				<% rsv.movenext
			wend 
			rsv.close
		end if
		for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% 	if i = 0 then %>
					<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
				<% 	end if %>
				<td class="content" colspan="6">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_art_nome_<%= Application("LINGUE")(i) %>" value="<%= textEncode(rs("art_nome_"& Application("LINGUE")(i))) %>" maxlength="250" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next 
		
		categoria = IIF(cInteger(request("tfn_art_tipologia_id"))>0, request("tfn_art_tipologia_id"), rs("art_tipologia_id"))%>
		<tr>
			<td class="label">categoria:</td>
			<td class="content" colspan="6">
				<%CALL dropDown(conn, cat_modelli.QueryElenco(true, ""), "tip_id", "NAME", "tfn_art_tipologia_id", categoria, true, " onchange=""form1.submit()""", LINGUA_ITALIANO)%>
				(*)
			</td>
		</tr>
		<% sql = "SELECT COUNT(*) FROM gtb_tipologie_raggruppamenti"
		if cIntero(getValueList(conn, rsv, sql))>0 then %>
			<tr>
				<td class="label">&nbsp;</td>
				<td class="label" colspan="2">raggruppamento di pubblicazione:</td>
				<td class="content" colspan="5">
					<% if cInteger(categoria)>0 then
						sql = " SELECT * FROM gtb_tipologie_raggruppamenti WHERE rag_tipologia_id=" & categoria
						rsv.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
						if rsv.eof then %>
							<span class="note">Nessun raggruppamento disponibile per questa categoria di prodotti</span>
							<input type="hidden" name="nfn_art_raggruppamento_id" value="NULL">
						<% else
							CALL DropDownRecordset(rsv, "rag_id", "rag_nome_it", "nfn_art_raggruppamento_id", rs("art_raggruppamento_id"), false, "", LINGUA_ITALIANO)
						end if
						rsv.close
					end if %>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label" nowrap>marchio / produttore:</td>
			<td class="content" colspan="6">
				<%CALL dropDown(conn, "SELECT mar_id, mar_nome_it FROM gtb_marche ORDER BY mar_nome_it", _
							    "mar_id", "mar_nome_it", "tfn_art_marca_id", rs("art_marca_id"), true, "", LINGUA_ITALIANO)%>
			</td>
		</tr>
		<tr><th colspan="7">DATI PER LA GESTIONE</th></tr>
		<tr>
			<td class="label" colspan="2">non a catalogo:</td>
			<td class="content" colspan="5"><input type="checkbox" class="checkbox" name="chk_art_disabilitato" <%= chk(rs("art_disabilitato")) %>></td>
		</tr>
		
		<tr><th colspan="7">DESCRIZIONE</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="7">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="5" name="tft_art_descr_<%= Application("LINGUE")(i) %>"><%= rs("art_descr_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		<%next %>
	</table>
	
	<% sql = " SELECT ct_id FROM gtb_carattech INNER JOIN gtb_tip_ctech ON gtb_carattech.ct_id = gtb_tip_ctech.rct_ctech_id " & _
			 " WHERE rct_tipologia_id = " & IIF(cInteger(request("tfn_art_tipologia_id"))>0, cInteger(request("tfn_art_tipologia_id")), rs("art_tipologia_id"))
		if cIntero(GetValueList(conn, NULL, sql)) > 0 then %>	
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
				<tr><th colspan="7">CARATTERISTICHE TECNICHE</th></tr>
				<% 	sql = " SELECT *" + _
						  " FROM gtb_carattech"& _
						  " INNER JOIN gtb_tip_ctech ON (gtb_carattech.ct_id = gtb_tip_ctech.rct_ctech_id AND rct_tipologia_id=" & IIF(cInteger(request("tfn_art_tipologia_id"))>0, cInteger(request("tfn_art_tipologia_id")), rs("art_tipologia_id")) &")" + _
						  " LEFT JOIN grel_art_ctech ON (gtb_carattech.ct_id = grel_art_ctech.rel_ctech_id AND grel_art_ctech.rel_art_id=" & rs("art_id") &")"& _
						  " LEFT JOIN gtb_carattech_raggruppamenti ON gtb_carattech.ct_raggruppamento_id = gtb_carattech_raggruppamenti.ctr_id " & _
						  " ORDER BY ctr_ordine, ctr_id, rct_ordine"
				CALL DesForm  (conn, sql, "gtb_carattech", "ct_id", "ct_nome_it", "ct_tipo", "ct_unita_it", "", "rel_ctech_", "rel_ctech_", "ctr_titolo_it", cIntero(request("ID")) = 0, 7)
				%>
			</table>
	<%	end if %>
	
	<% 	CALL oArticoliFoto.Elenco(cIntero(request("ID")), "FOTO") %>
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<tr><th colspan="7">NOTE INTERNE</th></tr>
		<tr>
			<td class="content" colspan="7">
				<textarea style="width:100%;" rows="3" name="tft_art_note"><%= rs("art_note") %></textarea>
			</td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<% CALL Form_DatiModifica(conn, rs, "art_") %>
		<tr>
			<td class="footer" colspan="7">
				(*) Campi obbligatori.
				<input type="submit" style="width:23%;" class="button" name="salva_elenco" value="SALVA & TORNA ALL'ELENCO">
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<%
set rs = nothing
set rsv = nothing
set rsp = nothing
conn.Close
set conn = nothing
%>