<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim conn, rs, rsp, rsv, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.Recordset")
set rsp = Server.CreateObject("ADODB.Recordset")
set rsv = Server.CreateObject("ADODB.Recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_ARTICOLI_SQL"), "art_id", "ArticoliGiacenze.asp")
end if

sql = " SELECT * FROM gtb_articoli INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " + _
	  " LEFT JOIN gtb_scontiq_classi ON gtb_articoli.art_scontoQ_id = gtb_scontiq_classi.scc_id " + _
	  " WHERE art_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

dim dicitura, tipo, listino
if rs("art_se_bundle") then
	tipo = "bundle"
elseif rs("art_se_confezione") then
	tipo = "confezione"
elseif rs("art_varianti") then
	tipo ="articolo con varianti"
else
	tipo ="articolo singolo"
end if
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione articoli - giacenze " & tipo
dicitura.puls_new = "INDIETRO;SCHEDA ARTICOLO;"
dicitura.link_new = "Articoli.asp;ArticoliMod.asp?ID=" & request("ID")
dicitura.puls_2a_riga.Add "PREZZI","ArticoliPrezzi.asp?ID=" & request("ID")
if Session("ATTIVA_FAQ_ARTICOLI") then
	dicitura.puls_2a_riga.Add "FAQ","ArticoliFaq.asp?ID=" & request("ID")
end if
if Session("ATTIVA_COMMENTI") then
	dicitura.puls_2a_riga.Add "COMMENTI","ArticoliCommenti.asp?ID=" & request("ID")
end if
dicitura.scrivi_con_sottosez()
%>


<div id="content_abbassato">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati <%= tipo %> con codice &quot;<%= rs("art_cod_int") %>&quot;</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="articolo precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="articolo successiva" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="7">DATI DELL'ARTICOLO</th></tr>
		<% CALL ArticoloScheda (conn, rs, rsp) %>
		<% if not rs("art_varianti") then%>
			<tr>
				<td class="label" colspan="2">giacenza minima per ogni magazzino</td>
				<td class="content" colspan="5"><%= rs("art_giacenza_min") %></td>
			</tr>
			<tr>
				<td class="label" colspan="2">quantit&agrave; minima ordinabile</td>
				<td class="content" colspan="5"><%= rs("art_qta_min_ord") %></td>
			</tr>
			<tr>
				<td class="label" colspan="2">loto di riordino</td>
				<td class="content" colspan="5"><%= rs("art_lotto_riordino") %></td>
			</tr>
			<% if rs("art_se_bundle") then %>
				<tr><th colspan="7">COMPONENTI DEL BUNDLE</th></tr>
				<tr>
					<td colspan="7">
						<% sql = " SELECT * FROM gv_articoli INNER JOIN gtb_bundle ON gv_articoli.rel_id = gtb_bundle.bun_articolo_id " + _
								 " WHERE gtb_bundle.bun_bundle_id IN (SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & cIntero(request("ID")) & ") "
						rsp.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch %>
						<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
							<tr>
								<td class="label" colspan="4" style="width:74%">
									<% if rsp.eof then %>
										Nessun componente definito.
									<% else %>
										Trovati n&ordm; <%= rsp.recordcount %> record
									<% end if %>
								</td>
							</tr>
							<% if not rsp.eof then %>
								<tr>
									<th class="l2_center" width="14%">codice</th>
									<th class="L2">descrizione componente</th>
									<th class="l2_center" width="8%">quantit&agrave;</th>
									<th class="l2_center" width="7%" >operazioni</th>
								</tr>
								<% while not rsp.eof %>
									<tr>
										<td class="content"><%= rsp("rel_cod_int")%></td>
										<td class="content">
											<% ArticoloLink rsp("art_id"), rsp("art_nome_it"), rsp("rel_cod_int") %>
											<% if rsp("art_varianti") then %>
												<%= ListValoriVarianti(conn, rsv, rsp("rel_id")) %>
											<% end if %>
										</td>
										<td class="content_center"><%= rsp("bun_quantita")%></td>
										<td class="content_center">
											<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la gestione delle giacenze del componente" <%= ACTIVE_STATUS %>
											   onclick="OpenAutoPositionedScrollWindow('ArticoliGiacenze.asp?ID=<%= rsp("art_id") %>', 'COMArt', 760, 400, true)">
												GIACENZE
											</a>
										</td>
									</tr>
									<%rsp.movenext
								wend 
							end if%>
						</table>
						<% rsp.close %>
					</td>
				</tr>
			<% end if
		else %>
			<tr><th colspan="7">VARIANTI DELL'ARTICOLO</th></tr>
			<tr>
				<td colspan="7">
					<%sql = " SELECT * FROM grel_art_valori WHERE rel_art_id=" & request("ID") & " ORDER BY rel_ordine "
					rsp.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch%>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr>
							<td class="label" style="width:100%" colspan="8">
								<% if rsp.eof then %>
									Nessuna variante definita per il prodotto
								<% else %>
									Trovati n&ordm; <%= rsp.recordcount %> record
								<% end if %>
							</td>
						</tr>
						<tr>
							<th class="L2">codice</th>
							<th class="L2">variante</th>
							<th class="l2_center">a catalogo</th>
							<th class="l2_center">giac. min.</th>
							<th class="l2_center">qta min. ord.</th>
							<th class="l2_center">qta lotto</th>
						</tr>
						<% while not rsp.eof %>
							<tr>
								<td class="content"><%= rsp("rel_cod_int") %></td>
								<td class="content">
									<% CALL TableValoriVarianti(conn, rsv, rsp("rel_id"), IIF(rsp("rel_disabilitato"), "content_disabled", "content")) %>
								</td>
								<td class="content_center"><input type="checkbox" class="checkbox" disabled <%= chk(not rsp("rel_disabilitato")) %>></td>
								<td class="content_center"><%= rsp("rel_giacenza_min") %></td>
								<td class="content_center"><%= rsp("rel_qta_min_ord") %></td>
								<td class="content_center"><%= rsp("rel_lotto_riordino") %></td>
							</tr>
							<% rsp.movenext
						wend %>
					</table>
					<% rsp.close %>
				</td>
			</tr>
		<% end if %>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>Giacenze dell'articolo</caption>
		<tr><th colspan="3">QUANTIT&Agrave; GLOBALE</th></tr>
		<% sql = " SELECT (SUM(gia_qta)) AS GIACENZA, (SUM(gia_impegnato)) AS IMPEGNATO, (SUM(gia_ordinato)) AS ORDINATO, (COUNT(*)) AS MAGAZZINI " + _
				 " FROM gtb_articoli INNER JOIN grel_Art_valori ON gtb_articoli.art_id = grel_Art_valori.rel_art_id " + _
				 " INNER JOIN grel_giacenze ON grel_art_valori.rel_id = grel_giacenze.gia_art_var_id " + _
				 " WHERE gtb_articoli.art_id=" & cIntero(request("ID"))
		rsp.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText		%>
		<tr>
			<td class="label" style="width:20%;">giacenza nei magazzini</td>
			<% if rs("art_varianti") then %>
				<td class="content"><%= rsp("GIACENZA") %></td>
			<% elseif rsp("GIACENZA")<1 then	'esaurita
			%>
				<td class="content alert" title="esaurito in tutti i magazzini"><%= rsp("GIACENZA") %></td>
			<%elseif rsp("GIACENZA") <= (cInteger(rs("art_giacenza_min")) * rsp("MAGAZZINI")) then		'in esaurimento
			%>
				<td class="content warning" title="in esaurimento o esaurito in almeno un magazzino"><%= rsp("GIACENZA") %></td>
			<% else %>
				<td class="content ok"><%= rsp("GIACENZA") %></td>
			<% end if %>
		</tr>
		<tr>
			<td class="label">merce impegnata da ordini</td>
			<td class="content" colspan="2"><%= rsp("IMPEGNATO") %></td>
		</tr>
		<% if not rs("art_se_bundle") then%>
			<tr>
				<td class="label">merce ordinata a fornitore</td>
				<td class="content" colspan="2"><%= rsp("ORDINATO") %></td>
			</tr>
		<% end if 
		rsp.close
		if rs("art_varianti") then %>
			<tr><th colspan="7">QUANTIT&Agrave; GLOBALI PER OGNI VARIANTE</th></tr>
			<tr>
				<td colspan="7">
					<%sql = " SELECT *, (SELECT SUM(gia_qta) FROM grel_giacenze WHERE gia_art_var_id = grel_art_valori.rel_id) AS GIACENZA, " + _
							" (SELECT SUM(gia_impegnato) FROM grel_giacenze WHERE gia_art_var_id = grel_art_valori.rel_id) AS IMPEGNATO, " + _
							" (SELECT SUM(gia_ordinato) FROM grel_giacenze WHERE gia_art_var_id = grel_art_valori.rel_id) AS ORDINATO, " + _
							" (SELECT COUNT(*) FROM grel_giacenze WHERE gia_art_var_id = grel_art_valori.rel_id) AS MAGAZZINI " + _
							" FROM grel_art_valori WHERE rel_art_id=" & cIntero(request("ID")) & " ORDER BY rel_ordine "
					rsp.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch%>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr>
							<td class="label" style="width:100%" colspan="8">
								<% if rsp.eof then %>
									Nessuna variante definita per il prodotto
								<% else %>
									Trovati n&ordm; <%= rsp.recordcount %> record
								<% end if %>
							</td>
						</tr>
						<tr>
							<th class="L2">codice</th>
							<th class="L2">variante</th>
							<th class="l2_center">a catalogo</th>
							<th class="l2_center">giacenza</th>
							<th class="l2_center">qta. impegnata</th>
							<th class="l2_center">qta. ordinata f.</th>
						</tr>
						<% while not rsp.eof %>
							<tr>
								<td class="content"><%= rsp("rel_cod_int") %></td>
								<td class="content">
									<% CALL TableValoriVarianti(conn, rsv, rsp("rel_id"), IIF(rsp("rel_disabilitato"), "content_disabled", "content")) %>
								</td>
								<td class="content_center"><input type="checkbox" class="checkbox" disabled <%= chk(not rsp("rel_disabilitato")) %>></td>
								<% if rsp("GIACENZA")<1 then	'esaurita
								%>
									<td style="text-align:center;" class="content alert" title="esaurito in tutti i magazzini"><%= rsp("GIACENZA") %></td>
								<%elseif rsp("GIACENZA") <= (cInteger(rsp("rel_giacenza_min")) * rsp("MAGAZZINI")) then		'in esaurimento
								%>
									<td style="text-align:center;" class="content warning" title="in esaurimento o esaurito in almeno un magazzino"><%= rsp("GIACENZA") %></td>
								<% else %>
									<td style="text-align:center;" class="content ok"><%= rsp("GIACENZA") %></td>
								<% end if %>
								<td class="content_center"><%= rsp("IMPEGNATO") %></td>
								<td class="content_center"><%= rsp("ORDINATO") %></td>
							</tr>
							<% rsp.movenext
						wend 
						rsp.close%>
					</table>
				</td>
			</tr>
		<% end if %>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<%dim magazzino
		sql = " SELECT * FROM grel_art_valori INNER JOIN grel_giacenze ON grel_art_valori.rel_id = grel_giacenze.gia_art_var_id " + _
			  " INNER JOIN gtb_magazzini ON grel_giacenze.gia_magazzino_id = gtb_magazzini.mag_id " + _
			  " WHERE rel_art_id=" & cIntero(request("ID")) & " ORDER BY mag_nome, mag_id, rel_ordine "
		rsp.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch
		magazzino = 0
		while not rsp.eof
			if magazzino <> rsp("mag_id") then %>
				<tr><th colspan="7"><%= rsp("mag_nome") %></th></tr>
				<% if rs("art_varianti") then %>
					<form action="MagazziniInventario.asp?ID=<%= rsp("mag_id") %>" method="post" target="magazzino" id="form_<%= rsp("mag_id") %>" name="form_<%= rsp("mag_id") %>">
					<tr>
						<td class="content_right" colspan="7">
							<input type="hidden" name="search_nome" value="<%= TextEncode(rs("art_nome_it")) %>">
							<input type="hidden" name="search_categoria" value="<%= rs("art_tipologia_id") %>">
							<input type="hidden" name="search_marchio" value="<%= rs("art_marca_id") %>">
							<input type="submit" name="cerca" value="APRI MAGAZZINO" class="button_L2" title="Apre l'inventario di magazzino relativo" onclick="OpenAutoPositionedScrollWindow('', 'magazzino', 760, 450, true);">
						</td>
					</tr>
					</form>
				<% end if %>
				<tr>
					<% if rs("art_varianti") then %>
						<th class="L2">codice</th>
						<th class="L2">variante</th>
					<% else %>
						<th class="L2" colspan="2">codice</th>
					<% end if %>
					<th class="l2_center">giacenza</th>
					<th class="l2_center">impegnato</th>
					<th class="l2_center">ordinato a fornitore</th>
					<th class="l2_center">data arrivo</th>
					<th class="l2_center" style="width:8%;">modifica</th>
				</tr>
				<% magazzino = rsp("mag_id")
			end if %>
			<tr>
				<% if rs("art_varianti") then %>
					<td class="content"><%= rsp("rel_cod_int") %></td>
					<td class="<%= IIF(rsp("rel_disabilitato"), "content_disabled", "content") %>">
						<% CALL TableValoriVarianti(conn, rsv, rsp("rel_id"), IIF(rsp("rel_disabilitato"), "content_disabled", "content")) %>
					</td>
				<% else %>
					<td class="content" colspan="2"><%= rsp("rel_cod_int") %></td>
				<% end if %>
				<td class="content_center"><%= rsp("gia_qta") %></td>
				<td class="content_center"><%= rsp("gia_impegnato") %></td>
				<td class="content_center"><%= rsp("gia_ordinato") %></td>
				<td class="content_center"><%= rsp("gia_ordinato_data_arrivo") %></td>
				<td class="content_center">
					<% if rs("art_se_bundle") then %>
						<a class="button_L2_disabled" title="Le giacenze dei bundle non sono modificabili. Vengono calcolate automaticamente dal sistema sulla base delle giacenze dei prodotti componenti.">
							MODIFICA
						</a>
					<% else %>
						<a class="button_L2" href="javascript:void(0);" title="Apre la modifica delle giacenze dell'articolo" <%= ACTIVE_STATUS %>
				   		   onclick="OpenAutoPositionedScrollWindow('ArticoliGiacenze_Mod.asp?ID=<%= rsp("gia_id") %>', 'Giacenze', 510, 400, false)">
							MODIFICA
						</a>
					<% end if %>
				</td>
			</tr>
			<% rsp.movenext
		wend
		rsp.close %>
	</table>
	&nbsp;
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