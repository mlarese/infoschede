<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" 

dim conn, rs, sql,rsc,nomeC
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	Session("B2B_SQL_GIACENZE") = ""
	CALL GotoRecord(conn, rs, Session("B2B_MAG_SQL"), "mag_id", "MagazziniInventario.asp")
end if

response.buffer = false
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione magazzino - inventario"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Magazzini.asp"
dicitura.scrivi_con_sottosez()  

'imposta ricerca
dim sql_where
if Request.ServerVariables("REQUEST_METHOD")="POST" and (request("cerca")<>"" or request("tutti")<>"") then
	session("B2B_INVENTARIO_SEARCHED") = true
	if request("tutti")<>"" then
		CALL SearchSession_Reset("mag_")
		session("B2B_INVENTARIO_SEARCHED_ALL") = true
	else
		session("B2B_INVENTARIO_SEARCHED_ALL") = false
		CALL SearchSession_Reset("mag_")
		CALL SearchSession_Set("mag_")
	end if
	sql_where = ""
	'filtra per codice interno
	if Session("mag_codice_int")<>"" then
		if sql_where <>"" then sql_where = sql_where & " AND "
        sql_where = sql_where & SQL_FullTextSearch(Session("mag_codice_int"), "rel_cod_int")
	end if
	'filtra per codice produttore
	if Session("mag_codice_pro")<>"" then
		if sql_where <>"" then sql_where = sql_where & " AND "
        sql_where = sql_where & SQL_FullTextSearch(Session("mag_codice_pro"), "rel_cod_pro")
	end if
	'filtra per nome
	if Session("mag_nome")<>"" then
		if sql_where <>"" then sql_where = sql_where & " AND "
		sql_where = sql_where & sql_FullTextSearch(Session("mag_nome"), FieldLanguageList("art_nome_"))
	end if
	'filtra per categoria
	if Session("mag_categoria")<>"" then
		if sql_where <>"" then sql_where = sql_where & " AND "
		sql_where = sql_where & " art_tipologia_id IN (" & categorie.FoglieID(Session("mag_categoria")) & " ) "
	end if
	'filtra per marca
	if Session("mag_marchio")<>"" then
		if sql_where <>"" then sql_where = sql_where & " AND "
		sql_where = sql_where & " art_marca_id=" & Session("mag_marchio")
	end if
	
	'ricerca per stato a catalogo
	if Session("mag_stato_catalogo")<>"" then
		if not (instr(1, Session("mag_stato_catalogo"), "1", vbTextCompare)>0 AND _
			    instr(1, Session("mag_stato_catalogo"), "0", vbTextCompare)>0 ) then
			if sql_where <>"" then sql_where = sql_where & " AND "
			if instr(1, Session("mag_stato_catalogo"), "1", vbTextCompare)>0 then
				'articolo a catalogo
				sql_where = sql_where & " NOT (" & SQL_IsTrue(conn, "art_disabilitato") & ") "
			elseif instr(1, Session("mag_stato_catalogo"), "0", vbTextCompare)>0 then
				'articolo fuori catalogo
				sql_where = sql_where & SQL_IsTrue(conn, "art_disabilitato")
			end if
		end if
	end if
	
	'ricerca full-text
	if Session("mag_full_text")<>"" then
		if sql <>"" then sql = sql & " AND "
		sql_where = sql_where & SQL_FullTextSearch(Session("mag_full_text"), FieldLanguageList("art_nome_;art_descr_"))
	end if
	
	if sql_where <>"" then sql_where = " AND " & sql_where
	Session("B2B_SQL_GIACENZE") = " WHERE gia_magazzino_id = " & cIntero(request("ID")) & sql_where
	Session("B2B_WHERE_SQL_GIACENZE") = sql_where
end if

if Session("B2B_SQL_GIACENZE") = "" and Session("B2B_WHERE_SQL_GIACENZE")<>"" then
		Session("B2B_SQL_GIACENZE") = " WHERE gia_magazzino_id = " & cIntero(request("ID")) & Session("B2B_WHERE_SQL_GIACENZE")
elseif Session("B2B_SQL_GIACENZE") = "" then
		Session("B2B_SQL_GIACENZE") = " WHERE gia_magazzino_id = " & cIntero(request("ID"))
end if
%>

<div id="content_liquid" style="width:900px;">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="20%" valign="top">
				<% CALL Ricerca(conn, rs) %>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
				<form action="" method="post" id="form2" name="form2">
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>	
						<table align="right" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td align="right" style="font-size: 1px;">
									<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="listino precedente" <%= ACTIVE_STATUS %>>
										&lt;&lt; PRECEDENTE
									</a>
									&nbsp;
									<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="listino successivo" <%= ACTIVE_STATUS %>>
										SUCCESSIVO &gt;&gt;
									</a>
								</td>
							</tr>
						</table>
						<% sql = "SELECT mag_nome FROM gtb_magazzini WHERE mag_id=" & cIntero(request("ID")) %>
						Gestione inventario del magazzino "<%= GetValueList(conn, rs, sql) %>"
					</caption>
					<% if session("B2B_INVENTARIO_SEARCHED") then
						sql = " SELECT * " + _
							  " FROM grel_giacenze INNER JOIN grel_art_valori ON grel_giacenze.gia_art_var_id = grel_art_valori.rel_id " + _
							  "      INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + _
							  Session("B2B_SQL_GIACENZE")
						rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
						if rs.eof then	%>
							<tr>
								<td class="noRecords">Nessun articolo trovato</td>
							</tr>
						<% else %>
							<tr><th colspan="8">Elenco articoli del magazzino
							</th></tr>
							<tr>
								<td class="label" colspan="8">
									Selezionati n&ordm; <%= rs.recordcount %> articoli
								</td>
							</tr>
							<form action="" method="post" id="form1" name="form1">
							<tr>
								<th class="L2" colspan="2" style="border-bottom:0px;">articolo</th>
								<th class="l2_center" colspan="2" style="border-bottom:0px;">giacenza</th>
								<th class="l2_center" colspan="3" style="border-bottom:0px;">situazione</th>
								<th class="l2_center" rowspan="2">operazioni</th>
							</tr>
							<tr>
								<th class="L2">codice</th>
								<th class="L2" style="width:30%;">nome</th>
								<th class="L2_right" nowrap>a magazzino</th>
								<th class="L2_right" nowrap>all'ordine</th>
								<th class="L2_right" nowrap>impegnato</th>
								<th class="L2_right" nowrap>ordinato</th>
								<th class="L2_right" nowrap>data arrivo</th>
							</tr>
							<% while not rs.eof %>
								<tr>
									<td class="content"><%= rs("rel_cod_int") %></td>
									<td class="content">
										<% CALL ArticoloLink(rs("rel_art_id"), rs("art_nome_it"), rs("rel_cod_int"))
										if rs("art_varianti") then %>
											<%= ListValoriVarianti(conn, rsc, rs("rel_id")) %>
										<% end if %>
									</td>
									<% if rs("gia_qta") <= 0 then %>	
										<td class="content alert" style="text-align:right;">
									<% elseif rs("gia_qta") <= rs("rel_giacenza_min") then%>
										<td class="content warning" style="text-align:right;">
									<% else %>
										<td class="content ok" style="text-align:right;">
									<% end if %>
										<%= rs("gia_qta") %></td>
									<td class="content_right"><%= rs("gia_qta") - rs("gia_impegnato") %></td>
									<td class="content_right"><%= rs("gia_impegnato") %></td>
									<td class="content_right"><%= rs("gia_ordinato") %></td>
									<td class="content_right"><%= DateIta(rs("gia_ordinato_data_arrivo")) %></td>
									<td class="Content_center">
										<a class="button_L2" href="javascript:void(0);" title="Apre la modifica del prezzo in una nuova finestra" <%= ACTIVE_STATUS %>
										   onclick="OpenAutoPositionedScrollWindow('ArticoliGiacenze_Mod.asp?ID=<%= rs("gia_id") %>', 'Giacenze', 510, 380, false)">
											MODIFICA
										</a>
									</td>
								</tr>
								<%rs.movenext
							wend
						end if
					else %>
						<tr>
							<td class="noRecords">Per visualizzare l'elenco degli articoli eseguire prima una ricerca.</td>
						</tr>
					<%end if%>
				</table>
				</form>
			</td> 
		</tr>
	</table>	
</div>

</body>
</html>
<% 
conn.close 
set rs = nothing
set rsc = nothing
set conn = nothing

sub Ricerca(conn, rsc) %>
	<form action="" method="post" id="ricerca" name="ricerca">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
			<tr>
				<td>
					<table cellspacing="1" cellpadding="0" class="tabella_madre">
						<caption>Opzioni di ricerca</caption>
						<tr>
							<td class="footer" colspan="2">
								<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
								<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
							</td>
						</tr>
						<tr><th colspan="2" <%= Search_Bg("mag_codice_int;mag_codice_pro") %>>CODICI</td></tr>
						<tr>
							<td class="label">interno:</td>
							<td class="content">
								<input type="text" name="search_codice_int" value="<%= TextEncode(Session("mag_codice_int")) %>" style="width:100%;">
							</td>
						</tr>
						<tr>
							<td class="label">produttore:</td>
							<td class="content">
								<input type="text" name="search_codice_pro" value="<%= TextEncode(Session("mag_codice_pro")) %>" style="width:100%;">
							</td>
						</tr>
						<tr><th colspan="2" <%= Search_Bg("mag_nome") %>>NOME</td></tr>
						<tr>
							<td class="content" colspan="2">
								<input type="text" name="search_nome" value="<%= TextEncode(Session("mag_nome")) %>" style="width:100%;">
							</td>
						</tr>
						<tr><th colspan="2" <%= Search_Bg("mag_stato_catalogo") %>>STATO ARTICOLO A CATALOGO</td></tr>
						<tr>
							<td class="content" style="width:45%;">
								<input type="checkbox" class="checkbox" name="search_stato_catalogo" value="1" <%= chk(instr(1, session("mag_stato_catalogo"), "1", vbTextCompare)>0) %>>
								visibile
							</td>
							<td class="content">
								<input type="checkbox" class="checkbox" name="search_stato_catalogo" value="0" <%= chk(instr(1, Session("mag_stato_catalogo"), "0", vbTextCompare)>0) %>>
								non visibile
							</td>
						</tr>
						<tr><th colspan="2" <%= Search_Bg("mag_categoria") %>>CATEGORIA</td></tr>
						<tr>
							<td class="content" colspan="2">
								<% CALL categorie.WritePicker("ricerca", "search_categoria", session("mag_categoria"), false, true, 32) %>
							</td>
						</tr>
						<tr><th colspan="2" <%= Search_Bg("mag_marchio") %>>MARCHIO</td></tr>
						<tr>
							<td class="content" colspan="2">
								<%	sql = "SELECT * FROM gtb_marche ORDER BY mar_nome_it"
								CALL dropDown(conn, sql, "mar_id", "mar_nome_it", "search_marchio", Session("mag_marchio"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
							</td>
						</tr>
						<tr><th colspan="2" <%= Search_Bg("mag_full_text") %>>FULL-TEXT (tutti i campi)</td></tr>
						<tr>
							<td class="content" colspan="2">
								<input type="text" name="search_full_text" value="<%= TextEncode(Session("mag_full_text")) %>" style="width:100%;">
							</td>
						</tr>
						<tr>
							<td class="footer" colspan="2">
								<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
								<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</form>	
<% end sub %>