<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<%
dim Pager
set Pager = new PageNavigator

%>
<%'--------------------------------------------------------
sezione_testata = "selezione dell'articolo" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("sap_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("sap_")
	end if
end if

dim conn, sql, rs, rsc, rso, qta
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")
set rso = Server.CreateObject("ADODB.RecordSet")

if request("CAR_ID")<>"" then
	'selezione delle righe di un carico di magazzino
	sql = "SELECT *, (0) AS riv_LstCod_id FROM gtb_carichi WHERE car_id=" & cIntero(request("CAR_ID"))
	rso.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

	'elenca gli articoli per il carico
	sql = " SELECT * FROM gv_articoli INNER JOIN grel_giacenze ON gv_articoli.rel_id = grel_giacenze.gia_art_var_id " + _
		  " WHERE gia_magazzino_id=" & rso("car_magazzino_id") & " AND NOT " + SQL_IsTrue(conn, "art_se_bundle") + _
		  " AND rel_id NOT IN ( SELECT rcv_art_var_id FROM grel_carichi_var WHERE rcv_car_id=" & cIntero(request("CAR_ID")) & ")"
	
else 
	'selezione delle righe di un ordine
	sql = " SELECT * FROM gtb_ordini INNER JOIN gv_rivenditori ON gtb_ordini.ord_riv_id = gv_rivenditori.riv_id " + _
		  " INNER JOIN gtb_listini ON gv_rivenditori.riv_listino_id = gtb_listini.listino_id " + _
		  " LEFT JOIN gtb_lista_codici ON gv_rivenditori.riv_LstCod_id = gtb_lista_codici.lstCod_id " + _
		  " WHERE ord_id=" & cIntero(request("ORD_ID"))
	rso.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		  
	'elenca articoli per l'ordine
	sql = " SELECT * FROM gv_listino_vendita INNER JOIN grel_giacenze " & _
		  " ON gv_listino_vendita.rel_id = grel_giacenze.gia_art_var_id AND gia_magazzino_id=" & rso("ord_magazzino_id")
	if cInteger(rso("riv_LstCod_id"))>0 then
		'aggiunge codifica dei codici personalizzata per rivenditore
		sql = sql + " LEFT JOIN gtb_codici ON (gv_listino_vendita.rel_id = gtb_codici.cod_variante_id AND cod_lista_id=" & rso("riv_LstCod_id") &") "
	end if
	sql = sql + " WHERE " + RivenditoreListinoCondition(rso("riv_listino_id"), rso("listino_base_attuale"))
end if

'filtra per codice interno
if Session("sap_codice")<>"" then
	sql = sql &" AND "& SQL_FullTextSearch(Session("sap_codice"), "art_cod_int;rel_cod_int;art_cod_alt;rel_cod_alt")
end if

'filtra per codice produttore
if Session("sap_cod_pro")<>"" then
    sql = sql &" AND "& SQL_FullTextSearch(Session("sap_cod_pro"), "art_cod_pro;rel_cod_pro")
end if

'filtra per codice rivenditore
if Session("sap_cod_cli")<>"" then
    sql = sql &" AND "& SQL_FullTextSearch(Session("sap_cod_cli"), "cod_codice")
end if

'filtra per nome
if Session("sap_nome")<>"" then
	sql = sql &" AND "& SQL_FullTextSearch(Session("sap_nome"), FieldLanguageList("art_nome_"))
end if

'filtra per categoria
if Session("sap_categoria")<>"" then
	sql = sql & " AND art_tipologia_id IN (" & categorie.FoglieID(Session("sap_categoria")) & " ) "
end if

'filtra per marca
if Session("sap_marchio")<>"" then
	sql = sql & " AND art_marca_id=" & Session("sap_marchio")
end if

sql = sql + " ORDER BY art_nome_it, rel_cod_int"
CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)%>
<div id="content_ridotto">
<form action="" method="post" id="ricerca" name="ricerca">
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption>
		<table border="0" cellspacing="0" cellpadding="1" align="right">
			<tr>
				<td style="font-size: 1px; padding-right:1px;" nowrap>
					<input type="submit" name="cerca" value="CERCA" class="button">
					&nbsp;
					<input type="submit" name="tutti" value="VEDI TUTTI" class="button">
				</td>
			</tr>
		</table>
		Opzioni di ricerca
	</caption>
	<tr>
		<th colspan="6" <%= Search_Bg("sap_codice;sap_cod_pro;sap_cod_cli") %>>CODICE ARTICOLO</th>
		<th <%= Search_Bg("sap_nome") %>>NOME ARTICOLO</th>
	</tr>
	<tr>
		<td class="label" style="width:8%;">interno:</td>
		<td class="content" style="width:15%;">
			<input type="text" class="text" name="search_codice" value="<%= session("sap_codice") %>" maxlength="50" size="7">
		</td>
		<td class="label" style="width:10%;">produttore:</td>
		<td class="content" style="width:15%;">
			<input type="text" class="text" name="search_cod_pro" value="<%= session("sap_cod_pro") %>" maxlength="50" size="7">
		</td>
		<% If request("ORD_ID")<>"" AND cInteger(rso("riv_LstCod_id"))>0 then %>
			<td class="label" style="width:6%;">cliente:</td>
			<td class="content" style="width:15%;">
				<input type="text" class="text" name="search_cod_cli" value="<%= session("sap_cod_cli") %>" maxlength="50" size="7">
			</td>
			<td class="content">
		<% else %>
			<td class="content" colspan="3">
		<% End If %>
			<input type="text" class="text" name="search_nome" value="<%= session("sap_nome") %>" maxlength="50" style="width=100%">
		</td>
	</tr>
	<tr>
		<th colspan="3" <%= Search_Bg("sap_marchio") %>>MARCHIO / PRODUTTORE</th>
		<th colspan="4" <%= Search_Bg("sap_categoria") %>>CATEGORIA</th>
	</tr>
	<tr>
		<td class="content" colspan="3">
			<%	sql = "SELECT * FROM gtb_marche ORDER BY mar_nome_it"
			CALL dropDown(conn, sql, "mar_id", "mar_nome_it", "search_marchio", session("sap_marchio"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
		</td>
		<td class="content" colspan="4">
			<% CALL categorie.WritePicker("ricerca", "search_categoria", session("sap_categoria"), false, false, 60) %>
		</td>
	</tr>
</table>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption class="border">Elenco articoli</caption>
	<tr>
		<td>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
				<tr>
					<td class="label_no_width" colspan="8">
						<% if rs.eof then %>
							Nessun articolo trovato.
						<% else %>
							Trovati n&ordm; <%= Pager.recordcount %> articoli in n&ordm; <%= Pager.PageCount %> pagine
						<% end if %>
					</td>
				</tr>
				<%if not rs.eof then %>
					<tr>
						<th class="L2" colspan="2" style="border-bottom:0px;">codici</th>
						<th class="L2" rowspan="2">ARTICOLO</th>
						<% If request("ORD_ID")<>"" then %>
							<th class="l2_center" rowspan="2">dispo.</th>
							<th class="l2_center" colspan="3">PREZZO</th>
						<% else %>
							<th class="l2_center" colspan="2" style="border-bottom:0px;">QUANTIT&Agrave;</th>
						<% End If %>
						<th class="l2_center" rowspan="2" width="14%">SELEZIONA</th>
					</tr>
					<tr>
						<th class="L2">interno</th>
						<% If request("ORD_ID")<>"" AND cInteger(rso("riv_LstCod_id"))>0 then %>
							<th class="L2">cliente</th>
						<% else %>
							<th class="L2">produttore</th>
						<% end if 
						If request("CAR_ID")<>"" then%>
							<th class="L2">giacenza</th>
							<th class="L2">impegnato</th>
						<% else %>
							<th class="l2_center">netto</th>
							<th class="l2_center">&nbsp;</th>
							<th class="l2_center">i.v.a.</th>
						<% End If %>
					</tr>
					<%rs.AbsolutePage = Pager.PageNo
					while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
						<tr>
							<td class="content"><%= rs("rel_cod_int") %></td>
							<% If request("ORD_ID")<>"" AND cInteger(rso("riv_LstCod_id"))>0 then %>
								<td class="content">
							<%		if NOT isNull(rs("cod_codice")) then %>
							<%= 		rs("cod_codice") %>
							<% 		else %>
							<%= 		rs("rel_cod_pro") %>
							<% 		end if %>
								</td>
							<% else %>
								<td class="content"><%= rs("rel_cod_pro") %></td>
							<% End If %>
							<td class="content" width="45%">
								<% CALL ArticoloLink(rs("art_id"), rs("art_nome_it"), rs("rel_cod_int")) %>
								<% if rs("art_varianti") then %>
									<%= ListValoriVarianti(conn, rsc, rs("rel_id")) %>
								<% end if %>
							</td>
							<% if request("ORD_ID")<>"" then 
								qta = rs("gia_qta") - rs("gia_impegnato")%>
								<% if qta <= 0  then %>
									<td class="content_center alert">0</td>
								<% else %>
									<td class="content_center ok"><%= qta %></td>
								<% end if%>
								<%'recupera prezzi articolo
								CALL ScontiQ(conn, rs, rsc, 1, rso("valu_cambio"), rso("valu_simbolo")) %>
								<td class="content_center">
									<% if rs("listino_offerte") then %>
										<span class="Icona Offerte" title="prodotto in offerta speciale">&nbsp;</span>
									<% elseif rs("prz_promozione") then %>
										<span class="Icona Promozioni" title="prodotto in promozione">&nbsp;</span>
									<% else %>
										&nbsp;
									<% end if %>
								</td>
								<td class="content_center">
									<%= FormatPrice(rs("iva_valore"), 2, true) %>%
								</td>
								<td class="content_center">
									<a tabindex="<%= rs.AbsolutePosition %>" class="button" href="Ordini_dettagliNew.asp?ORD_ID=<%= request("ORD_ID") %>&PRZ_ID=<%= rs("prz_id") %>&ART_ID=<%= rs("rel_id") %>" title="Seleziona questo articolo" <%= ACTIVE_STATUS %>>SELEZIONA</a>
								</td>
							<% else %>
								<% qta = cInteger(rs("gia_qta"))
								if qta < 1 then %>
									<td class="content_center alert">
								<% elseif cInteger(rs("rel_giacenza_min"))=0 OR qta > cInteger(rs("rel_giacenza_min")) then %>
									<td class="content_center ok">
								<% else %>
									<td class="content_center warning">
								<% end if %>
									<%= rs("gia_qta") %>
								</td>
								<td class="content"><%= rs("gia_impegnato") %></td>
								<td class="content_center">
									<a tabindex="<%= rs.AbsolutePosition %>" class="button" href="MagazziniCarichi_dettagliNew.asp?CAR_ID=<%= request("CAR_ID") %>&ART_ID=<%= rs("rel_id") %>&MAG_ID=<%= rso("car_magazzino_id") %>" title="Seleziona questo articolo" <%= ACTIVE_STATUS %>>SELEZIONA</a>
								</td>
							<% End If %>
							
						</tr>
						<% rs.MoveNext
					wend%>
					<tr>
						<td colspan="8" class="footer">
							<table width="100%" cellpadding="0" cellspacing="0">
								<tr>
									<td><% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%></td>
									<td align="right">
										<a class="button" href="javascript:window.close();" title="chiudi la finestra" <%= ACTIVE_STATUS %>>
											CHIUDI</a>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				<% end if %>
			</table>
		</td>
	</tr>
</form>
</table>
</div>
</body>
</html>

<%
rs.close
set rsc = nothing
set rs = nothing
conn.close
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>
