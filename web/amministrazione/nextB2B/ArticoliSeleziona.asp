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

'imposta le variabili iniziali
if request.querystring("SelectedPage")<>"" then
	Pager.Reset()
	Session("SelectedPage") = request("SelectedPage") & ".asp"		'pagina di destinazione dopo la selezione
	Session("TYPE") = request("TYPE")								'indica il tipo di selezione dell'articolo: se componente (C) o accessorio (A)
	Session("ExcludeID") = cInteger(request("Exclude_ID"))			'indica quale id di prodotto escludere
	response.redirect "ArticoliSeleziona.asp"
end if

%>
<%'--------------------------------------------------------
sezione_testata = "selezione dell'articolo" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("SelArt_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("SelArt_")
	end if
end if

dim conn, sql, rs, rsc, FieldID
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

if Session("TYPE")="C" then
	FieldID = "rel_id"
	'selezione componenti di bundle o confezioni
	sql = " SELECT * FROM gv_articoli " + _
		  " WHERE NOT " + SQL_IsTrue(conn, "art_se_bundle") + " AND NOT " & SQL_IsTrue(conn, "art_se_confezione") + _
		  " AND rel_id NOT IN (SELECT bun_articolo_id FROM gtb_bundle WHERE bun_bundle_id=" & Session("ExcludeID") & ")"
else
	FieldID = "art_id"
	'selezione di accessori
	sql = " SELECT * FROM gtb_articoli " + _
		  " WHERE art_id <> " & Session("ExcludeID")
end if


'filtra per codice interno
if Session("SelArt_cod_int")<>"" then
	if sql <>"" then sql = sql & " AND "
	if Session("TYPE")="C" then
		sql = sql & SQL_FullTextSearch(Session("SelArt_cod_int"), "art_cod_int;rel_cod_int")
	else
		sql = sql & "( " & SQL_FullTextSearch(Session("SelArt_cod_int"), "art_cod_int") & _
					" or (art_id IN (SELECT rel_art_id FROM grel_art_valori WHERE " & SQL_FullTextSearch(Session("SelArt_cod_int"), "rel_cod_int") & ") ) ) "
	end if
end if

'filtra per codice alternativo
if Session("SelArt_cod_alt")<>"" then
	if sql <>"" then sql = sql & " AND "
	if Session("TYPE")="C" then
		sql = sql & SQL_FullTextSearch(Session("SelArt_cod_alt"), "art_cod_alt;rel_cod_alt")
	else
		sql = sql & "( " & SQL_FullTextSearch(Session("SelArt_cod_alt"), "art_cod_alt") & _
					" or (art_id IN (SELECT rel_art_id FROM grel_art_valori WHERE " & SQL_FullTextSearch(Session("SelArt_cod_alt"), "rel_cod_alt") & ") ) ) "
	end if
end if

'filtra per codice produttore
if Session("SelArt_cod_pro")<>"" then
	if sql <>"" then sql = sql & " AND "
	if Session("TYPE")="C" then
		sql = sql & SQL_FullTextSearch(Session("SelArt_cod_pro"), "art_cod_pro;rel_cod_pro")
	else
		sql = sql & "( " & SQL_FullTextSearch(Session("SelArt_cod_pro"), "art_cod_pro") & _
					" or (art_id IN (SELECT rel_art_id FROM grel_art_valori WHERE " & SQL_FullTextSearch(Session("SelArt_cod_pro"), "rel_cod_pro") & ") ) ) "
	end if
end if

'filtra per nome
if Session("SelArt_nome")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & SQL_FullTextSearch(Session("SelArt_nome"), FieldLanguageList("art_nome_"))
end if

'filtra per categoria
if Session("SelArt_categoria")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & " art_tipologia_id IN (" & categorie.FoglieID(Session("SelArt_categoria")) & " ) "
end if

'filtra per marca
if Session("SelArt_marchio")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & " art_marca_id=" & Session("SelArt_marchio")
end if

sql = sql + " ORDER BY art_nome_it"

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
		<th colspan="6" <%= Search_Bg("SelArt_cod_int;SelArt_cod_alt;SelArt_cod_pro") %>>CODICE</th>
		<th <%= Search_Bg("SelArt_marchio") %> nowrap>MARCHIO / PRODUTTORE</th>
	</tr>
	<tr>
		<td class="label" style="width:10%;">interno:</td>
		<td class="content"><input type="text" class="text" name="search_cod_int" value="<%= session("SelArt_cod_int") %>" maxlength="50" size="7"></td>
		<td class="label" style="width:6%;">alt.:</td>
		<td class="content">
			<input type="text" class="text" name="search_cod_alt" value="<%= session("SelArt_cod_alt") %>" maxlength="50" size="7">
		</td>
		<td class="label" style="width:6%;">prod.:</td>
		<td class="content">
			<input type="text" class="text" name="search_cod_pro" value="<%= session("SelArt_cod_pro") %>" maxlength="50" size="7">
		</td>
		<td class="content">
			<%	sql = "SELECT * FROM gtb_marche ORDER BY mar_nome_it"
			CALL dropDown(conn, sql, "mar_id", "mar_nome_it", "search_marchio", session("SelArt_marchio"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
		</td>
	</tr>
	<tr>
		<th colspan="2" <%= Search_Bg("SelArt_nome") %>>NOME</th>
		<th colspan="5" <%= Search_Bg("SelArt_categoria") %>>CATEGORIA</th>
	</tr>
	<tr>
		<td class="content" colspan="2"><input type="text" class="text" name="search_nome" value="<%= session("SelArt_nome") %>" maxlength="50" style="width=95%"></td>
		<td class="content" colspan="5">
			<% CALL categorie.WritePicker("ricerca", "search_categoria", session("SelArt_categoria"), false, false, 48) %>
		</td>
	</tr>
</table>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption class="border">Elenco articoli</caption>
	<tr>
		<td>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
				<tr>
					<td class="label_no_width" colspan="5">
						<% if rs.eof then %>
							Nessun articolo trovato.
						<% else %>
							Trovati n&ordm; <%= Pager.recordcount %> articoli in n&ordm; <%= Pager.PageCount %> pagine
						<% end if %>
					</td>
				</tr>
				<%if not rs.eof then %>
					<tr>
						<th class="L2">CODICE</th>
						<th class="L2">ARTICOLO</th>
						<% if Session("TYPE") = "A" then %>
							<th class="l2_center">VARIANTI</th>
						<% end if %>
						<th class="l2_center">A CATALOGO</th>
						<th class="l2_center" width="15%">SELEZIONA</th>
					</tr>
					<%rs.AbsolutePage = Pager.PageNo
					while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
						<tr>
							<td class="content">
								<% if Session("TYPE") = "C" then %>
									<%= rs("rel_cod_int") %>
								<% else %>
									<%= rs("art_cod_int") %>
								<% end if %>
							</td>
							<td class="content">
								<% CALL ArticoloLink(rs("art_id"), rs("art_nome_it"), rs("art_cod_int")) %>
								<% if Session("TYPE") = "C" then %>
									<% if rs("art_varianti") then %>
										<%= ListValoriVarianti(conn, rsc, rs("rel_id")) %>
									<% else %>
										&nbsp;
									<% end if %>
								<% else %>
									</td>
									<td class="content_center">
										<input type="checkbox" class="checkbox" disabled <%= chk(rs("art_varianti")) %>>
								<% end if %>
							</td>
							<td class="content_center"><input type="checkbox" class="checkbox" disabled <%= chk(not rs("art_disabilitato")) %>></td>
							<td class="content_center">
								<a tabindex="<%= rs.AbsolutePosition %>" class="button" href="<%= Session("SelectedPage") %>?EXT_ID=<%= Session("ExcludeID") %>&ARTID=<%= rs(FieldID) %>" title="Seleziona questo articolo" <%= ACTIVE_STATUS %>>SELEZIONA</a>
							</td>
						</tr>
						<% rs.MoveNext
					wend%>
					<tr>
						<td colspan="5" class="footer">
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