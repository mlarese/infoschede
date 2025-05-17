<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../NEXTb2b/Tools_B2B.asp" -->
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
	CALL SearchSession_Reset("SelArt_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("SelArt_")
	end if
end if

dim conn, sql, rs, rsc, FieldID, fieldCod
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

FieldID = "rel_id"
fieldCod = "rel_cod_int"
sql = " SELECT * FROM gv_articoli WHERE (1=1) "

'filtra per codice interno
if Session("SelArt_cod_int")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("SelArt_cod_int"), "art_cod_int;rel_cod_int")
end if
'filtra per codice alternativo
if Session("SelArt_cod_alt")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("SelArt_cod_alt"), "art_cod_alt;rel_cod_alt")
end if

'filtra per codice produttore
if Session("SelArt_cod_pro")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("SelArt_cod_pro"), "art_cod_pro;rel_cod_pro")
end if

'filtra per nome
if Session("SelArt_nome")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("SelArt_nome"), FieldLanguageList("art_nome_"))
end if

'filtra per categoria
if Session("SelArt_categoria")<>"" then
	sql = sql & " AND art_tipologia_id IN (" & categorie.FoglieID(Session("SelArt_categoria")) & " )"
end if

'filtra per marca
if Session("SelArt_marchio")<>"" then
	sql = sql & " AND art_marca_id=" & Session("SelArt_marchio")
end if

sql = sql + " ORDER BY art_nome_it"
CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)

%>

<script language="JavaScript">
function seleziona(id) {
	opener.<%= request("formname") %>.<%= request("inputname") %>.value = id
	var nome = document.getElementById("nome_articolo_" + id);
	opener.<%= request("formname") %>.view_<%= request("inputname") %>.value = nome.innerText;
	window.close()
}
</script>
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
		<th colspan="6">CODICE</th>
		<th nowrap>MARCHIO / PRODUTTORE</th>
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
		<th colspan="2">NOME</th>
		<th colspan="5">CATEGORIA</th>
	</tr>
	<tr>
		<td class="content" colspan="2"><input type="text" class="text" name="search_nome" value="<%= session("SelArt_nome") %>" maxlength="50" style="width=95%"></td>
		<td class="content" colspan="5">
			<% CALL categorie.WritePicker("ricerca", "search_categoria", session("SelArt_categoria"), false, false, 42) %>
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
						<th class="l2_center" width="15%">SELEZIONA</th>
					</tr>
					<%rs.AbsolutePage = Pager.PageNo
					while not rs.eof and rs.AbsolutePage = Pager.PageNo
						%>
						<tr>
							<td class="content">
								<%= rs("rel_cod_int") %>
							</td>
							<td class="content" id="nome_articolo_<%= rs("rel_id") %>">
								<% CALL ArticoloLink(rs("art_id"), rs("art_nome_it"), rs("art_cod_int")) %>
								<% if rs("art_varianti") then %>
									<%= ListValoriVarianti(conn, rsc, rs("rel_id")) %>
								<% else %>
									&nbsp;
								<% end if %>
							</td>
							<td class="content_center">
								<a tabindex="<%= rs.AbsolutePosition %>" class="<%= IIF(cInteger(request("selected"))=rs("rel_id"), "button_disabled", "button") %>" href="javascript:seleziona('<%= rs("rel_id") %>')" title="Seleziona questo articolo" <%= ACTIVE_STATUS %>>SELEZIONA</a>
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