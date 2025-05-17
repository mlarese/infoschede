<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede_Categorie.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede_Const.asp" -->
<!--#INCLUDE FILE="../nextB2B/Tools_B2B.asp" -->
<%
dim Pager, conn, sql
set Pager = new PageNavigator
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")


if request("SALVA") = "true" then
	sql = " INSERT INTO srel_problemi_articoli(rpa_problema_id,rpa_articolo_rel_id) " & _
		  " VALUES (" & cIntero(request("IDPRB")) & ", " & cIntero(request("RELID")) & ") "
	conn.Execute(sql)
	%>
	<script language="JavaScript" type="text/javascript">
		opener.location.reload(true);
		window.close();
	</script>
	<%
end if

'imposta le variabili iniziali
if request.querystring("TYPE")<>"" then
	Pager.Reset()
	if request("SelectedPage") <> "" then
		Session("SelectedPage") = request("SelectedPage") & ".asp"		'pagina di destinazione dopo la selezione
	end if
	if request("IDPRB")	<> "" then
		Session("IDPRB") = request("IDPRB")								'id problema di riferimento
	end if
	if request("IDSCH")	<> "" then
		Session("IDSCH") = request("IDSCH")								'id scheda di riferimento
	end if
	if request("ID_EXT") <> "" then
		Session("ID_EXT") = request("ID_EXT")								'id di riferimento
	end if
	Session("TYPE") = request("TYPE")								'indica il tipo di selezione dell'articolo: se ricambio (R), se modelli (M)
	if cString(request("Exclude_IDS")) <> "" then
		Session("ExcludeID") = cString(request("Exclude_IDS"))			'indica quale id di prodotto escludere
	end if
	
	
	'dalla funzione WritePicker_ArticoloVariante
	'formname=form1&inputname=tfn_sc_modello_id&selected=34711
	if request("formname") <> "" then
		Session("formname") = request("formname")
	else
		Session("formname") = ""
	end if
	if request("inputname") <> "" then
		Session("campo_principale") = request("inputname")
	end if

	if request("selected") <> "" then
		Session("selected") = request("selected")
	end if
	
	if request("COSTR_ID") <> "" then
		Session("COSTR_ID") = request("COSTR_ID")
	end if
	
	if request("SUBMIT_AFTER") = "true" then
		Session("SUBMIT_AFTER") = true
	end if

	response.redirect "ArticoliSeleziona.asp"
end if


if Session("COSTR_ID") <> "" then
	CALL SearchSession_Reset("SelArt_")
	Session("SelArt_costruttore") = Session("COSTR_ID")
	Session("COSTR_ID") = ""
end if

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("SelArt_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("SelArt_")
	else
		Session("COSTR_ID") = ""
	end if
end if


dim rs, rsc, FieldID, labelSing, labelPlur
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

if Trim(cString(Session("ExcludeID"))) = "" then
	Session("ExcludeID") = 0
end if 

FieldID = "rel_id"
sql = " SELECT rel_cod_int, rel_id, art_id, art_nome_it, art_cod_int, art_varianti, art_disabilitato FROM gv_articoli " + _
	  " WHERE rel_id NOT IN (" & Session("ExcludeID") & ")"

if Session("TYPE")="R" then
	'selezione ricambi
	labelSing = "Ricambio"
	labelPlur = "Ricambi"
elseif Session("TYPE")="M" then
	'selezione modelli
	labelSing = "Modello"
	labelPlur = "Modelli"
else
	'selezione articoli
	labelSing = "Articolo"
	labelPlur = "Articoli"
end if


%>
<%'--------------------------------------------------------
sezione_testata = "selezione " & LCase(labelSing) %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 



'filtra per codice interno
if Session("SelArt_cod_int")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & SQL_FullTextSearch(Session("SelArt_cod_int"), "art_cod_int;rel_cod_int")
end if

'filtra per codice alternativo
if Session("SelArt_cod_alt")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & SQL_FullTextSearch(Session("SelArt_cod_alt"), "art_cod_alt;rel_cod_alt")
end if

'filtra per codice produttore
if Session("SelArt_cod_pro")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & SQL_FullTextSearch(Session("SelArt_cod_pro"), "art_cod_pro;rel_cod_pro")
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

'filtra per costruttore
if Session("SelArt_costruttore")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & " mar_anagrafica_id=" & Session("SelArt_costruttore")
end if

'filtra per marca
if Session("SelArt_marchio")<>"" then
	if sql <>"" then sql = sql & " AND "
	sql = sql & " art_marca_id=" & Session("SelArt_marchio")
end if


if Session("TYPE")="R" then
	sql = sql + " AND art_tipologia_id IN ("&cat_ricambi.DiscendentiID(0)&") " + " ORDER BY art_nome_it"
elseif Session("TYPE")="M" then
	sql = sql + " AND art_tipologia_id IN ("&cat_modelli.DiscendentiID(0)&") " + " ORDER BY art_nome_it"
else
	sql = sql + " AND art_tipologia_id IN ("&cat_articoli.DiscendentiID(0)&") " + " ORDER BY art_nome_it"
end if


CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)


if Session("formname")<>"" AND Session("campo_principale")<>"" then
	%>
	<script language="JavaScript">
	function seleziona(id) {
		opener.<%= Session("formname") %>.<%= Session("campo_principale") %>.value = id
		var nome = document.getElementById("nome_articolo_" + id);
		opener.<%= Session("formname") %>.view_<%= Session("campo_principale") %>.value = nome.innerText;
		<% if cBoolean(Session("SUBMIT_AFTER"), false) then %>
			opener.<%= Session("formname") %>.submit();
		<% end if %>
		window.close()
	}
	</script>
	<%
end if

%>

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
	</tr>
	<tr>
		<td class="label" style="width:10%;">interno:</td>
		<td class="content">
			<input type="text" class="text" name="search_cod_int" value="<%= session("SelArt_cod_int") %>" maxlength="50" size="10">
		</td>
		<td class="label" style="width:10%;">alternativo:</td>
		<td class="content">
			<input type="text" class="text" name="search_cod_alt" value="<%= session("SelArt_cod_alt") %>" maxlength="50" size="10">
		</td>
		<td class="label" style="width:10%;">produttore:</td>
		<td class="content">
			<input type="text" class="text" name="search_cod_pro" value="<%= session("SelArt_cod_pro") %>" maxlength="50" size="10">
		</td>
	</tr>
	<tr>
		<th <%= Search_Bg("SelArt_costruttore") %> nowrap colspan="3">COSTRUTTORE</th>
		<th <%= Search_Bg("SelArt_marchio") %> nowrap colspan="3">MARCHIO / PRODUTTORE</th>
	</tr>
	<tr>
		<td class="content" colspan="3">
			<%	sql = "SELECT * FROM gv_rivenditori WHERE riv_profilo_id = "&COSTRUTTORI&" ORDER BY NomeOrganizzazioneElencoIndirizzi "
			CALL dropDown(conn, sql, "riv_id", "NomeOrganizzazioneElencoIndirizzi", "search_costruttore", session("SelArt_costruttore"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
		</td>
		<td class="content" colspan="3">
			<%	sql = "SELECT * FROM gtb_marche ORDER BY mar_nome_it"
			CALL dropDown(conn, sql, "mar_id", "mar_nome_it", "search_marchio", session("SelArt_marchio"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
		</td>
	</tr>
	<tr>
		<th colspan="3" <%= Search_Bg("SelArt_nome") %>>NOME</th>
		<th colspan="3" <%= Search_Bg("SelArt_categoria") %>>CATEGORIA</th>
	</tr>
	<tr>
		<td class="content" colspan="3">
			<input type="text" class="text" name="search_nome" value="<%= session("SelArt_nome") %>" maxlength="50" style="width:95%;">
		</td>
		<td class="content" colspan="3">
			<% if Session("TYPE")="M" then 
				CALL cat_modelli.WritePicker("ricerca", "search_categoria", session("SelArt_categoria"), false, false, 48)
			elseif Session("TYPE")="R" then 
				CALL cat_ricambi.WritePicker("ricerca", "search_categoria", session("SelArt_categoria"), false, false, 48)
			else
				CALL cat_articoli.WritePicker("ricerca", "search_categoria", session("SelArt_categoria"), false, false, 48)
			end if %>
		</td>
	</tr>
</table>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption class="border">Elenco <%=lCase(labelPlur)%></caption>
	<tr>
		<td>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
				<tr>
					<td class="label_no_width" colspan="5">
						<% if rs.eof then %>
							Nessun&nbsp;<%=lCase(labelSing)%> trovato.
						<% else %>
							Trovati n&ordm; <%= Pager.recordcount %>&nbsp;<%=lCase(labelPlur)%> in n&ordm; <%= Pager.PageCount %> pagine
						<% end if %>
					</td>
				</tr>
				<%if not rs.eof then %>
					<tr>
						<th class="L2">CODICE</th>
						<th class="L2"><%=uCase(labelSing)%></th>
						<th class="l2_center">A CATALOGO</th>
						<th class="l2_center" width="15%">SELEZIONA</th>
					</tr>
					<%rs.AbsolutePage = Pager.PageNo
					while not rs.eof and rs.AbsolutePage = Pager.PageNo %>
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
							<td class="content_center"><input type="checkbox" class="checkbox" disabled <%= chk(not rs("art_disabilitato")) %>></td>
							<td class="content_center">
								<%
								'response.write Session("formname") 
								'response.end
								%>
								<% if Session("formname")<>"" AND not Session("TYPE")="R" then %>
									<a tabindex="<%= rs.AbsolutePosition %>" class="<%= IIF(cInteger(Session("selected"))=rs(FieldID), "button_disabled", "button") %>" href="javascript:seleziona('<%= rs(FieldID) %>')" title="Seleziona questo <%=LCase(labelSing)%>" <%= ACTIVE_STATUS %>>SELEZIONA</a>
								<% else %>
									<% if Session("TYPE")="M" then %>
										<a tabindex="<%= rs.AbsolutePosition %>" class="button" href="ArticoliSeleziona.asp?SALVA=true&IDPRB=<%=Session("IDPRB")%>&RELID=<%= rs(FieldID) %>" title="Associa questo modello" <%= ACTIVE_STATUS %>>ASSOCIA</a>
									<% elseif Session("TYPE")="R" then %>
										<a tabindex="<%= rs.AbsolutePosition %>" class="button" href="<%= Session("SelectedPage") %>?IDSCH=<%= Session("IDSCH")%>&RELID=<%= rs(FieldID) %>" title="Associa questo ricambio" <%= ACTIVE_STATUS %>>SELEZIONA</a>
									<% else %>
										<a tabindex="<%= rs.AbsolutePosition %>" class="button" href="<%= Session("SelectedPage") %>?ID_EXT=<%= Session("ID_EXT")%>&RELID=<%= rs(FieldID) %>" title="Associa questo articolo" <%= ACTIVE_STATUS %>>SELEZIONA</a>
									<% end if %>
								<% end if %>
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

<% if Session("TYPE")="R" then %>
	<div id="pulsanti" style="position:absolute; top:600px; left:4px; width:99%;">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption class="border">Ricambio non trovato?</caption>
			<tr>
				<td class="content" style="padding:2px; text-align:right;">
					<a nowrap target="DettScheda" class="button_L2" href="ArticoliNew.asp?IDSCH=<%=Session("IDSCH")%>&CATEGORIA=ricambio&STANDALONE=true" style="margin-left:4px;">
						AGGIUNGI IL RICAMBIO
					</a>
				</td>

			</tr>
		</table>
	</div>
<% end if %>

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