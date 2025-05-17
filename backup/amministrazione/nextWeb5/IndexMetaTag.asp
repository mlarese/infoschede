<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="IndexMetaTag_TOOLS.asp" -->
<% 
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_strumenti_accesso, 0))

if request.querystring("FROM") <> "" then
	Session("FROM_AREA") = lcase(request.querystring("FROM"))
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
if Session("FROM_AREA") = "indice" then
	dicitura.sezione = "Indice generale - Gestione meta tag"
	dicitura.puls_new = "INDIETRO"
	dicitura.link_new = "IndexGenerale.asp"
else
	dicitura.sezione = "Gestione meta tag"
	dicitura.puls_new = "INDIETRO"
	dicitura.link_new = "SitoAnalisi.asp"
end if
dicitura.scrivi_con_sottosez()  

dim conn, rs, rsi, sql, Pager, LingueVisualizzate
dim Lingua, Visual, TagTitle, MetaDescription, MetaKeywords, Title, TextValue
dim IndexValue, ContentValue, PageValue, BaseValue, DefaultLingueList
set Pager = new PageNavigator

if isArray(Session("LINGUE")) then
	if ubound(Session("LINGUE"))>=0 then
		DefaultLingueList = Session("LINGUE")
	else 
		DefaultLingueList = Application("LINGUE")
	end if
else
	DefaultLingueList = Application("LINGUE")
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsi = Server.CreateObject("ADODB.RecordSet")

'imposta eventuali filtri di ricerca
sql = IndexSearchEngineSetFilter(conn, true)

if cString(Session("idx_indicizzabile")) = "" then
	Session("idx_indicizzabile") = "1"
end if

'applica filtro visualizzando solo indicizzabili
if cintero(Session("idx_indicizzabile"))=1 then
	sql = sql + IIF(sql<>"", " AND ", "") + _
		  " ((ISNULL(TIP_L0.idx_principale,0) = 1 AND " + _
		  "  " + _ 
		  " TIP_L0.idx_link_pagina_id IN (SELECT id_paginesito FROM tb_pagineSito WHERE ISNULL(riservata,0) = 0 AND  ISNULL(indicizzabile,0) = 1 ) " + _
		  " ) " + _
		  " OR isNull(TIP_L0.idx_link_url_it,'')<>'') "
end if

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	if request("lingue_visualizzate")<>"" then
		Session("meta_tag_lingue_visualizzate") = request("lingue_visualizzate")
	else
		Session("meta_tag_lingue_visualizzate") = join(DefaultLingueList, ",")
		LingueVisualizzate = DefaultLingueList
	end if
end if

if Session("meta_tag_lingue_visualizzate") <> "" then
	LingueVisualizzate = split(replace(Session("meta_tag_lingue_visualizzate"), " ", ""), ",")
else
	Session("meta_tag_lingue_visualizzate") = join(DefaultLingueList, ",")
	LingueVisualizzate = DefaultLingueList
end if

sql = index.QueryElenco(false, sql)
%>
 
<div id="content_liquid">
<%
CALL Pager.OpenSmartRecordset(conn, rs, sql, 50)
%>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption>Gestione meta tag dell'indice - opzioni di ricerca e filtri di visualizzazione</caption>
		<tr>
			<th <%= Search_Bg("idx_chiave") %> style="width:10%">CODICE UNIVOCO</th>
			<th <%= Search_Bg("idx_titolo") %> style="width:14%">TITOLO</th>
			<th colspan="2" <%= Search_Bg("idx_visibile") %>>VISIBILE</th>
			<th <%= Search_Bg("idx_livello") %> style="width:12%">LIVELLO DELLA VOCE</th>
			<th <%= Search_Bg("idx_tipoContenuto") %>>TIPO DELLA VOCE</th>
			<th <%= Search_Bg("idx_categoria") %>>VOCI COLLEGATE A</th>
		</tr>
		<form action="" method="post" id="ricerca" name="ricerca">
		<tr>
			<td class="content">
				<input type="text" name="search_chiave" value="<%= TextEncode(session("idx_chiave")) %>" style="width:100%;">
			</td>
			<td class="content">
				<input type="text" name="search_titolo" value="<%= TextEncode(session("idx_titolo")) %>" style="width:100%;">
			</td>
			<td class="content" style="width:8%">
				<input type="checkbox" class="checkbox" name="search_visibile" value="0" <%= chk(instr(1, Session("idx_visibile"), "1", vbTextCompare)>0) %>>
				visibile
			</td>
			<td class="content" style="width:8%">
				<input type="checkbox" class="checkbox" name="search_visibile" value="0" <%= chk(instr(1, session("idx_visibile"), "0", vbTextCompare)>0) %>>
				non visibile
			</td>
			<td class="content">
			<% 	sql = "SELECT MAX(idx_livello) FROM tb_contents_index"
				dim levels, i
				set levels = Server.CreateObject("Scripting.Dictionary")
				CALL levels.Add("0", "Voci base")
				for i=1 to cInteger(GetValueList(Index.conn, NULL, sql))
					CALL levels.Add(cString(i), "Voci livello " & i)
				next
				CALL DropDownDictionary(levels, "search_livello", Session("idx_livello"), false, "style=""width:100%;""", LINGUA_ITALIANO)%>
			</td>
			<td class="content">
				<% CALL index.content.DropDownTipi("search_tipoContenuto", "", session("idx_tipoContenuto")) %>
			</td>
			<td class="content" style="width:325px;">
				<% CALL index.WritePicker("", "", "ricerca", "search_categoria", session("idx_categoria"), 0, false, false, 40, false, false) %>
			</td>
		</tr>
		<tr>
			<th colspan="4">VISUALIZZAZIONE ED ATTIVAZIONE LINGUE</th>
			<th colspan="3">VISIBILIT&Agrave; AI MOTORI DI RICERCA</th>
		</tr>
		<tr>
			<td class="label_no_width">Lingue visualizzate:</td>
			<td class="content" colspan="3">
				<% for each lingua in Application("LINGUE") %>
					<input type="Checkbox" class="noBorder" name="lingue_visualizzate" value="<%= lingua %>" <%= chk(Session("meta_tag_lingue_visualizzate") = "" OR instr(1, Session("meta_tag_lingue_visualizzate"), lingua, vbTextCompare)>0) %>>
					<img src="<%= GetAmministrazionePath() %>grafica/flag_<%= lingua %>.jpg" style="margin-right:20px;">
				<% next %>
			</td>
			<td class="content">
				<input type="checkbox" class="checkbox" name="search_indicizzabile" value="1" <%= chk(instr(1, Session("idx_indicizzabile"), "1", vbTextCompare)>0) %>>
				solo indicizzabili dai motori
			</td>
			<td class="content" colspan="2">
				<input type="checkbox" class="checkbox" name="search_indicizzabile" value="0" <%= chk(instr(1, session("idx_indicizzabile"), "0", vbTextCompare)>0) %>>
				tutti
				<span class="note">Vengono visualizzati anche gli url non principali, o le pagine protette o non indicizzabili.</span>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="7">
				<input type="submit" name="cerca" value="CERCA" class="button" style="width:90px;">
				<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width:90px;">
			</td>
		</tr>
		</form>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0;">
		<caption>Gestione meta tag dell'indice - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</caption>
		<tr>
			<th colspan="3">LEGENDA COLORI UTILIZZATI</th>
		</tr>
		<% CALL WriteLegenda( "width:4%;", 2 ) %>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<% if not rs.eof then %>
            <tr>
               <th rowspan="3">Voce</th>
			   <th class="center" colspan="<%= 3 * (ubound(LingueVisualizzate) + 1) %>" style="border-bottom:0px;">STATO META TAG</th>
			   <th rowspan="3" class="center" style="width:60px;">OPERAZIONI</th>
            </tr>
			<tr>
				<th class="center" style="border-bottom:0px;" colspan="<%= (ubound(LingueVisualizzate) + 1) %>">TITLE</th>
				<th class="center" style="border-bottom:0px;" colspan="<%= (ubound(LingueVisualizzate) + 1) %>">KEYWORDS</th>
				<th class="center" style="border-bottom:0px;" colspan="<%= (ubound(LingueVisualizzate) + 1) %>">DESCRIPTIONS</th>
			</tr>
			<tr>
				<% for each lingua in LingueVisualizzate %>
					<th class="center"><img src="<%= GetAmministrazionePath() %>grafica/flag_mini_<%= lingua %>.jpg"></td>
				<% next %>
				<% for each lingua in LingueVisualizzate %>
					<th class="center"><img src="<%= GetAmministrazionePath() %>grafica/flag_mini_<%= lingua %>.jpg"></td>
				<% next %>
				<% for each lingua in LingueVisualizzate %>
					<th class="center"><img src="<%= GetAmministrazionePath() %>grafica/flag_mini_<%= lingua %>.jpg"></td>
				<% next %>
			</tr>
			<% rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo 
				sql = " SELECT v_indice.*, " + _
					  SQL_MultiLanguage("nome_ps_<LINGUA>, PAGE_keywords_<LINGUA>, PAGE_description_<LINGUA>", ",") + "," + _
					  SQL_MultiLanguage("META_keywords_<LINGUA>, META_description_<LINGUA>", ",") + _
					  " FROM (v_indice " + _
					  " LEFT JOIN tb_paginesito ON v_indice.idx_link_pagina_id = tb_paginesito.id_pagineSito ) " + _
					  " LEFT JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + _
					  " WHERE idx_id=" & rs("idx_id")
				rsi.open sql, conn, adOpenStatic, adLockOptimistic, adcmdtext
				%>
				<tr>
					<td class="content">
						<% if not index.content.IsRaggruppamento(rsi("tab_name")) then %>
							<span style="float:right;">
								<% if index.content.IsSito(rsi("tab_name")) AND cIntero(rsi("co_F_key_id")) = cIntero(Session("AZ_ID")) then %>
									<a class="button_L2" title="Apre la modifica del sito." target="sito_<%= rs("idx_id") %>"
									   href="SitoMod.asp?ID=<%= rsi("co_F_key_id") %>" style="margin-right:3px;">
										SITO
									</a>
								<% elseif index.content.IsPagina(rsi("tab_name")) then %>
									<a class="button_L2" title="Apre la modifica della pagina" target="pagina_<%= rsi("co_F_key_id") %>"
									   href="SitoPagineMod.asp?ID=<%= rsi("co_F_key_id") %>" style="margin-right:3px;">
										PAGINA
									</a>
								<% end if %>
								<a class="button_L2" title="Apre la modifica del nodo dell'indice." target="indice_<%= rs("idx_id") %>"
								   href="IndexGestione.asp?ID=<%= rs("idx_id") %>&FROM=<%= FROM_ELENCO %>" style="margin-right:3px;">
									VOCE
								</a>
								<% CALL index.WriteButton(rsi("tab_name"), rs("co_F_key_id"), POS_INDICE) %>
							</span>
						<% end if %>
						<%= rs("NAME") %>&nbsp;
						<%= Index.Content.WriteTipoRS(rsi) %>
					</td>
					<% if cString(rsi("idx_link_url_it"))="" AND index.content.IsRaggruppamento(rsi("tab_name")) then %>
						<td class="label_no_width" colspan="<%= 1 + (3 * (ubound(LingueVisualizzate) + 1)) %>">I dati non sono impostabili per i raggruppamenti.</td>
					<% elseif cString(rsi("idx_link_url_it"))="" AND cIntero(rsi("idx_link_pagina_id")) = 0 AND lcase(rsi("tab_name")) <> "tb_webs" then %>
						<td class="label_no_width" colspan="<%= 1 + (3 * (ubound(LingueVisualizzate) + 1)) %>">I dati non sono impostabili per i contenuti che puntano a link esterni.</td>
					<% else
						for each lingua in LingueVisualizzate 
							TagTitle = GetTitle(cString(rsi("idx_alt_" + Lingua)), cString(rsi("co_alt_" + Lingua)), cString(rsi("nome_ps_" + Lingua)), cString(rsi("co_titolo_" + Lingua)), _
												Lingua, Visual, TextValue, Title)
												
							
							%>
							<td class="label_no_width" title="<%= Title %>" style="white-space:nowrap;">
								<% writeIcon(visual) %>
								<% CALL WriteSyncroLock(rsi("tab_field_titolo_alt_" & lingua)) %>
								<span style="white-space:nowrap;"><%= TextValue %></span>
							</td>
						<% next
						
						for each lingua in LingueVisualizzate
							MetaKeywords = GetKeywords(cString(rsi("idx_meta_keywords_" + Lingua)), cString(rsi("co_meta_keywords_" + Lingua)), cString(rsi("PAGE_keywords_" + Lingua)), cString(rsi("META_keywords_" + Lingua)), _
									   Lingua, Visual, TextValue, Title)
							%>
							<td class="label_no_width" title="<%= Title %>" style="white-space:nowrap;">
								<% writeIcon(visual) %>
								<% CALL WriteSyncroLock(rsi("tab_field_titolo_alt_" & lingua)) %>
								<span style="white-space:nowrap;"><%= TextValue %></span>
							</td>
						<% next
						
						for each lingua in LingueVisualizzate
							MetaDescription = GetDescription(cString(rsi("idx_meta_description_" + Lingua)), cString(rsi("co_meta_description_" + Lingua)), cString(rsi("PAGE_description_" + Lingua)), cString(rsi("META_description_" + Lingua)), _
											 Lingua, Visual, TextValue, Title)
							%>
							<td class="label_no_width" title="<%= Title %>" style="white-space:nowrap;">
								<% writeIcon(visual) %>
								<% CALL WriteSyncroLock(rsi("tab_field_titolo_alt_" & lingua)) %>
								<span style="white-space:nowrap;"><%= TextValue %></span>
							</td>
						<% next %>
						<td class="content_center">
							<a HREF="javascript:void(0);" onClick="OpenAutoPositionedScrollWindow('IndexMetaTagModifica.asp?ID=<%= rs("idx_id") %>', 'modifica_metatag_<%= rs("idx_id") %>', 740, 350, true)" class="button_block"
						   	   title="Modifica delle parole chiave e della descrizione per i motori di ricerca per la voce dell'indice ed il contenuto." <%= ACTIVE_STATUS %>>
								MODIFICA
							</a>
						</td>
					<% end if %>
				</tr>
				<% rsi.close
				rs.movenext
			wend %>
            <tr>
			    <td colspan="<%= 3 + (3 * (ubound(LingueVisualizzate) + 1)) %>" class="footer" style="text-align:left;">
			        <% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
		        </td>
			</tr>
        <% else %>
            <caption>Voci dell'indice</caption>
            <tr><td class="noRecords">Nessun record trovato</th></tr>
        <% end if
        rs.close %>
	</table>
</div>
</body>
</html>
<%
conn.close
set rs = nothing
set rsi = nothing
set conn = nothing

%>