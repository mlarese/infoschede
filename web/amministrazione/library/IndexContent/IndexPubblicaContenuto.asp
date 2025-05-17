<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="ClassIndexAlberi.asp" -->
<!--#INCLUDE FILE="../ClassPageNavigator.asp" -->
<script language="JavaScript" type="text/javascript">
    function Selezione(objID, objNome){
        window.location.href = "?SELEZIONA=" + objID.value;
		//SelezionaNodo(objID.value, objNome.value);
    }
	
	function SelezionaNodo(IdValue, NameValue){
		var input = opener.document.getElementById("<%= Session("prmIdx_InputName") %>");
		<%	if Session("prmIdx_viewName") <> "" then %>
			var view = opener.document.getElementById("<%= Session("prmIdx_viewName") %>");
		<%	else %>
        	var view = opener.document.getElementById("view_<%= Session("prmIdx_InputName") %>");
		<%	end if %>
        input.value=IdValue;
        view.value=NameValue;

		<%	'se sto scegliendo il padre submit per refresh permessi associati
        if Session("prmIdx_InputName") = "idx_padre_id" then %>
            opener.<%= Session("prmIdx_FormName") %>.submit()
        <% 	end if %>
        window.close();
	}
	
	function VisualizzaCome_onclick(Albero){
		ricerca.VisualizzaComeAlbero.value = Albero;
		ricerca.submit();
	}
</script>
<%
'--------------------------------------------------------
sezione_testata = "pubblicazione del contenuto sul sito" %>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim Pager, rs, sql
set rs = Server.CreateObject("ADODB.RecordSet")
set Pager = new PageNavigator

'dichiara oggetto per generazione albero
dim oTree
set oTree = new ObjIndexTrees
set oTree.Index = Index


if cIntero(request("SELEZIONA"))>0 then
	%>
	<script language="JavaScript" type="text/javascript">
		//SelezionaNodo('<%= request("SELEZIONA") %>', '<%= replace(Index.NomeCompleto(request("SELEZIONA")), "'", "") %>');
	</script>
	<%
	Session("prmIdx_SELECTED") = ""
	CALL oTree.Tree.ResetExpansionState()
	response.redirect GetAmministrazionePath() & "library/IndexContent/IndicizzaAssocia.asp" & _
									"?co_F_table_id="&Session("co_F_table_id")&"&co_F_key_id="&Session("co_F_key_id")&_
									"&tab_pagina_default_id="&Session("tab_pagina_default_id")&"&idx_padre_id="&cIntero(request("SELEZIONA"))
									
	response.end
end if


'imposta le variabili iniziali
if request.querystring("formname")<>"" AND request.servervariables("REQUEST_METHOD") <> "POST" then
	Pager.Reset()
	'imposta parametri iniziali per apertura elenco
	Session("prmIdx_FormName") = request.querystring("formname")
	Session("prmIdx_InputName") = request.querystring("InputName")
	Session("prmIdx_viewName") = request.querystring("viewName")
    Session("prmIdx_SoloFoglie") = (request("SoloFoglie")<>"")
    Session("prmIdx_Selected") = request.querystring("selected")
	Session("prmIdx_WebId") = request.querystring("WebIdFilter")
	CALL oTree.Tree.ResetExpansionState()
	if cIntero(Session("prmIdx_Selected"))>0 then
		CALL oTree.ExpandNodes(Session("prmIdx_Selected"))
	end if
	
    Session("prmIdx_selezione_voci") = session("CONDIZIONE_SELEZIONE_VOCI_" & request.querystring("formname") & "_" & request.querystring("InputName"))
    Session("prmIdx_selezione_tipi") = session("CONDIZIONE_SELEZIONE_TIPI_" & request.querystring("formname") & "_" & request.querystring("InputName"))
    response.redirect "IndexPubblicaContenuto.asp"
end if

if Request.ServerVariables("REQUEST_METHOD")="POST" then
    Pager.Reset()
	CALL SearchSession_Reset("SelIdx_")
	if not(request("tutti")<>"") then
	    CALL SearchSession_Set("SelIdx_")
	end if
	
	if request.form("VisualizzaComeAlbero")<>"" then
		Session("VisualizzaComeAlbero") = cIntero(request.form("VisualizzaComeAlbero"))
	end if
end if

sql = ""
    
'imposta filtri di selezione degli elementi
if Session("prmIdx_selezione_voci")<>"" then
    sql = sql + IIF(sql<>"", " AND ", "") + Session("prmIdx_selezione_voci")
end if

if cIntero(Session("prmIdx_WebId"))>0 then
	sql = sql + IIF(sql<>"", " AND ", "") + " TIP_L0.idx_webs_id=" & Session("prmIdx_WebId")
end if

if Session("prmIdx_selezione_tipi") <> "" then
    sql = sql + IIF(sql<>"", " AND ", "") + " TIP_C0.co_F_table_id IN (SELECT tab_id FROM tb_siti_tabelle WHERE " & Session("prmIdx_selezione_tipi") & ") "
end if
    
if session("SelIdx_titolo") <> "" then
    sql = sql + IIF(sql<>"", " AND ", "") + SQL_FullTextSearch(Session("SelIdx_titolo"), FieldLanguageList("TIP_C0.co_titolo_"))
end if
	
if session("SelIdx_tipo") <> "" then
    sql = sql + IIF(sql<>"", " AND ", "") + "  TIP_C0.co_F_table_id = " & cIntero(session("SelIdx_tipo"))
end if


dim tab_pagina_default_id, co_F_table_id, co_F_key_id
if cIntero(request("tab_pagina_default_id")) > 0 then
	tab_pagina_default_id = cIntero(request("tab_pagina_default_id"))
	Session("tab_pagina_default_id") = tab_pagina_default_id
else
	tab_pagina_default_id = Session("tab_pagina_default_id")
end if

if cIntero(request("co_F_table_id")) > 0 then
	co_F_table_id = cIntero(request("co_F_table_id"))
	Session("co_F_table_id") = co_F_table_id
else
	co_F_table_id = Session("co_F_table_id")
end if

if cIntero(request("co_F_key_id")) > 0 then
	co_F_key_id = cIntero(request("co_F_key_id"))
	Session("co_F_key_id") = co_F_key_id
else
	co_F_key_id = Session("co_F_key_id")
end if
%>

<div id="content_ridotto">
	<form action="" method="post" id="ricerca" name="ricerca">
	<input type="hidden" name="VisualizzaComeAlbero" value="<%= Session("VisualizzaComeAlbero") %>">
	<input type="hidden" name="tab_pagina_default_id" value="<%= tab_pagina_default_id %>">
	<input type="hidden" name="co_F_table_id" value="<%= co_F_table_id %>">
	<input type="hidden" name="co_F_key_id" value="<%= co_F_key_id %>">
	<% if cIntero(Session("VisualizzaComeAlbero"))<>0 or cString(Session("VisualizzaComeAlbero"))="" then %>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
			<caption class="border">
				<table cellpadding="1" cellspacing="0" align="right">
					<tr>
						<td>
							<a class="button_L2" href="#" onclick="VisualizzaCome_onclick('0')" title="Apre la visualizzazione ad elenco.">
								VISUALIZZA COME ELENCO
							</a>
						</td>
					</tr>
				</table>
				Albero dell'indice generale
			</caption>
		</table>
		<% 
	    CALL oTree.AlberoIndiceSelezione(Session("prmIdx_SELECTED"), Session("SoloFoglie"), Session("AZ_ID"))
	else 
		sql = index.QueryElenco(Session("SoloFoglie"), sql)
		rs.open sql, index.conn, adOpenStatic, adLockReadOnly, adCmdText
		rs.PageSize = 20 %>
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
				<th<%= Search_Bg("SelIdx_tipo") %>>TIPO</th>
				<th<%= Search_Bg("SelIdx_titolo") %>>TITOLO</th>
			</tr>
			<tr>
				<td class="content"><% CALL index.content.DropDownTipi("search_tipo", Session("prmIdx_selezione_tipi"), session("selIdx_tipo")) %></td>
				<td class="content"><input type="text" class="text" name="search_titolo" value="<%= session("SelIdx_titolo") %>" maxlength="50" style="width=98%"></td>
			</tr>
		</table>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
			<caption class="border">
				<table cellpadding="1" cellspacing="0" align="right">
					<tr>
						<td>
							<a class="button_L2" href="#" onclick="VisualizzaCome_onclick('1')" title="Apre la visualizzazione ad albero.">
								VISUALIZZA COME ALBERO
							</a>
						</td>
					</tr>
				</table>
				Elenco voci dell'indice generale
			</caption>
			<tr>
				<td>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr>
							<td class="label_no_width" colspan="5">
								<% if rs.eof then %>
									Nessuna voce trovata.
								<% else %>
									Trovati n&ordm; <%= rs.recordcount %> voci in n&ordm; <%= rs.PageCount %> pagine.
								<% end if %>
							</td>
						</tr>
						<%if not rs.eof then %>
							<tr>
								<th class="l2_center">SEL.</th>
								<th class="L2">TITOLO COMPLETO INDICE</th>
								<th class="L2">VIS.</th>
							</tr>
							<%rs.AbsolutePage = Pager.PageNo
							dim i
							while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
								<input type="hidden" name="NAME_<%= rs("idx_id") %>" value="<%= replace(rs("NAME"), """", "'") %>">
								<tr>
									<td width="4%" class="content_center">
										<input type="radio" name="seleziona_<%= rs("idx_id") %>" class="checkbox" value="<%= rs("idx_id") %>" <%= Chk(cInteger(Session("prmIdx_SELECTED")) = rs("idx_id")) %>
												   title="Click per selezionare l'indice" onclick="Selezione(this, NAME_<%= rs("idx_id") %>)">
										</td>
									<td class="<%= IIF(rs("co_visibile"), "content", "content_disabled"" title=""voce non visibile") %>">
										<a href="javascript:void(0)" onclick="seleziona_<%= rs("idx_id") %>.click()" title="<%= rs("NAME") %>" <%= ACTIVE_STATUS %>>
											<% 	i = InStrRev(rs("NAME"), " - ")
												if i > 0 then
													response.write Left(rs("NAME"), i) & _
																   "<b>"& Right(rs("NAME"), Len(rs("NAME")) - i)
												else
													response.write "<b>"& rs("NAME")
												end if %></b>&nbsp;
											<% 	CALL index.content.WriteTipo(rs("co_F_table_id")) %>
										</a>
									</td>
									<td class="content"><input type="checkbox" class="checkbox" disabled <%= chk(rs("co_visibile")) %>></td>
								</tr>
								<% rs.MoveNext
							wend%>
							<tr>
								<td colspan="5" class="footer">
									<table width="100%" cellpadding="0" cellspacing="0">
										<tr>
											<td><% 	CALL Pager.Render_GroupNavigator(10, rs.PageCount, "", "button", "button_disabled")%></td>
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
		</table>
		<% rs.close
	end if %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption class="border" <%= IIF(cIntero(Session("VisualizzaComeAlbero"))<>0 or cString(Session("VisualizzaComeAlbero"))="", "style=""border-top:0px;""", "") %>>
			<table cellpadding="1" cellspacing="0" align="right">
				<tr>
					<td>
						<a class="button_L2" href="<%=GetAmministrazionePath() & "library/IndexContent/Indicizza.asp?co_F_table_id="&Session("co_F_table_id")&"&co_F_key_id="&Session("co_F_key_id")&"&MODE=standard" %>" title="Apre la finestra classica di collegamento all'indice.">
							MODALITA' STANDARD
						</a>
					</td>
				</tr>
			</table>
		</caption>
	</table>
	</form>
</div>
</body>
</html>

<%

set rs = nothing

%>
<script language="JavaScript" type="text/javascript">
	//VisualizzaCome_onclick('1');
<!--
	FitWindowSize(this);
//-->
</script>