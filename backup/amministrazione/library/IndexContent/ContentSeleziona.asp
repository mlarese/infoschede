<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../ClassPageNavigator.asp" -->
<%
'--------------------------------------------------------
sezione_testata = "selezione del contenuto" %>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim Pager, sql, rs, aux, lingua, campi_lingua
set Pager = new PageNavigator
set rs = Server.CreateObject("ADODB.RecordSet")

'imposta le variabili iniziali
if request.querystring("formname")<>"" AND request.servervariables("REQUEST_METHOD") <> "POST" then
    Pager.Reset()
	'imposta parametri iniziali per apertura elenco
	Session("prmCo_FormName") = request.querystring("formname")
	Session("prmCo_InputName") = request.querystring("InputName")
	Session("prmCo_Selected") = request.querystring("selected")
	
    if request.querystring("tipo") = "gestione" AND cInteger(request.querystring("selected"))>0 then
        response.redirect "ContentGestione.asp?FROM=selezione&co_F_key_id="& request("co_F_key_id") &"&co_F_table_id="& request("co_F_table_id") & "&ID=" & request.querystring("selected") & "&selection_disabled=" & request.querystring("selection_disabled")
    else
        response.redirect "ContentSeleziona.asp?co_F_key_id="& request("co_F_key_id") &"&co_F_table_id="& request("co_F_table_id")
    end if
end if
	
if Request.ServerVariables("REQUEST_METHOD")="POST" then
    Pager.Reset()
	CALL SearchSession_Reset("selco_")
	if not(request("tutti")<>"") then
	    CALL SearchSession_Set("selco_")
    end if
end if

sql = ""
if session("SelCo_titolo") <> "" then
    sql = sql & " AND " & SQL_FullTextSearch(Session("SelCo_titolo"), FieldLanguageList("co_titolo_"))
end if

if session("SelCo_tipo") <> "" then
    sql = sql & " AND co_F_table_id = "& cIntero(session("SelCo_tipo"))
end if

campi_lingua = ""
for each lingua in Application("LINGUE")
	campi_lingua = campi_lingua & " co_link_url_"& lingua & ", co_link_url_rw_" & lingua & ", "
next

sql = " SELECT "&campi_lingua&" co_id, co_titolo_it, co_link_tipo, co_link_pagina_id, co_visibile, co_F_table_id, co_F_key_id, " & _
	  " tab_id, tab_field_url_it, tab_colore, tab_titolo " & _
	  " FROM tb_contents c"& _
      " INNER JOIN tb_siti_tabelle t ON c.co_F_table_id = t.tab_id"& _
      " WHERE "& index.content.SQLPermessi() & sql & _
      " ORDER BY co_titolo_it"

dim tabIndex
tabIndex = index.GetTable("tb_contents_index")
rs.open sql, index.conn, adOpenStatic, adLockOptimistic, adCmdText
rs.PageSize = 20 %>
    <script language="JavaScript" type="text/javascript">
	    function ResetLink() {
		    DisableControl(opener.<%= Session("prmCo_FormName") %>.idx_link_tipo_<%= lnk_interno %>, false)
			DisableControl(opener.<%= Session("prmCo_FormName") %>.idx_link_pagina_id, false)
			opener.<%= Session("prmCo_FormName") %>.idx_link_pagina_id.value = ""
			DisableControl(opener.<%= Session("prmCo_FormName") %>.idx_link_tipo_<%= lnk_esterno %>, false)
			<% 	for each lingua in Application("LINGUE") %>
			    DisableControl(opener.<%= Session("prmCo_FormName") %>.idx_link_url_<%= lingua %>, true)
				opener.<%= Session("prmCo_FormName") %>.idx_link_url_<%= lingua %>.value = ""
		    <% 	next %>
        }

        function Selezione(ID, close){
		    var v, e
			
            // setto il link
			ResetLink()
			v = document.getElementById("LINK_VINCOLATO_" + ID).value
			try {
			    if (v != "-1") {	// -1 = link non trovato
				    if (document.getElementById("LINK_TIPO_" + ID).value == "<%= lnk_interno %>") {
					    opener.<%= Session("prmCo_FormName") %>.idx_link_tipo_<%= lnk_interno %>.click()
						opener.<%= Session("prmCo_FormName") %>.idx_link_pagina_id.value = document.getElementById("LINK_PAG_" + ID).value
                    } 
                    else {
                        opener.<%= Session("prmCo_FormName") %>.idx_link_tipo_<%= lnk_esterno %>.click()
                        <% 	for each lingua in Application("LINGUE") %>
                            opener.<%= Session("prmCo_FormName") %>.idx_link_url_<%= lingua %>.value = document.getElementById("LINK_<%= lingua %>_" + ID).value
                        <% 	next %>
					}
				}
					
				// link vincolato
				if (v == "1") {
				    DisableControl(opener.<%= Session("prmCo_FormName") %>.idx_link_tipo_<%= lnk_interno %>, true)
					DisableControl(opener.<%= Session("prmCo_FormName") %>.idx_link_pagina_id, true)
					DisableControl(opener.<%= Session("prmCo_FormName") %>.idx_link_tipo_<%= lnk_esterno %>, true)
					<% 	for each lingua in Application("LINGUE") %>
					    DisableControl(opener.<%= Session("prmCo_FormName") %>.idx_link_url_<%= lingua %>, true)
					<% 	next %>
				}
			} 
            catch (e) {}
				
		    // setto il contenuto
			v = document.getElementById("NAME_" + ID).value
			try {
			    opener.<%= Session("prmCo_FormName") %>.<%= Session("prmCo_InputName") %>.value= ID;
				opener.<%= Session("prmCo_FormName") %>.view_<%= Session("prmCo_InputName") %>.value = v;
                try{
                    opener.<%= Session("prmCo_FormName") %>_<%= Session("prmCo_InputName") %>_UpdateTitle(<%= Session("prmCo_FormName") %>.view_<%= Session("prmCo_InputName") %>)
                } 
                catch(except){ }
            } 
            catch (e) {}
				
			if (close)
			    window.close();
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
				<th<%= Search_Bg("SelCo_tipo") %>>TIPO</th>
				<th<%= Search_Bg("SelCo_titolo") %>>TITOLO</th>
			</tr>
			<tr>
				<td class="content"><% CALL index.content.DropDownTipi("search_tipo", "", session("selCo_tipo")) %></td>
				<td class="content"><input type="text" class="text" name="search_titolo" value="<%= session("SelCo_titolo") %>" maxlength="50" style="width=98%"></td>
			</tr>
		</table>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
			<caption class="border">Elenco contenuti</caption>
			<tr>
				<td>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr>
							<td class="label_no_width" colspan="5">
								<% if rs.eof then %>
									Nessun contenuto trovato.
								<% else %>
									Trovati n&ordm; <%= rs.recordcount %> contenuti in n&ordm; <%= rs.PageCount %> pagine.
								<% end if %>
							</td>
						</tr>
						<%if not rs.eof then %>
							<tr>
								<th class="l2_center">SEL.</th>
								<th class="L2">TITOLO</th>
								<th class="l2_center" width="40px">VISIBILE</th>
								<th class="l2_center">OPERAZIONI</th>
							</tr>
							<%	rs.AbsolutePage = Pager.PageNo
								while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
								<input type="hidden" id="NAME_<%= rs("co_id") %>" name="NAME_<%= rs("co_id") %>" value="<%= TextEncode(rs("co_titolo_it")) %>">
								
							<%		'url vincolato
									if index.content.LinkPrecalcola(rs, true) then %>
								<input type="hidden" id="LINK_VINCOLATO_<%= rs("co_id") %>" name="LINK_VINCOLATO_<%= rs("co_id") %>" value="1">
							<%		elseif CString(rs("co_link_tipo")) <> "" then %>
								<input type="hidden" id="LINK_VINCOLATO_<%= rs("co_id") %>" name="LINK_VINCOLATO_<%= rs("co_id") %>" value="0">
							<%		else %>
								<input type="hidden" id="LINK_VINCOLATO_<%= rs("co_id") %>" name="LINK_VINCOLATO_<%= rs("co_id") %>" value="0">
							<%		end if %>
							<%		if CString(rs("co_link_tipo")) <> "" then %>
								<input type="hidden" id="LINK_TIPO_<%= rs("co_id") %>" name="LINK_TIPO_<%= rs("co_id") %>" value="<%= rs("co_link_tipo") %>">
								<input type="hidden" id="LINK_PAG_<%= rs("co_id") %>" name="LINK_PAG_<%= rs("co_id") %>" value="<%= rs("co_link_pagina_id") %>">
							<% 			for each lingua in Application("LINGUE") %>
								<input type="hidden" id="LINK_<%= lingua %>_<%= rs("co_id") %>" name="LINK_<%= lingua %>_<%= rs("co_id") %>" value="<%= rs("co_link_url_"& lingua) %>">
							<% 			next %>
							<% 		end if %>
								<tr>
									<td width="4%" class="content_center">
										<input type="radio" id="seleziona_<%= rs("co_id") %>" name="seleziona_<%= rs("co_id") %>" class="checkbox" value="<%= rs("co_id") %>" <%= Chk(cInteger(Session("prmCo_SELECTED")) = rs("co_id")) %>
												   title="Click per selezionare l'indice" onclick="Selezione(<%= rs("co_id") %>, true)">
										<%	if cInteger(Session("prmCo_SELECTED")) = rs("co_id") then %>
										<script type="text/javascript">
											Selezione(<%= rs("co_id") %>, false)
										</script>
										<%	end if %>
									</td>
									<td class="<%= IIF(rs("co_visibile"), "content", "content_disabled"" title=""voce non visibile") %>">
										<a href="javascript:void(0)" onclick="seleziona_<%= rs("co_id") %>.click()" title="<%= rs("co_titolo_it") %>" <%= ACTIVE_STATUS %>>
											<% CALL index.content.WriteNomeETipo(rs) %>
										</a>
									</td>
									<td class="content_center"><input type="checkbox" class="checkbox" disabled <%= chk(rs("co_visibile")) %>></td>
									<td class="content_center" width="90">
										<a class="button_L2" href="ContentGestione.asp?FROM=selezione&co_F_table_id=<%= rs("co_F_table_id") %>&co_F_key_id=<%= rs("co_F_key_id") %>&ID=<%= rs("co_id") %>">
											<% 	if index.content.IsAllFull(rs, "tab_") AND rs("tab_id") <> tabIndex then %>
											    VEDI
											<% 	else %>
											    COMPLETA I DATI
											<% 	end if %>
										</a>
									</td>
								</tr>
						<% 		rs.MoveNext
							wend %>
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
    </form>
</div>
<%
rs.close
set rs = nothing
%>
</body>
</html>
<script language="JavaScript" type="text/javascript">
    <!--
    FitWindowSize(this);
	//-->
</script>
