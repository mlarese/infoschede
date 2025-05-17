<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 10000000 %>
<%
Imposta_Proprieta_Sito("ID_WEB")
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim ID
if request("ID") <> "" then
	ID = cIntero(request("ID"))
elseif request("INDICE") <> "" then
	ID = cIntero(GetValueList(index.conn, NULL, "SELECT idx_link_pagina_id FROM tb_contents_index WHERE idx_id="& cIntero(request("INDICE"))))
end if

if NOT index.content.ChkPrmF("tb_pagineSito", ID) then
	session("ERRORE") = "Non si possiedono i permessi per modificare la pagina."
	response.redirect "SitoPagine.asp"
end if

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoPagineSalva.asp")
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - indice delle pagine - modifica della pagina"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = IIF(request("FROM")=FROM_ALBERO, "SitoPagineAlbero.asp", "SitoPagine.asp")
CALL dicitura.InitializeIndex(Index, "tb_pagineSito", ID)
dicitura.scrivi_con_sottosez()

dim conn, rs, rsA, sql, i, lingua, n_lingue
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")
set rsA = server.CreateObject("ADODB.recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("WEB_PAGINE_SQL"), "id_pagineSito", "SitoPagineMod.asp?FROM="& request("FROM"))
end if

sql = " SELECT * FROM tb_PagineSito INNER JOIN tb_webs " &_
	  " ON tb_PagineSito.id_web=tb_webs.id_webs WHERE id_paginesito=" & ID
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>

<script language="JavaScript" type="text/javascript">
	//pubblicazione
	function azione_pagina(action, page_source, page_dest, lingua, nome_lingua){
		OpenAutoPositionedWindow('SitoPagineCopia.asp?ID_S=' + page_source + '&ID_D=' + page_dest + '&lingua=' + lingua + '&nome_lingua=' + nome_lingua + '&azione=' + action, 
				 				 'action', 500, 200);
    }
	
	function azione_template(page_source, nome_lingua){
		OpenAutoPositionedWindow('SitoPagineTemplate.asp?ID_STAGE=' + page_source + '&nome_lingua=' + nome_lingua, "template", 500, 250)
	}

	function azione_modifica( id_pagina ) {
		OpenAutoPositionedScrollWindow('loadshock.asp?PAGINA='+id_pagina, 'editor', document.body.clientWidth, screen.height, true);
	}
	
	function azione_vedi(id_pagina, lingua) {
		OpenPositionedScrollWindow('dynalay.asp?PAGINA='+id_pagina+'&lingua='+lingua, 'vedi', window.screenLeft, window.screenTop, document.body.clientWidth, 600, true);
	}
</script>

<div id="content">
<% CALL Ceck_page_exists(conn, rs) %>
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_id_web" value="<%=rs("id_web")%>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption class="border">	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption" title="<%= GetPageNumbers(rs) %>">Modifica layout della pagina n&ordm; <%= ID %></td>
					<td align="right" style="font-size:1px; padding-bottom:1px;">
						<a class="button" href="?FROM=<%= request("FROM") %>&ID=<%= ID %>&goto=PREVIOUS" title="pagina precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?FROM=<%= request("FROM") %>&ID=<%= ID %>&goto=NEXT" title="pagina successiva" <%= ACTIVE_STATUS %>>
							SUCCESSIVA &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<% if cIntero(rs("id_home_page")) = ID OR _
			  cIntero(rs("id_home_page_riservata")) = ID OR _
			  cIntero(rs("sito_in_aggiornamento_pagina")) = ID OR _
			  cIntero(rs("sito_in_costruzione_pagina")) = ID OR _
			  cIntero(rs("errore_pagina")) = ID then %>
			<tr>
				<td class="label">stato</td>
				<td colspan="2">
					<table border="0" cellpadding="0" cellspacing="0" width="100%">
						<% if cIntero(rs("id_home_page")) = ID then %>
							<tr>
								<td class="content_b homepage">Home page del sito</td>
							</tr>
						<% end if
						if cIntero(rs("id_home_page_riservata")) = ID then%>
							<tr>
								<td class="content_b homeareariservata">Home page dell'area riservata</td>
							</tr>
						<% end if
                        if cIntero(rs("id_login_page_riservata")) = ID then%>
							<tr>
								<td class="content_b loginareariservata">Pagina di login dell'area riservata</td>
							</tr>
						<% end if
						if cIntero(rs("errore_pagina")) = ID then%>
							<tr>
								<td class="content_b paginaerrore">Pagina di errore</td>
							</tr>
						<% end if
						if cIntero(rs("sito_in_costruzione_pagina")) = ID then%>
							<tr>
								<td class="content_b incostruzione">Home page del sito in fase di costruzione</td>
							</tr>
						<% end if
						if cIntero(rs("sito_in_aggiornamento_pagina")) = ID then%>
							<tr>
								<td class="content_b inaggiornamento">Home page del sito in fase di aggiornamento</td>
							</tr>
						<% end if %>
					</table>
				</td>
			</tr>
		<%end if %>
        <tr>
            <td class="label">riferimento interno</td>
            <td class="content">
                <input type="text" class="text" name="tft_nome_ps_interno" value="<%= rs("nome_ps_interno") %>" maxlength="255" size="35">
            </td>
            <td class="note">
                nota aggiuntiva visualizzata ed utilizzata solo all'interno del NEXT-web.
            </td>
        </tr>
    </table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<% dim page_id_STAGE, page_id_DYN
		n_lingue = 0
			for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
				lingua = Application("LINGUE")(i)
				if Session("LINGUA_" + lingua) then
					n_lingue = n_lingue + 1
					page_id_STAGE = cInteger(rs("id_pagStage_" & lingua))
					page_id_DYN = cInteger(rs("id_pagDyn_" & lingua)) %>
			<tr>
				<th class="L2" style="padding-top:1px;padding-bottom:1px;"><img src="../grafica/flag_mini_<%= lingua %>.jpg" border="0"></th>
				<th class="L2" colspan="5"><span style="text-transform:lowercase;">VERSIONE LINGUA</span> <%= GetNomeLingua(lingua)  %></th>
			</tr>
			<tr>
				<td class="label_no_width" style="width:3%;">titolo:</td>
				<td class="content" colspan="5" style="width:97%;">
					<input type="text" class="text" name="tft_nome_ps_<%= lingua %>" value="<%= rs("nome_ps_"& lingua) %>" maxlength="250" style="width:80%;">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2" style="width:30%;">apri la pagina di lavoro in modifica</td>
				<td class="content" style="width:5%;">
					<a HREF="javascript:void(0);" onClick="azione_modifica(<%=page_id_STAGE%>)" class="button_L2_block"
					   title="apri la pagina di lavoro n&ordm; <%= page_id_STAGE %> in modifica" <%= ACTIVE_STATUS %>>
						MODIFICA
					</a>
				</td>
				<td class="content" style="width:16%;">&nbsp;</td>
				<td class="label_right" colspan="2" style="">
					pagina di lavoro: 
					<a class="button_L2" href="javascript:void(0);" onclick="azione_vedi('<%= page_id_STAGE %>', '<%= lingua %>')"
					   title="apre in una nuova finestra la pagina di lavoro n&ordm; <%= page_id_STAGE %>" <%= ACTIVE_STATUS %>>
						VEDI
					</a>
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">pubblica la pagina sostituendo la pagina attualmente visibile al pubblico</td>
				<td class="content">
					<a HREF="javascript:void(0);" class="button_L2_block" onclick="azione_pagina('PUBBLICA', <%= page_id_STAGE %>, <%= page_id_DYN %>, '<%= lingua %>', '<%=  GetNomeLingua(lingua) %>')"
					   title="pubblica la pagina di lavoro sostituendo la pagina attualmente visibile al pubblico" <%= ACTIVE_STATUS %>>
						PUBBLICA
					</a>
				</td>
				<% if must_be_published(conn, rsA, page_id_STAGE, page_id_DYN) then %>
					<td class="content dapubblicare">
						da pubblicare
					</td>
				<% else %>
					<td class="content">&nbsp;</td>
				<% end if %>
				
				<td class="label_right" colspan="2">
					pagina pubblicata: 
					<a class="button_L2" href="javascript:void(0);" onclick="azione_vedi('<%= page_id_DYN %>', '<%= lingua %>')"
					   title="apre in una nuova finestra la pagina pubblicata n&ordm; <%= page_id_DYN %>" <%= ACTIVE_STATUS %>>
						VEDI
					</a>
				</td>
			</tr>
			<%	'recupera dati del template
				sql = "SELECT tb_templates.id_page, tb_templates.nomepage, tb_templates.semplificata " + _
					  " FROM tb_pages LEFT JOIN tb_pages tb_templates " & _
				 	  " ON tb_pages.id_template=tb_templates.id_page " & _
					  " WHERE tb_pages.id_page=" & page_id_STAGE
				rsA.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText %>
			<tr>
				<td class="label_no_width" colspan="2">
					apri la finestra di gestione dell'associazione ai template
				</td>
				<td class="content">
					<a HREF="javascript:void(0);" class="button_L2_block" onclick="azione_template('<%=page_id_STAGE%>', '<%=  GetNomeLingua(lingua) %>')"
					   title="apri la finestra di gestione dell'associazione ai template" <%= ACTIVE_STATUS %>>
						TEMPLATE
					</a>
				</td>
				<td class="label_no_width">
					Template:
					<% 	if cString(rsA("nomepage"))<>"" then  %>
						<a HREF="dynalay.asp?PAGINA=<%=rsA("id_page")%>&lingua=it" target="_blank" 
						   title="<%= rsA("nomepage") & vbCrLf%>apre in una nuova finestra il template" <%= ACTIVE_STATUS %>>
							<%= rsA("nomepage") %>
							<% if rsA("semplificata") then %>
								<img src="../grafica/notReadKnow.gif" border="0" alt="Template per email con visualizzazione semplificata.">
							<% end if %>
						</a>
					<%else %>
						nessuno
					<% end if %>
				</td>
				<td class="label_right" colspan="2">
					strumenti di pagina:
					<a HREF="javascript:void(0);" onClick="OpenAutoPositionedWindow('SitoPagineStrumenti.asp?PAGINA=<%= ID %>&LINGUA=<%= lingua %>', 'strumenti', 500, 250)" class="button_L2">
						APRI
					</a>
				</td>
			</tr>
			<% 	rsA.close
			end if
		next %>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption class="border">Modifica propriet&agrave; generali della pagina</caption>
		<tr>
			<td class="label_no_width">gestisci la descrizione della pagina e le parole chiave per i motori di ricerca</td>
			<td class="content_right" style="width:21%;">
				<a HREF="javascript:void(0);" onClick="OpenAutoPositionedScrollWindow('SitoPagineMetaTag.asp?ID=<%= ID %>', 'metatag', 640, 250, true)" class="button_L2_block"
				   title="Modifica delle parole chiave e della descrizione per i motori di ricerca." <%= ACTIVE_STATUS %>>
					DESCRIZIONE E META TAG
				</a>
			</td>
		</tr>
		<tr>
			<td class="label_no_width">pubblica tutte le pagine di lavoro sostituendo le pagine attualmente visibili al pubblico</td>
			<td class="content_right">
				<a HREF="javascript:void(0);" onClick="OpenAutoPositionedWindow('SitoPaginePubblicaTutte.asp?PAGINA=<%= ID %>', 'action', 500, 250)" class="button_L2_block"
				   title="Permette la pubblicazione di tutte le versioni della pagina." <%= ACTIVE_STATUS %>>
					PUBBLICA TUTTE LE LINGUE
				</a>
			</td>
		</tr>
		<tr>
			<td class="label_no_width">apri la finestra di gestione dei template della pagina</td>
			<td class="content_right">
				<a HREF="javascript:void(0);" onClick="OpenAutoPositionedWindow('SitoPagineTemplateTutte.asp?PAGINA=<%= ID %>', 'template', 500, 250)" class="button_L2_block"
				   title="apri la finestra di gestione dei template di tutte le pagine" <%= ACTIVE_STATUS %>>
					TEMPLATE DELLE PAGINE
				</a>
			</td>
		</tr>
		<!--
		<tr>
			<td class="label_no_width">apri l'elenco dei plugin usati nella pagina</td>
			<td class="content_right">
				<a HREF="SitoPlugin.asp?PAGINA=<%=ID%>" target="_blank" class="button_L2_block" title="apri l'elenco dei plugin usati nelle pagine" <%= ACTIVE_STATUS %>>
					PLUGIN USATI NELLE PAGINE
				</a>
			</td>
		</tr>
		-->
		<% sql = " SELECT COUNT(*) FROM v_indice WHERE NOT (tab_name LIKE 'tb_paginesito') AND " + _
                 " ( co_link_pagina_id=" & rs("id_pagineSito") & " OR idx_link_pagina_id = " & rs("id_pagineSito") & ")"
		if cIntero(GetValueList(conn, rsA, sql))>0 then %>
			<tr>
				<td class="label_no_width">visualizza l'elenco dei nodi dell'indice e dei relativi contenuti pubblicati per mezzo della pagina</td>
				<td class="content_right">
					<a HREF="javascript:void(0);" onClick="OpenAutoPositionedScrollWindow('SitoPagineElencoContenuti.asp?PAGINA=<%= ID %>', 'elenco_contenuti', 700, 400, true)" class="button_L2_block"
					   title="visualizza l'elenco dei nodi dell'indice e dei relativi contenuti pubblicati per mezzo della pagina" <%= ACTIVE_STATUS %>>
						CONTENUTI PUBBLICATI
					</a>
				</td>
			</tr>
		<% end if %>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<tr><th class="L2" colspan="5">stato della pagina</th></tr>
		<% if IsAreaRiservataActive(conn) then %>
			<tr>
				<td class="label_no_width">
	                <img src="../grafica/padlock.gif" border="0" alt="Pagina appartenente all'area protetta">
	                protetta
	            </td>
				<td class="content" style="width:15%;">
					<input type="radio" class="checkbox" value="1" name="tfn_riservata" value="1" <%= chk(rs("riservata")) %>>
					si
					<input type="radio" class="checkbox" value="0" name="tfn_riservata" value="0" <%= chk(not rs("riservata")) %>>
					no
				</td>
				<td class="content_right" colspan="3">
					<span class="note">
						Proteggendo la pagina se ne permette la visualizzazione solo agli utenti dell'area riservata.
					</span>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label">
                <img src="../grafica/archiviata.gif" border="0" alt="Pagina archiviata">
                archiviata
            </td>
			<td class="content" colspan="4">
				<input type="radio" class="checkbox" value="1" name="tfn_archiviata" value="1" <%= chk(rs("archiviata")) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_archiviata" value="0" <%= chk(not rs("archiviata")) %>>
				no
			</td>
		</tr>
		<tr>
			<td class="label">
                <img src="../grafica/indicizzazione.gif" border="0" alt="Pagina indicizzabile">
                indicizzabile
            </td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_indicizzabile" value="1" <%= chk(rs("indicizzabile")) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_indicizzabile" value="0" <%= chk(not rs("indicizzabile")) %>>
				no
			</td>
			<td class="content_right" colspan="3" style="width:70%;">
				<span class="note">
					Scegliere se rendere indicizzabile questa pagina dai motori di ricerca come Google, Bing, Yahoo, ecc.
				</span>
			</td>
		</tr>
		<% if cIntero(rs("id_home_page")) = ID then
			sql = " SELECT * FROM v_indice "& _
					 " WHERE co_F_key_id = "& rs("id_webs") & _
					 " AND co_F_table_id = "& index.GetTable("tb_webs") & _
					 " ORDER BY idx_principale DESC, idx_ordine_assoluto, idx_ordine, idx_tipologie_padre_lista"
		else
			sql = " SELECT * FROM v_indice "& _
					 " WHERE co_F_key_id = "& rs("id_pagineSito") & _
					 " AND co_F_table_id = "& index.GetTable("tb_pagineSito") & _
					 " ORDER BY idx_principale DESC, idx_ordine_assoluto, idx_ordine, idx_tipologie_padre_lista"
		end if
		
		rsA.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		if not rsA.eof then  %>
			<tr><th class="L2" colspan="5">pubblicazioni sull'indice</th></tr>
			<% while not rsA.eof 
				for each lingua in Application("LINGUE")
					if Session("LINGUA_" + lingua) then %>
						<tr>
							<% if lingua = LINGUA_ITALIANO then %>
								<td class="label_no_width <% if not(rsa("idx_principale") or rsa.recordcount=1) then %> notes_disabled<%end if %>" rowspan="<%= n_lingue%>">
									id: <%= rsa("idx_id")%>
								</td>
							<% end if %>
							<td class="content" colspan="3">
								<table cellpadding="0" cellspacing="0">
									<tr>
										<td style="vertical-align:top; padding-right:10px;">
											<img src="../grafica/flag_mini_<%= lingua %>.jpg" border="0">
										</td>
										<% if rsa("idx_principale") or rsa.recordcount=1 then %>
											<td style="vertical-align:top;" title="Pubblicazione principale del contenuto.">
												<% CALL index.WriteNodeLink(rsa, "", lingua) %>
												- 
												<% CALL index.WriteNodeLabelLink(rsa, "versione di lavoro", "?stage=1", "style=""font-size:9px; """, lingua)
										else%>
											<td class="content notes_disabled" style="vertical-align:top; " title="Pubblicazione alternativa">
												<%= index.NomeCompleto(rsa("idx_id"))%>
										<%end if%>
											   
										</td>
									</tr>
								</table>
							</td>
							<% if lingua = LINGUA_ITALIANO then %>
								<td  style="width:13%;" class="content_right" rowspan="<%= n_lingue%>">
									<% CALL index.WriteCollegamentoButton("_L2", rsa("co_F_table_id"), rsa("co_F_key_id"), rsa("idx_id")) %>
								</td>
							<% end if %>
						</tr>
					<% end if
				next
				rsa.movenext
			wend
		end if
		rsA.close
		%>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<%CALL Form_DatiModifica_EX(conn, rs, "ps_", "Dati del record", "L2") %>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="elenco" value="SALVA & TORNA ALL'ELENCO" style="width:23%;">
				<input type="submit" class="button" name="mod" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<%
rs.close
set rs = nothing
set rsA = nothing
conn.Close
set conn = nothing
%>