<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_menu_accesso, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoMenuSalva.asp")
end if

dim conn, rs, rsl, sql, i, lingua
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")
set rsl = server.CreateObject("ADODB.recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("WEB_MENU_SQL"), "m_id", "SitoMenuMod.asp")
end if

sql = "SELECT * FROM tb_menu WHERE m_id="& cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - menu - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoMenu.asp"
dicitura.scrivi_con_sottosez() 
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati del menu</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="menu precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="menu successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="2">DATI DEL MENU</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
			lingua = Application("LINGUE")(i)
			if Session("LINGUA_" & lingua) then%>
				<tr>
				<% 	if i = 0 then %>
					<td class="label" rowspan="<%= Session("LINGUE_ATTIVE") %>">nome:</td>
				<% 	end if %>
					<td class="content">
						<img src="../grafica/flag_<%= lingua %>.jpg" width="26" height="15" alt="" border="0">
						<input type="text" class="text" name="tft_m_nome_<%= lingua %>" value="<%= rs("m_nome_"& lingua) %>" maxlength="255" size="100">
						<% 	if lingua = LINGUA_ITALIANO then response.write "(*)" end if %>
					</td>
				</tr>
			<%end if
		next %>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<tr><th colspan="4">GESTIONE DEI LINK</th></tr>
		<tr>
			<td class="content">
				<input type="radio" class="noBorder" name="metodo_gestione" id="metodo_gestione_manuale" onclick="SetStato()" value="manuale" <%= chk(cInteger(rs("m_index_id"))=0) %>>
			</td>
			<td class="content">
				inserimento manuale dei link
			</td>
			<td class="label_no_width" style="text-align:right;">
				copia dei link da un altro menu o da una voce dell'indice
			</td>
			<td class="content_center" style="width:21%;">
				<a class="button_L2_block" href="javascript:void(0)" id="link_copia"
				   onclick="OpenAutoPositionedWindow('SitoMenuLinksCopia.asp?ID=<%= rs("m_id") %>', 'MenuGestione', 500, 300)">
					COPIA LINKS
				</a>
			</td>
		</tr>
		<tr>
			<td class="content" rowspan="3" style="width:4%;">
				<input type="radio" class="noBorder" name="metodo_gestione" id="metodo_gestione_agganciata" onclick="SetStato()" value="agganciata" <%= chk(cInteger(rs("m_index_id"))>0) %>>
			</td>
			<td class="content" rowspan="3" style="width:25%;"">agganciato ad una voce dell'indice:</td>
			<% if cInteger(rs("m_index_id"))>0 then %>
				<td class="content_b" colspan="2">
					<%= index.NomeCompleto(rs("m_index_id"))  %>
				</td>
			<% else %>
				<td class="content_disabled" colspan="2">
					- -
				</td>
			<% end if %>
		</tr>
		<tr>
			<td class="label_no_width" style="text-align:right;">
				aggancia il menu ad una voce dell'indice
			</td>
			<td class="content_center">
				<a class="button_L2_block" href="javascript:void(0)" id="link_aggancia"
				   onclick="OpenAutoPositionedWindow('SitoMenuLinksAggancia.asp?ID=<%= rs("m_id") %>', 'MenuGestione', 500, 300)">
					<%= IIF(cInteger(rs("m_index_id"))>0, "CAMBIA VOCE AGGANCIATA", "AGGANCIA VOCE") %>
				</a>
			</td>
		</tr>
		<tr>
			<td class="label_no_width" style="text-align:right;">
				sgancia il menu dalla voce permettendo l'inserimento manuale dei link
			</td>
			<td class="content_center">
				<% if cInteger(rs("m_index_id"))>0 then %>
					<a class="button_L2_block" href="javascript:void(0)" id="link_sgancia"
					   onclick="OpenAutoPositionedWindow('SitoMenuLinksSgancia.asp?ID=<%= rs("m_id") %>', 'MenuGestione', 500, 300)">
					   	SGANCIA DALLA VOCE
					</a>
				<% else %>
					<a class="button_L2_block_disabled">
					   	SGANCIA DALLA VOCE
					</a>
				<% end if %>
			</td>
		</tr>
	</table>
	<script language="JavaScript" type="text/javascript">
		function SetStato(){
			var gestione_manuale = document.getElementById("metodo_gestione_manuale");
			
			EnableIfChecked(gestione_manuale, document.getElementById("link_copia"));
			DisableIfChecked(gestione_manuale,  document.getElementById("link_aggancia"));
			
			var sgancia = document.getElementById("link_sgancia");
			if (sgancia)
				DisableIfChecked(gestione_manuale, sgancia);
		}
		
		SetStato();
	</script>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<tr><th colspan="2">ELENCO DEI LINK</th></tr>
		<% if cInteger(rs("m_index_id"))>0 then
			'menu agganciato all'indice
			sql = " SELECT * FROM tb_contents_index INNER JOIN tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id" + _
				  "	WHERE idx_padre_id = " & rs("m_index_id") & _
				  " ORDER BY idx_ordine_assoluto, co_ordine "
		else
			'menu con link normali
			sql = " SELECT * FROM tb_menuItem WHERE mi_menu_id="& rs("m_id") &" ORDER BY mi_ordine "
		end if
		session("WEB_LINKS_SQL") = sql
		rsl.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>
		<tr>
			<td colspan="2">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<td class="label_no_width" colspan="2">
							<% if rsl.eof then %>
								Nessun link inserito.
							<% else %>
								Trovati n&ordm; <%= rsl.recordcount %> record
							<% end if %>
						</td>
						<% if cInteger(rs("m_index_id"))<1 then %>
							<td colspan="5" class="content_right" style="padding-right:0px;">
								<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('SitoMenuLinkNew.asp?MENU=<%= request("ID") %>', '_blank', 680, 405, true)"
								   title="Apre la finestra per l'inserimento di un nuovo link del menu." <%= ACTIVE_STATUS %>>
									NUOVO LINK
								</a>
								&nbsp;
								<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('SitoMenuLinkNewFromIndex.asp?MENU=<%= request("ID") %>', '_blank', 680, 405, true)"
								   title="Apre la finestra per la selezione di una voce dall'indice da cui copiare il nuovo link." <%= ACTIVE_STATUS %>>
									COPIA LINK DA VOCE DELL'INDICE
								</a>
							</td>
						<% end if %>
					</tr>
					<%if not rsl.eof then %>
						<tr>
							<th class="L2">TITOLO</th>
							<th class="L2">IMMAGINE</th>
							<th class="l2_center" width="7%;">ORDINE</th>
							<th class="l2_center" width="8%;">VISIBILE</th>
							<th class="l2_center" width="11%;">TIPO</th>
							<% if cInteger(rs("m_index_id"))<1 then %>
								<th class="l2_center" colspan="2" width="10%">OPERAZIONI</th>
							<% end if %>
						</tr>
						<% if cInteger(rs("m_index_id"))>0 then
							while not rsl.eof %>
								<tr>
									<td class="content"><%= rsl("co_titolo_it") %></td>
									<td class="content"><%= rsl("co_foto_thumb") %></td>
									<td class="content_center"><%= rsl("co_ordine") %></td>
									<td class="content_center"><input <%= chk(rsl("co_visibile")) %> type="checkbox" name="attivo" value="1" disabled class="checkbox"></td>
									<td class="content_center"><%= IIF(cIntero(rsl("co_link_tipo"))= LNK_INTERNO, "INTERNO", "ESTERNO")%></td>
								</tr>
								<% rsl.movenext
							wend
						else
							while not rsl.eof %>
								<tr>
									<td class="content"><%= rsl("mi_titolo_it") %></td>
									<td class="content"><%= rsl("mi_image_it") %></td>
									<td class="content_center"><%= rsl("mi_ordine") %></td>
									<td class="content_center"><input <%= chk(rsl("mi_attivo")) %> type="checkbox" name="attivo" value="1" disabled class="checkbox"></td>
									<td class="content_center"><%= IIF(cIntero(rsl("mi_index_id")) > 0, "INTERNO", "ESTERNO")%></td>
									<td style="vertical-align:middle;" class="content_center" width="5%">
										<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('SitoMenuLinkMod.asp?MENU=<%= request("ID") %>&ID=<%= rsl("mi_id") %>', 'MenuMod', 680, 405, true)"
									   	   title="Apre la finestra per l'inserimento di un nuovo link del menu." <%= ACTIVE_STATUS %>>
											MODIFICA
										</a>
									</td>
									<td style="vertical-align:middle;" class="content_center" width="5%">
										<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('MENUITEM','<%= rsl("mi_id") %>');" >
											CANCELLA
										</a>
									</td>
								</tr>
								<%rsl.movenext
							wend
						end if
					end if %>
				</table>
			</td>
		</tr>
		<% rsl.close %>
		<tr>
			<td class="footer" colspan="2">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
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
set rsl = nothing
conn.Close
set conn = nothing
%>