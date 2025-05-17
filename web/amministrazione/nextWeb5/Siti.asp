<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% 
Reset_Proprieta_Sito()
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - elenco"
if index.ChkPrm(prm_siti_gestione, 0) then
	dicitura.puls_new = "NUOVO SITO"
	dicitura.link_new = "SitoNew.asp"
end if
dicitura.scrivi_con_sottosez()

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT * FROM tb_webs w "& _
	  " LEFT JOIN tb_pagineSito p ON w.id_home_page = p.id_pagineSito "& _
	  " ORDER BY nome_webs"
session("WEB_SITI_SQL") = sql
rs.open sql, conn, adOpenStatic, adLockReadOnly, adAsyncFetch
%>
<div id="content">
	<% if not rs.eof then
		while not rs.eof %>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
			<caption class="border" style="padding-bottom:1px;">
				<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
						<td class="caption" >
						<%if rs("sito_mobile") then%>
							<table border="0" cellspacing="0" cellpadding="0" align="left">
								<tr>
									<td style="padding-top:1px;font-size:1px;">
										<img src="../grafica/mobile_icon.png" border="0" alt="Sito per dispositivi mobili.">
									</td>
								</tr>
							</table>
						<%end if%>						
						<%=rs("nome_webs") %>						
						</td>
						<% 	if index.ChkPrm(prm_siti_gestione, 0) then %>
						<td width="20%">
							<a class="button" href="SitoMod.asp?ID=<%= rs("id_webs") %>">
								MODIFICA
							</a>
                            &nbsp;
                            <% sql = "SELECT COUNT(*) FROM tb_pages WHERE id_webs = " & rs("id_webs")
                            if cIntero(GetValueList(conn, NULL, sql))>0 then %>
                                <a class="button_disabled" title="Sito non cancellabile: sono presenti delle pagine o dei template.">
                            <% else %>
							    <a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('SITI','<%= rs("id_webs") %>');" >
                            <% end if %>
								CANCELLA
							</a>
						</td>
						<% 	end if %>
					</tr>
				</table>
			</caption>
            <tr>
				<td class="label" colspan="2">home page:</td>
				<td class="content">
					<%= IIF(isNull(rs("nome_ps_it")), "non impostata", PaginaSitoNome(rs, "")) %>
				</td>
				<td class="label" style="width:10%;">lingue attive:</td>
				<td class="content_right">
					<img src="../grafica/flag_mini_it.jpg" alt="italiano" border="0">
				<% 	if rs("lingua_en") then %>
					<img src="../grafica/flag_mini_en.jpg" alt="inglese" border="0">
				<% 	end if %>
				<% 	if rs("lingua_fr") then %>
					<img src="../grafica/flag_mini_fr.jpg" alt="francese" border="0">
				<% 	end if %>
				<% 	if rs("lingua_es") then %>
					<img src="../grafica/flag_mini_es.jpg" alt="spagnolo" border="0">
				<% 	end if %>
				<% 	if rs("lingua_de") then %>
					<img src="../grafica/flag_mini_de.jpg" alt="tedesco" border="0">
				<% 	end if %>
				<% 	if FieldExists(rs, "lingua_ru") then
						if rs("lingua_ru") then %>
							<img src="../grafica/flag_mini_ru.jpg" alt="russo" border="0">
				<% 		end if 
					end if %>
				<% 	if FieldExists(rs, "lingua_cn") then
						if rs("lingua_cn") then %>
							<img src="../grafica/flag_mini_cn.jpg" alt="cinese" border="0">
				<% 		end if 
					end if %>
				<% 	if FieldExists(rs, "lingua_pt") then
						if rs("lingua_pt") then %>
							<img src="../grafica/flag_mini_pt.jpg" alt="portoghese" border="0">
				<% 		end if 
					end if %>
				</td>
			</tr>
            <tr>
			<%
			dim rows
			rows=1
			if cString(rs("url_secure"))<>"" AND cString(rs("url_alternativo"))<>"" then 
				rows=3
			elseif cString(rs("url_secure"))<>"" OR cString(rs("url_alternativo"))<>"" then
				rows=2
			end if
			%>
                <td class="label_no_width" rowspan="<%=rows%>">url:</td>
                <td class="label_no_width">principale:</td>
                <td class="content" ><a target="_blank" href="<%= rs("url_base") %>"><%= rs("url_base") %></a></td>
				<td class="label_no_width">gestione url:</td>
				<td class="content"><%= IIF(rs("URL_rewriting_attivo"), "<strong>url statici</strong>", "url dinamici") %></td>
            </tr>
            <% if cString(rs("url_secure"))<>"" then %>
                <tr>
                    <td class="label_no_width">sicuro:</td>
					<td class="content" colspan="3"><a target="_blank" href="<%= rs("url_secure") %>"><%= rs("url_secure") %></a></td>
                </tr>
			<% 	end if %>
			<% if cString(rs("url_alternativo"))<>"" then %>
                <tr>
                    <td class="label_no_width">alternativo:</td>
					<td class="content" colspan="3"><a target="_blank" href="<%= rs("url_alternativo") %>"><%= rs("url_alternativo") %></a></td>
                </tr>
			<% 	end if %>
			<tr>
				<td class="label_no_width" colspan="2">tipo di sito:</td>
				<% if rs("sito_mobile") then %>
					<td class="label_no_width" colspan="3">sito per dispositivi mobili</td>
				<% else %>
					<td class="label_no_width" colspan="3">normale</td>
				<% end if %>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">stato del sito:</td>
				<% if rs("sito_in_aggiornamento") then %>
					<td class="content inaggiornamento">sito in aggiornamento</td>
				<% elseif rs("sito_in_costruzione") then %>	
					<td class="content incostruzione">sito in costruzione</td>
				<% else %>
					<td class="content attivo">sito attivo</td>
				<% end if %>
				<td class="label_no_width">id:</td>
				<td class="content"><%= rs("id_webs") %></td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">accessibilit&agrave;:</td>
                <td class="content" colspan="3"><%= IIF(rs("sito_accessibile"), "sito accessibile", "sito non accessibile") %></td>
            </tr>
			<tr>
				<td class="label_no_width" colspan="2">indicizzazione:</td>
                <td class="content" colspan="3"><%= IIF(rs("sito_indicizzabile"), "sito indicizzabile", "sito non indicizzabile") %></td>
            </tr>
            <% if index.ChkPrm(prm_strumenti_accesso, 0) then %>
			<tr>
				<td colspan="4" class="label_no_width">
					Strumenti di analisi e gestione del sito
				</td>
				<td class="content_center">
					<a class="button_block" href="SitoAnalisi.asp?ID=<%= rs("id_webs") %>">
						STRUMENTI
					</a>
				</td>
			</tr>
			<% 	end if %>
			<% 	if index.ChkPrm(prm_menu_accesso, 0) then %>
			<tr>
				<td colspan="4" class="label_no_width">
					Creazione e gestione dei menu interattivi del sito
				</td>
				<td class="content_center">
					<a class="button_block" href="SitoMenu.asp?ID=<%= rs("id_webs") %>">
						MENU
					</a>
				</td>
			</tr>
			<%	end if %>
			<tr>
				<td colspan="4" class="label_no_width" style="width:80%;">
					Gestione dei files trasferiti online
				</td>
				<td class="content_center">
					<a class="button_block" href="SitoFileManager.asp?ID=<%= rs("id_webs") %>">
						FILES
					</a>
				</td>
			</tr>
			<% 	if IsNextAim() then %>
			<tr>
				<td colspan="4" class="label_no_width">
					Gestione delle propriet&agrave; dei plugin dinamici utilizzati
				</td>
				<td class="content_center">
					<a class="button_block" href="SitoPlugin.asp?ID=<%= rs("id_webs") %>">
						PLUGIN
					</a>
				</td>
			</tr>
			<% 	end if %>
			<% 	if index.ChkPrm(prm_template_accesso, 0) then %>
			<tr>
				<td colspan="4" class="label_no_width">
					Gestione dei templates di pagina
				</td>
				<td class="content_center">
					<a class="button_block" href="SitoTemplate.asp?ID=<%= rs("id_webs") %>">
						TEMPLATES
					</a>
				</td>
			</tr>
			<% 	end if %>
			<tr>
				<td colspan="4" class="label_no_width">
					Gestione delle pagine del sito
				</td>
				<td class="content_center">
					<a class="button_block" href="SitoPagine.asp?ID=<%= rs("id_webs") %>">
						PAGINE
					</a>
				</td>
			</tr>
			<% 	if index.ChkPrm(prm_stili_accesso, 0) then %>
			<tr>
				<td colspan="4" class="label_no_width">
					Gestione stili di testo per la composizione delle pagine.
				</td>
				<td class="content_center">
					<a class="button_block" href="SitoStili.asp?ID=<%= rs("id_webs") %>">
						STILI
					</a>
				</td>
			</tr>
			<% 	end if %>
		</table>
			<%rs.movenext
		wend
	else%>
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Gestione siti</caption>
			<tr><td class="noRecords">Nessun record trovato</th></tr>
		</table>
	<% end if %>
</div>
</body>
</html>
<%
rs.close
conn.close
set rs = nothing
set conn = nothing%>
