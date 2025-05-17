<%
class testata
	
	'variabili e proprieta' locali
	Public sottosezioni			'array di sottosezioni	 
	Public links 				'array di campi visibili nell'elenco
	Public sezione 				'dicitura
	Public fields_intervento    'array di campi di intervento accessibili  dall'elenco
	Public puls_new
	Public link_new
	Public puls_2a_riga
	
	'variabili per indicizzazione
	Private Index_object
	Private Index_F_Table
	Private Index_F_ID
	
	Public sub iniz_sottosez(ByVal qfel)
		sottosezioni = Array(qfel)
		links = Array(qfel)
		redim preserve sottosezioni(qfel)
		redim preserve links(qfel)
		set puls_2a_riga = Server.CreateObject("Scripting.Dictionary")
	end sub
	
	Public Sub scrivi_con_sottosez()
		scrivi_con_sottosez_advanced(false)		
	end sub
	
	Public Sub scrivi_con_sottosez_advanced(ridotta)
		CALL scrivi_con_sottosez_advanced_extended(ridotta, "", true)
	end sub
	
	'titoloSezioneEsterna (string) 		se diverso da stringa vuota attiva la sezione esterna a destra con il link al contenuto
	'modificaIndice (bool)				se false nasconde il pulsante COLLEGATO ALL'INDICE
	Public Sub scrivi_con_sottosez_advanced_extended(ridotta, titoloSezioneEsterna, modificaIndice)
		dim puls_list, link_list, a, b, i
		response.write "<tr><td><table class=""testata"" width='100%' cellspacing='0' cellpadding='0' border='0'><tr>"
		response.write "<td><table cellspacing='0' cellpadding='0' border='0'>"
		response.write "<tr>"		
		if Session("ERRORE") <> "" then
			response.write "<td style=""padding-left:15px;"">"
			response.write "<font class=""errore"">" & Session("ERRORE") & "</font>"
			response.write "</td>"
			Session("ERRORE") = ""
		else
			if sezione <> "" then
				if not ridotta then
					response.write "<td><font class='testo10b'>&nbsp;&nbsp;" + _
						ChooseValueByAllLanguages(Session("LINGUA"), "sezioni:", "section:", "", "", "", "", "", "") + _
						"&nbsp;</font>"
				else
					response.write "<td><font class='testo10b'>sez.:</font>"
				end if
			end if
			response.write "<font class='testo12b'>"+sezione+"&nbsp;&nbsp;</font>"
			if ubound(sottosezioni)>0 then
				if not ridotta then
					response.write "<font class='testo10b'>&nbsp;&nbsp;" + _
					ChooseValueByAllLanguages(Session("LINGUA"), "sottosezioni:", "subsection:", "", "", "", "", "", "") + _
					"&nbsp;</font>"
				else
					response.write "<font class='testo10b'>&nbsp;" + _
					ChooseValueByAllLanguages(Session("LINGUA"), "sottosez.:", "subsect.:", "", "", "", "", "", "") + _
					"</font>"
				end if
				for a = 1 to ubound(sottosezioni)
					response.write "<a href="""+links(a)+""" class='links1'>"+sottosezioni(a)+"</a>"
					if a < ubound(sottosezioni) then
						if not ridotta then
							response.write"<font class='testo10b'>&nbsp;&nbsp;&nbsp;</font>"
						else
							response.write"<font class='testo10b'>&nbsp;</font>"
						end if
					end if
				next
			end if
			response.write "</td>" 'Giacomo 20/11/2012
		end if
		response.write "</tr></table></td>"
		response.write "<td align='right'><table cellspacing='0' cellpadding='0' border='0'><tr>"
		if puls_new <> "" then
			puls_list = split(puls_new, ";")
			link_list = split(link_new, ";")
			for a = 0 to ubound(puls_list)
				if Trim(puls_list(a)) <> "" then
					if link_list(a)<>"" then
						'scrive pulsante con link
						response.write "<td style='padding-left:3px;' nowrap><a href=""" & link_list(a) & """ class='simul_puls_1'>" & puls_list(a) & "</a></td>"
					else
						'scrive label
						response.write "<td style='padding-left:10px;' nowrap>" & puls_list(a) & "</td>"
					end if
				end if
			next
		else
			response.write "<td><font class='testo10b'>&nbsp;</font></td>"
		end if

		if IsObject(Index_object) AND cString(titoloSezioneEsterna)="" then
			CALL Index_object.WriteButton(Index_F_Table, Index_F_ID, POS_TESTATA)
		end if
		
		if puls_2a_riga.Count > 0 then %>
		
			</tr>
		</table>
		<table cellspacing="0" cellpadding="0" border="0" width="100%">
			<tr>
				<td style="text-align:right; padding-top:3px;">
					<%
					a = puls_2a_riga.Items
					b = puls_2a_riga.Keys
					for i=0 to puls_2a_riga.Count-1 %>
						<a class="button_L2" href="<%=a(i)%>" style="margin-left:4px;">
						<%=  b(i) %>
						</a>
					<% next %>
				</td>
		<% end if
			
		response.write "</tr>"
		response.write "</table></td></tr></table></td></tr>"
		response.write "</table></div>"

		if IsObject(Index_object) AND cString(titoloSezioneEsterna)<>"" then
			%>
			<div id="pulsanti" style="position:absolute; top:93px; left:810px; width:210px; text-align:center;">
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption class="">Operazioni</caption>
					<%
					dim rs, sql
					set rs = server.CreateObject("ADODB.recordset")
					sql = " SELECT co_id FROM tb_contents " & _
						  " WHERE co_F_table_id IN (SELECT tab_id FROM tb_siti_tabelle WHERE tab_name='"&Index_F_Table&"') AND co_f_key_id="&Index_F_ID
					rs.open sql, Index_object.conn
					%>
					<% if cBoolean(modificaIndice, false) then %>
						<tr><th style="text-align:center;">MODIFICA COLLEGAMENTO ALL'INDICE	</th></tr>
						<tr>
							<td>
								<table cellspacing="1" cellpadding="0" class="" style="background-color:#f4f4f2; width:100%; text-align:center;">
									<tr><td style="line-height:2px;">&nbsp;</td></tr>
									<tr>
									<%
									if not rs.eof then
										%>
										<td>
										<%
										CALL Index_object.WriteButton(Index_F_Table, Index_F_ID, POS_ELENCO)
										%>
										</td>
										<%
									else
										%>
										<td class="label">Contenuto non pubblicabile.</td>
										<%
									end if
									%>
									</tr>
									<tr><td style="line-height:3px;">&nbsp;</td></tr>
								</table>
							</td>
						</tr>
					<% end if 
					rs.close
					
					
					dim lingua
					sql = ""
					for each lingua in Application("LINGUE")
						sql = sql & " idx_link_url_"&lingua&", idx_link_url_rw_"&lingua&", co_titolo_"&lingua&", "
					next
					set rs = server.CreateObject("ADODB.recordset")
					sql = " SELECT "&sql&" idx_id, idx_link_tipo, idx_webs_id FROM v_indice " & _
						  " WHERE co_F_table_id IN (SELECT tab_id FROM tb_siti_tabelle WHERE tab_name='"&Index_F_Table&"') AND co_f_key_id="&Index_F_ID

					rs.open sql, Index_object.conn
					if not rs.eof then
						%>
						<tr><th style="text-align:center;"><%= titoloSezioneEsterna%></th></tr>
						<tr>
							<td>
								<table cellpadding="0" cellspacing="1" style="width:100%;">
									<%
									for each lingua in Application("LINGUE")
									%>
									<tr>
										<td class="content" style="vertical-align:middle; padding-right:6px; width:18px;">
											<img src="../grafica/flag_mini_<%= lingua %>.jpg" border="0">
										</td>
										<td class="content" style="vertical-align:top; padding-top:3px; padding-bottom:4px;">
											<% CALL Index_object.WriteNodeLabelLink(rs, IIF(rs("co_titolo_"&lingua)<>"",rs("co_titolo_"&lingua),rs("co_titolo_it")) , "", "", lingua) %>
										</td>
									</tr>
									<%
									next
									%>
								</table>
							</td>
						</tr>
						<%
					end if
					rs.close
					%>
				</table>
			</div>
			<%
		end if		
	end sub
	
	
	
	public Sub scrivi_ridotta()%>
		<table width="400" cellspacing="0" cellpadding="0" border="0">
			<tr>
				<td bgcolor="#E6E6E6" align="right" style="padding-top:1px; padding-right:5px; border-bottom:1px solid #FFF;">
					<a href="javascript:window.close();" class="menu">
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "CHIUDI", "CLOSE", "", "", "", "", "", "")%>
					</a>
				</td>
			</tr>
			<tr>
				<td>
					<img src="<%= GetAmministrazionePath() %>grafica/<%= Application("PREFISSO_BARRA_AMMINISTRAZIONE_PERSONALIZZATA") %>barra_intestazione_rid.jpg" width="400" height="47" alt="" border="0">
				</td>
			</tr>
			<tr>
				<td class="testo10b">
                    <% if Session("ERRORE")<>"" then %>
                        <span class="errore"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Errore: ", "Error: ", "", "", "", "", "", "")%><%= Session("ERRORE") %></span>
                        <% Session("ERRORE") = ""
                    else 
						if sezione <> "" then %>
							&nbsp;&nbsp;<%= ChooseValueByAllLanguages(Session("LINGUA"), "sezione:", "section:", "", "", "", "", "", "")%>&nbsp;
							<span class="testo12b">
								<%= sezione %>
							</span>
						<% end if %>
                    <% end if %>
				</td>
			</tr>
		</table>
	<%end sub
	
	
	public Sub InitializeIndex(oIndex, Table, ID)
		if IsObject(oIndex) then
			if Table<>"" then
				if cInteger(ID)>0 then
					set Index_object = oIndex
					Index_F_Table = Table
					Index_F_ID = ID
				end if
			end if
		end if
	end sub
end class
 %>
