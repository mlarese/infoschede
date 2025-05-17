			<script language="JavaScript" type="text/javascript">
				function RicalcolaTotali(){
					var qta = toNumber(form1.tfn_det_qta.value);
					
					var prezzo_listino = toNumber(form1.tfn_det_prezzo_listino.value);
					form1.prezzo_listino_totale.value = FormatNumber(prezzo_listino * qta, 2);
					
					var prezzo_cliente = toNumber(form1.prezzo_cliente.value);
					form1.prezzo_cliente_totale.value = FormatNumber(prezzo_cliente * qta, 2);
					
					var prezzo_unitario = toNumber(form1.tfn_det_prezzo_unitario.value);
					form1.prezzo_unitario_totale.value = FormatNumber(prezzo_unitario * qta, 2);
					
					var iva_inclusa = toNumber(form1.iva_inclusa.value);
					form1.iva_inclusa_totale.value = FormatNumber(iva_inclusa * qta, 2);
				}
				
				
				function VariazioneSconto(){
					//variato lo sconto da listino cliente: ricalcola il prezzo
					CalcolaPrezzo(form1.prezzo_cliente, form1.tfn_det_prezzo_unitario, form1.sconto_cliente);
					VariazionePrezzo();
				}
				
				
				function VariazionePrezzo(){
					//variato il prezzo ricalcola lo sconto dal listino cliente
					CalcolaVariazione(form1.prezzo_cliente, form1.tfn_det_prezzo_unitario, form1.sconto_cliente);
					
					//variato il prezzo: ricalcola lo sconto dal listino base
					CalcolaVariazione(form1.tfn_det_prezzo_listino, form1.tfn_det_prezzo_unitario, form1.tfn_det_sconto);
					
					//ricalcola iva unitaria
					CalcolaPrezzo(form1.tfn_det_prezzo_unitario, form1.iva_inclusa, form1.tfn_det_iva);
					
					var qta = toNumber(form1.tfn_det_qta.value);
					
					//ricalcola prezzo totale netto
					var prezzo = toNumber(form1.tfn_det_prezzo_unitario.value);
					form1.prezzo_unitario_totale.value = FormatNumber(prezzo * qta, 2);
					
					//ricalcola prezzo iva inclusa
					var prezzo = toNumber(form1.iva_inclusa.value);
					form1.iva_inclusa_totale.value = FormatNumber(prezzo * qta, 2);
				}
				
				
				function VariazioneQuantita(){
					var qta = FormatNumber(toNumber(form1.tfn_det_qta.value),0);
					form1.tfn_det_qta.value = qta;
					
					//verifica se la quantita' immessa e' compatibile con il lotto di riordino
					if (qta > 0){
						if (ControllaQta(form1.tfn_det_qta.value, form1.tfn_det_qta.value, <%= CInteger(rs("rel_qta_min_ord")) %>, <%= CInteger(rs("rel_lotto_riordino")) %>, false)) {
							<% if cInteger(rs("prz_scontoQ_id"))>0 then 
								sql = "SELECT * FROM gtb_scontiQ WHERE sco_classe_id = " & rs("prz_scontoQ_id") & " ORDER BY sco_qta_da DESC"
								rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
								if not rsc.eof then%>
									//imposta sconto per quantit&agrave;
									<% while not rsc.eof 
										if rsc.absoluteposition>1 then%>
											else
										<% end if %>
										if (qta >= <%= rsc("sco_qta_da") %>){
											form1.sconto_cliente.value = FormatNumber(<%= Replace(FormatPrice(rsc("sco_sconto"), 2, false), ",", ".") %>, 2);
										}
										<% rsc.movenext
									wend %>
									else {
										form1.sconto_cliente.value = FormatNumber(0,2);
									}
									VariazioneSconto();
								<% end if
								rsc.close
							end if %>
							
							//quantita' conforme: esegue calcoli
							RicalcolaTotali();
						}
					}
					else{
						alert('errore nella quantita\' in ordine');
					}
				}
			</script>
			<tr><th colspan="7">DATI DETTAGLIO</th></tr>
			<tr>
				<td class="label">quantit&agrave;</td>
				<td class="content_right" colspan="3">
					<input type="text" class="number" tabindex="1" name="tfn_det_qta" value="<%= quantita %>" maxlength="10" size="3" onchange="VariazioneQuantita()">
				</td>
				<td class="content" colspan="3">(*)</td>
			</tr>
			<tr>
				<th class="L2" colspan="2">&nbsp;</th>
				<th class="l2_center">unitario</th>
				<th class="l2_center">totale</th>
				<th class="L2" colspan="4">&nbsp;</th>
			</tr>
			<tr>
				<td class="label" rowspan="2">
					prezzi a listino:
				</td>
				<td class="label">
					base attuale:
				</td>
				<td class="content_right">
					<input type="text" readonly class="number_disabled" name="tfn_det_prezzo_listino" value="<%= FormatPrice(prezzo_listino_base, 2, false) %>" size="7">
				</td>
				<td class="content_right">
					<input type="text" disabled class="number_disabled" name="prezzo_listino_totale" value="<%= FormatPrice(prezzo_listino_base * quantita, 2, false) %>" size="9">
				</td>
				<td class="content" colspan="3">&euro;</td>
			</tr>
			<tr>
				<td class="label">
					cliente:
				</td>
				<td class="content_right">
					<input type="text" readonly class="number_disabled" name="prezzo_cliente" value="<%= FormatPrice(prezzo_listino_cliente, 2, false) %>" size="7">
				</td>
				<td class="content_right">
					<input type="text" disabled class="number_disabled" name="prezzo_cliente_totale" value="<%= FormatPrice(prezzo_listino_cliente * quantita, 2, false) %>" size="9">
				</td>
				<td class="content" colspan="4">
					&euro;
					<span title="sconto assegnato al cliente da listino base">
						&nbsp;&nbsp;&nbsp;
						<%= FormatPrice(GetVarPercent(prezzo_listino_base, prezzo_listino_cliente), 2, true) %>%
					</span>
					<% if rs("listino_offerte") then %>
						<span title="prodotto in offerta speciale per tutti i clienti">
							<span class="Icona Offerte">&nbsp;</span>
							&nbsp;in offerta speciale
						</span>
					<% elseif rs("prz_promozione") then %>
						<span title="prodotto in promozione per il cliente">
							<span class="Icona Promozioni">&nbsp;</span>
							&nbsp;in promozione
						</span>
					<% else %>
						&nbsp;
					<% end if %>
				</td>
			</tr>
			<tr>
				<td class="label" rowspan="2">
					variazione prezzo da:
				</td>
				<td class="label">
					listino base:
				</td>
				<td class="content_right">&nbsp;</td>
				<td class="content_right">
					<input type="text" readonly class="number_disabled" name="tfn_det_sconto" value="<%= FormatPrice(sconto_listino_base, 2, false) %>" size="4">
				</td>
				<td class="content" colspan="4">%</td>
			</tr>
			<tr>
				<td class="label">
					listino cliente:
				</td>
				<td class="content_right">&nbsp;</td>
				<td class="content_right">
					<input type="text" class="number" name="sconto_cliente" value="<%= FormatPrice(sconto_listino_cliente, 2, false) %>" size="4" onchange="VariazioneSconto()">
				</td>
				<td class="content" colspan="4">%</td>
			</tr>
			<tr>
				<td class="label" rowspan="2">
					prezzo finale:
				</td>
				<td class="label">
					netto:
				</td>
				<td class="content_right">
					<input type="text" class="number" name="tfn_det_prezzo_unitario" value="<%= FormatPrice(prezzo_finale, 2, false) %>" size="7" onchange="VariazionePrezzo()">
				</td>
				<td class="content_right">
					<input type="text" disabled class="number_disabled" name="prezzo_unitario_totale" value="<%= FormatPrice(prezzo_finale * quantita, 2, false) %>" size="9">
				</td>
				<td class="content" colspan="4">
					&euro;
					&nbsp;&nbsp;&nbsp;
					<% if rs("iva_valore")>0 then %>
						+ <%= rs("iva_valore") %>% i.v.a.
					<% else %>
						esente i.v.a.
					<% end if %>
				</td>
			</tr>
			<input type="hidden" name="tfn_det_iva" value="<%= rs("iva_valore") %>">
			<tr>
				<td class="label">
					iva inclusa:
				</td>
				<td class="content_right">
					<input type="text" class="number_disabled" disabled name="iva_inclusa" value="<%= FormatPrice(prezzo_finale + GetIva(prezzo_finale, rs("iva_valore")), 2, false) %>" size="7">
				</td>
				<td class="content_right">
					<input type="text" class="number_disabled" disabled name="iva_inclusa_totale" value="<%= FormatPrice(prezzo_finale * quantita + GetIva(prezzo_finale * quantita, rs("iva_valore")), 2, false) %>" size="9">
				</td>
				<td class="content" colspan="4">&euro;</td>
			</tr>