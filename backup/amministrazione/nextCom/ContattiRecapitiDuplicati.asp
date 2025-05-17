<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_contatti.ASP" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<%
'--------------------------------------------------------
sezione_testata = "Elenco contatti con lo stesso recapito"  %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim conn, rs, rsr, rsv, sql, rubriche_visibili, var

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")

'recupera rubriche visibili all'utente
rubriche_visibili = GetList_Rubriche(conn, rs)

sql = " SELECT IDElencoIndirizzi, Lingua, lingua_nome_IT, LockedByApplication, ApplicationsLocker, SyncroApplication, " +_
      " isSocieta, CognomeElencoIndirizzi, NomeElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi, TitoloElencoIndirizzi, " + _
	  " indirizzoElencoIndirizzi, capElencoIndirizzi, cittaElencoIndirizzi,localitaElencoIndirizzi, statoProvElencoIndirizzi, countryElencoIndirizzi, DataIscrizione " + _
	  " FROM tb_indirizzario INNER JOIN tb_cnt_lingue ON tb_indirizzario.lingua=tb_cnt_lingue.lingua_codice WHERE " & _
	  " IDElencoIndirizzi IN (SELECT id_indirizzo FROM rel_rub_ind WHERE id_rubrica IN (" & IIF(rubriche_visibili<>"", rubriche_visibili, "0") & ")) " + _
	  " AND IdElencoIndirizzi<>" & cInteger(request("contatto")) & _
	  " AND IdElencoIndirizzi IN (SELECT ID_Indirizzario FROM tb_valoriNumeri WHERE LTRIM(RTRIM(ValoreNumero)) LIKE '" & ParseSql(Trim(request("recapito")), adChar) & "%' )"	+ _
	  " ORDER BY ModoRegistra "
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
%>
<div id="content_ridotto">
<table cellspacing="1" cellpadding="0" class="tabella_madre">
	<caption>Elenco contatti con recapito "<%= request("recapito") %>"</caption>
	<% if not rs.eof then %>
		<tr><th>Trovati n&ordm; <%= rs.recordcount %> contatti</th></tr>
		<%while not rs.eof %>
			<tr>
				<td class="body">
					<table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
						<tr>
							<td class="header" colspan="2">
								<table border="0" cellspacing="0" cellpadding="0" align="right">
									<tr>
										<% if rs("lingua")<>"" then %>
											<td style="padding-right:4px; vertical-align:bottom;">
												<img src="../grafica/flag_mini_<%= rs("lingua") %>.jpg" alt="Lingua: <%= rs("lingua_nome_IT") %>">
											</td>
										<% end if %>
										<td style="font-size: 1px;">
											<a class="button" target="contatto_<%= rs("IdElencoIndirizzi") %>" href="ContattiMod.asp?ID=<%= rs("IDElencoIndirizzi") %>" title="modifica dati anagrafici e rubriche">
												MODIFICA
											</a>
											&nbsp;
											<a class="button" target="contatto_<%= rs("IdElencoIndirizzi") %>" href="ContattiRecapiti.asp?ID=<%= rs("IDElencoIndirizzi") %>" title="gestione recapiti (telefono, fax, email, ecc..)">
												RECAPITI
											</a>
											&nbsp;
											<% if cInteger(rs("LockedByApplication"))>0 then
												sql = "SELECT sito_nome FROM tb_siti WHERE id_sito IN (" & rs("ApplicationsLocker") & "0 )"%>
												<a class="button_disabled" title="contatto non cancellabile perch&egrave; bloccato dalle applicazioni: <%= GetValueList(conn, rsr, sql) %>.">
													CANCELLA
												</a>
											<% elseif cInteger(rs("SyncroApplication"))>0 then
												sql = "SELECT sito_nome FROM tb_siti WHERE id_sito=" & rs("SyncroApplication")%>
												<a class="button_disabled" title="contatto gestito completamente dall'applicazione: <%= GetValueList(conn, rsr, sql) %>.">
													CANCELLA
												</a>
											<% else
												if Application("NextCrm") then 
													sql = "SELECT (COUNT(*)) AS N_PRATICHE FROM tb_pratiche WHERE pra_cliente_id=" & rs("IDElencoIndirizzi")
													rsv.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
													var = rsv("N_PRATICHE")>0
													rsv.close
												else
													var = 0
												end if
												if var= 0 then %>
													<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('CONTATTI','<%= rs("IDElencoIndirizzi") %>');">
														CANCELLA
													</a>
												<% else %>
													<a class="button_disabled" title="contatto non cancellabile perch&egrave; associato a delle pratiche.">
														CANCELLA
													</a>
												<% end if
											end if %>
										</td>
									</tr>
								</table>
								<%=ContactName(rs)%><%= IIF(cString(rs("TitoloElencoIndirizzi"))<>"", ",&nbsp;" & rs("TitoloElencoIndirizzi"), "")%>
							</td>
						</tr>
						<% if rs("isSocieta") then 
							if rs("CognomeElencoIndirizzi") & rs("NomeElencoIndirizzi")<>"" then  %>
								<tr>
									<td class="label">contatto:</td>
									<td class="content"><%= rs("CognomeElencoIndirizzi") %>&nbsp;<%= rs("NomeElencoIndirizzi") %><%= IIF(cString(rs("TitoloElencoIndirizzi"))<>"", ",&nbsp;" & rs("TitoloElencoIndirizzi"), "") %></td>
								</tr>
							<% end if
						else 
							if rs("NomeOrganizzazioneElencoIndirizzi")<>"" then  %>
								<tr>
									<td class="label">ente:</td>
									<td class="content"><%= rs("NomeOrganizzazioneElencoIndirizzi") %></td>
								</tr>
							<% end if
						end if %>
						<tr>
							<td class="label" style="width:19%;">rubriche:</td>
							<% sql = " SELECT '<span style=""white-space:nowrap"">' + nome_rubrica + '</span>' FROM tb_rubriche " &_
									 " INNER JOIN rel_rub_ind ON tb_rubriche.id_rubrica=rel_rub_ind.id_rubrica " &_
									 " WHERE rel_rub_ind.id_indirizzo=" & rs("IDElencoIndirizzi") %>
							<td class="content"><%= GetValueList(conn, rsr, sql) %></td>
						</tr>
						<tr>
							<td class="label">indirizzo:</td>
							<td class="content">
								<%= ContactAddress(rs) %>
							</td>
						</tr>
						<% sql = "SELECT * FROM tb_TipNumeri WHERE id_tipoNumero " &_
								 " IN (SELECT id_TipoNumero FROM tb_ValoriNumeri WHERE id_indirizzario=" & rs("IDElencoIndirizzi") & ") "
						rsv.Open sql, conn, AdOpenForwardOnly, adLockReadOnly, adCmdText
						while not rsv.eof
							sql = "SELECT * FROM tb_ValoriNumeri " &_
								  " WHERE id_TipoNumero=" & rsv("id_tipoNumero") & " AND  id_Indirizzario=" & rs("IDElencoIndirizzi")
							rsr.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
							if not rsr.eof then%>
								<tr>
									<td class="label" nowrap><%= Lcase(rsv("nome_TipoNumero")) %>:</td>
									<td class="content<%= IIF(cString(rsr("ValoreNumero")) = cString(request("recapito")), " warning", "") %>">
										<% if cString(rsr("ValoreNumero")) = cString(request("recapito")) then %>
											<table align="right" cellpadding="0" cellspacing="0">
												<tr>
													
													<td>
														<% if cString(rsr("SyncroField"))="" OR rsr("id_TipoNumero")=VAL_EMAIL then%>
															<a class="button_L2" target="contatto_<%= rs("IdElencoIndirizzi") %>" href="ContattiRecapiti.asp?ID=<%= rs("IDElencoIndirizzi") %>&RID=<%= rsr("id_ValoreNumero") %>" title="gestione recapiti (telefono, fax, email, ecc..)">
																MODIFICA
															</a>
														<% else %>
															<a class="button_L2_disabled" title="contatto sincronizzato con dei dati di un'applicazione esterna.">
																MODIFICA
															</a>
														<% end if %>
														&nbsp;
														<% if cString(rsr("SyncroField"))<>"" then%>
															<a class="button_L2_disabled" title="contatto sincronizzato con dei dati di un'applicazione esterna.">
																CANCELLA
															</a>
														<% else %>
															<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('RECAPITI','<%= rsr("id_ValoreNumero") %>');" >
																CANCELLA
															</a>
														<% end if %>
													</td>
												</tr>
											</table>
										<% end if
										while not rsr.eof
											select case rsr("id_TipoNumero")
												case 6	'email %>
													<a href="mailto:<%= rsr("ValoreNumero") %>"><%= rsr("ValoreNumero") %></a>
												<% case 7	'web %>
													<a href="http://<%= rsr("ValoreNumero") %>" target="_blank"><%= rsr("ValoreNumero") %></a>
												<% case else %>
													<%= rsr("ValoreNumero") %>
											<%end select
											rsr.movenext
											if not rsr.eof then%>
												,&nbsp;
											<%end if
										wend %>
									</td>
								</tr>
							<%end if
							rsr.close
							rsv.MoveNext
						wend
						rsv.Close
						if isDate(rs("DataIscrizione")) then%>
							<tr>
								<td class="label">data iscrizione:</td>
								<td class="content">
									<%= DateTimeIta(rs("DataIscrizione")) %>
								</td>
							</tr>
						<% 	end if
						%>
					</table>
				</td>
			</tr>
			<% rs.moveNext
		wend
	else%>
		<tr><td class="noRecords">Nessun record trovato</th></tr>
	<% end if %>
	<tr>
		<td class="footer" colspan="6">
			Tutti i campi sono obbligatori.
			<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
		</td>
	</tr>
</table>