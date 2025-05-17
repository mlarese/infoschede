<%@ Language=VBScript CODEPAGE=65001 %>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../nextCom/Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/classIndirizzarioLock.asp" -->
<!--#INCLUDE FILE="../library/ExportTools.asp" -->

<% 
dim conn
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")

if cIntero(request("ARCHIVIA_SC_ID")) > 0 then
	conn.execute("UPDATE sgtb_schede SET sc_stato_id = " & StatoSchedaConclusa & " WHERE sc_id = " & cIntero(request("ARCHIVIA_SC_ID")))
	response.redirect "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
end if


dim dicitura, assegnata, is_officina, sql_export

set dicitura = New testata
if request("ASSEGNATA") = "false" then
	assegnata = false
else
	assegnata = true
end if

if cString(Session("INFOSCHEDE_OFFICINA"))<>"" then
	is_officina = true
else
	is_officina = false
end if


id_centro_assistenza = GetIdCentroAssistenzaLoggato()
if id_centro_assistenza > 0 then
	sql = " AND sc_centro_assistenza_id = " & id_centro_assistenza
end if


dicitura.iniz_sottosez(0)
if assegnata then
	dicitura.sezione = "Gestione schede di assistenza - elenco"
	dicitura.puls_new = ""
	dicitura.link_new = ""
else
	dicitura.sezione = "Gestione richieste di assistenza - elenco"
	dicitura.puls_new = "INSERISCI RICHIESTA DI ASSISTENZA"
	dicitura.link_new = "SchedeNew.asp"
end if

if id_centro_assistenza > 0 AND not is_officina then
	dicitura.sezione = "Gestione schede di assistenza - elenco"
	dicitura.puls_new = "INSERISCI SCHEDA DI ASSISTENZA"
	dicitura.link_new = "SchedeNew.asp?ID_CENTRO_ASSISTENZA="&id_centro_assistenza
end if

dicitura.scrivi_con_sottosez() 

dim rs, last, Pager, sql, rsa, rows, color, dataStart, id_centro_assistenza, cnt_id
set Pager = new PageNavigator

set rs = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")


'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("sch_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("sch_")
	end if
end if

'ricerca per numero scheda
if cIntero(Session("sch_numero"))<>0 then
	sql = sql & " AND sc_numero = " & Session("sch_numero")
end if

'ricerca per garanzia
if Session("sch_garanzia")<>"" then
	if not (instr(1, Session("sch_garanzia"), "1", vbTextCompare)>0 AND _
		    instr(1, Session("sch_garanzia"), "0", vbTextCompare)>0 ) then
		sql = sql & " AND "
		if instr(1, Session("sch_garanzia"), "1", vbTextCompare)>0 then
			'in garanzia
			sql = sql & " ISNULL(sc_in_garanzia, 0)=1 "
		elseif instr(1, Session("sch_garanzia"), "0", vbTextCompare)>0 then
			'fuori garanzia
			sql = sql & " ISNULL(sc_in_garanzia, 0)=0 "
		end if
	end if
end if

'ricerca per richiesta garanzia
if Session("sch_richiesta_garanzia")<>"" then
	if not (instr(1, Session("sch_richiesta_garanzia"), "1", vbTextCompare)>0 AND _
		    instr(1, Session("sch_richiesta_garanzia"), "0", vbTextCompare)>0 ) then
		sql = sql & " AND "
		if instr(1, Session("sch_richiesta_garanzia"), "1", vbTextCompare)>0 then
			'in garanzia
			sql = sql & " ISNULL(sc_richiesta_garanzia, 0)=1 "
		elseif instr(1, Session("sch_richiesta_garanzia"), "0", vbTextCompare)>0 then
			'fuori garanzia
			sql = sql & " ISNULL(sc_richiesta_garanzia, 0)=0 "
		end if
	end if
end if

' di default mostro le schede di tutti gli stati tranne quelle concluse
if cString(Session("sch_stati")) = "" then
	rsa.open "SELECT sts_id FROM sgtb_stati_schede WHERE sts_id NOT IN ("&StatoSchedaConclusa&") ORDER BY sts_ordine", conn, adOpenForwardOnly, adLockOptimistic, adCmdText
	while not rsa.eof
		Session("sch_stati") = Session("sch_stati") & rsa("sts_id") & ","
		rsa.moveNext
	wend
	Session("sch_stati") = Left(Session("sch_stati"), len(Session("sch_stati")) - 1)
	rsa.close
end if

'ricerca per stato scheda
'if Session("sch_stato")<>"" then
'	sql = sql & " AND sc_stato_id = " & Session("sch_stato")
'end if
if Session("sch_stati")<>"" then
	sql = sql & " AND sc_stato_id IN (" & Session("sch_stati") & ")"
end if

'filtra per centro assistenza
if Session("sch_centro_assistenza")<>"" AND assegnata AND id_centro_assistenza = 0 then
    sql = sql & " AND sc_centro_assistenza_id IN (SELECT ag_id FROM gv_agenti WHERE " & _
				SQL_FullTextSearch_Contatto_Nominativo(conn, Session("sch_centro_assistenza")) & ")"
end if

'filtra per nome cliente
if Session("sch_cliente")<>"" then
	sql = sql & " AND sc_cliente_id IN (SELECT riv_id FROM gv_rivenditori WHERE " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("sch_cliente")) & ")" & _
				" OR ("& SQL_FullTextSearch(Session("sch_cliente"), "sc_rif_cliente") &")"
end if

'filtra per riv_id (pagina chiamata dall'elenco anagrafiche)
if request("sch_riv_id")<>"" then
	Session("sch_riv_id") = request("sch_riv_id")
	sql = sql & " AND sc_cliente_id IN (" & Session("sch_riv_id") & ") "
else
	Session("sch_riv_id") = ""
end if

'filtra per articolo variante
if Session("sch_articolo")<>"" then
	sql = sql & " AND sc_modello_id = " & Session("sch_articolo")
end if

'filtra per costruttore
if Session("sch_costruttore")<>"" then
	sql = sql & " AND sc_modello_id IN (SELECT rel_id FROM grel_art_valori WHERE rel_art_id IN (SELECT art_id FROM gtb_articoli WHERE " & _
				"		art_marca_id IN (SELECT mar_id FROM gtb_marche WHERE mar_anagrafica_id = "&cIntero(Session("sch_costruttore")) & "))) "
end if

'filtra per data ricevimento scheda
if isDate(Session("sch_data_ricevimento_from")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "sc_data_ricevimento", adCompareGreaterThan, Session("sch_data_ricevimento_from"))
end if
if isDate(Session("sch_data_ricevimento_to")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "sc_data_ricevimento", adCompareLessThan, Session("sch_data_ricevimento_to"))
end if

'filtra per data fine lavoro
if isDate(Session("sch_data_fine_from")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "sc_data_fine_lavoro", adCompareGreaterThan, Session("sch_data_fine_from"))
end if
if isDate(Session("sch_data_fine_to")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "sc_data_fine_lavoro", adCompareLessThan, Session("sch_data_fine_to"))
end if

'filtri per DDT di presa
if Session("sch_numero_DDT_presa")<>"" then
	sql = sql & " AND sc_numero_DDT_di_carico = " & Session("sch_numero_DDT_presa")
end if
if isDate(Session("sch_data_presa_from")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "sc_data_data_DDT_di_carico", adCompareGreaterThan, Session("sch_data_presa_from"))
end if
if isDate(Session("sch_data_presa_to")) then
	sql = sql & " AND "& SQL_CompareDateTime(conn, "sc_data_data_DDT_di_carico", adCompareLessThan, Session("sch_data_presa_to"))
end if

'filtri per DDT di riconsegna
if Session("sch_numero_DDT_riconsegna")<>"" then
	sql = sql & " AND sc_rif_DDT_di_resa_id IN (SELECT ddt_id FROM sgtb_ddt WHERE ddt_numero = " & Session("sch_numero_DDT_riconsegna") & ")"
end if
if isDate(Session("sch_data_riconsegna_from")) then
	sql = sql & " AND sc_rif_DDT_di_resa_id IN (SELECT ddt_id FROM sgtb_ddt WHERE " & _
				SQL_CompareDateTime(conn, "ddt_data", adCompareGreaterThan, Session("sch_data_riconsegna_from")) & ")"
end if
if isDate(Session("sch_data_riconsegna_to")) then
	sql = sql & " AND sc_rif_DDT_di_resa_id IN (SELECT ddt_id FROM sgtb_ddt WHERE " & _
				SQL_CompareDateTime(conn, "ddt_data", adCompareLessThan, Session("sch_data_riconsegna_to")) & ")"
end if


if assegnata then
	sql = " SELECT * FROM sgtb_schede WHERE ISNULL(sc_centro_assistenza_id, 0)>0 " + sql + " ORDER BY sc_numero DESC"
	Session("INFOSCHEDE_SCHEDE_SQL") = sql
else
	sql = " SELECT * FROM sgtb_schede WHERE ISNULL(sc_centro_assistenza_id, 0)=0 " + sql + " ORDER BY sc_numero DESC"
	Session("INFOSCHEDE_SCHEDE_SQL") = sql
end if

sql_export = "SELECT [sc_id] " & Right(sql,Len(sql)-InStr(sql,"FROM")+1)
sql_export = Left(sql_export,Len(sql_export)-(Len(sql_export)-InStr(sql_export,"ORDER BY") +1))

CALL Pager.OpenSmartRecordset(conn, rs, sql, 10)

%>
<div id="content">
<% 
'CALL listSession() 
'response.end
%>
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
	  		<td width="27%" valign="top">
			<!-- BLOCCO DI RICERCA -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
					<form action="" method="post" id="ricerca" name="ricerca">
					<tr>
						<td>
							<table cellspacing="1" cellpadding="0" class="tabella_madre">
								<caption>Opzioni di ricerca</caption>
								<tr>
									<td class="footer" colspan="2">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
								<tr><th colspan="2" <%= Search_Bg("sch_numero") %>>NUMERO SCHEDA</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="text" name="search_numero" value="<%= TextEncode(session("sch_numero")) %>" style="width:100%;">
									</td>
								</tr>
								
								<tr><th colspan="2" <%= Search_Bg("sch_garanzia;sch_richiesta_garanzia") %>>GARANZIA</th></tr>
								<tr>
									<td class="content" colspan="2">
										<input type="checkbox" class="checkbox" name="search_garanzia" value="1" <%= chk(instr(1, session("sch_garanzia"), "1", vbTextCompare)>0) %>>
										<b>in garanzia</b>
									</td>
								</tr>
								<tr>
									<td class="content">
										<input type="checkbox" class="checkbox" name="search_garanzia" value="0" <%= chk(instr(1, Session("sch_garanzia"), "0", vbTextCompare)>0) %>>
										fuori garanzia
									</td>
								</tr>
								<tr>
									<td class="content">
										<input type="checkbox" class="checkbox" name="search_richiesta_garanzia" value="1" <%= chk(instr(1, Session("sch_richiesta_garanzia"), "1", vbTextCompare)>0) %>>
										richiesta garanzia
									</td>
								</tr>
								
								<% if assegnata then %>
									<% sql = "SELECT * FROM sgtb_stati_schede ORDER BY sts_ordine" %>
									<tr><th colspan="2" <%= Search_Bg("sch_stato") %>>STATO SCHEDA</th></tr>
									<%
									rsa.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
									while not rsa.eof
										%>
										<tr>
											<td class="content" colspan="2">
												<input type="checkbox" class="checkbox" name="search_stati" value="<%= rsa("sts_id") %>" <%= chk(instr(1, ","&Replace(Session("sch_stati"), " ", "")&"," , ","&rsa("sts_id")&",", vbTextCompare)>0) %>>
												<%= rsa("sts_nome_it") %>
											</td>
										</tr>
										<%
										rsa.moveNext
									wend 
									rsa.close
									%>
									<!--
									<tr>
										<td class="content" colspan="2">
											<% 'CALL dropDown(conn, sql, "sts_id", "sts_nome_it", "search_stato", session("sch_stato"), false, "style=""width: 100%;""", Session("LINGUA")) 
											%>
										</td>
									</tr>
									-->
								<% end if %>
								
								<% if assegnata AND id_centro_assistenza = 0 then %>
									<tr><th colspan="2" <%= Search_Bg("sch_centro_assistenza") %>>CENTRO ASSISTENZA</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_centro_assistenza" value="<%= TextEncode(session("sch_centro_assistenza")) %>" style="width:100%;">
										</td>
									</tr>
								<% end if %>
								
								<tr><th colspan="2" <%= Search_Bg("sch_cliente;sch_riv_id") %>>CLIENTE</th></tr>
								<tr>
									<td class="content" colspan="2">
										<% if Session("sch_riv_id")<>"" then 
											sql = "SELECT riv_profilo_id, ut_abilitato, IDElencoIndirizzi, riv_profilo_id, riv_id, " & _
												  "  NomeElencoIndirizzi, CognomeElencoIndirizzi, isSocieta, NomeOrganizzazioneElencoIndirizzi, IndirizzoElencoIndirizzi, " & _
												  "  LocalitaElencoIndirizzi, CittaElencoIndirizzi, ut_login, cnt_insAdmin_id, cnt_modAdmin_id" & _
												  " FROM gv_rivenditori WHERE riv_id = " & Session("sch_riv_id")
											rsa.open sql, conn
											%>
											<%=ContactName(rsa)%>
											<% rsa.close %>
										<% else %>
											<input type="text" name="search_cliente" value="<%= TextEncode(session("sch_cliente")) %>" style="width:100%;">
										<% end if %>
									</td>
								</tr>

								<tr><th colspan="2" <%= Search_Bg("sch_articolo") %>>MODELLO</th></tr>
								<tr>
									<td class="content" colspan="2">
										<% CALL WritePicker_ArticoloVariante(conn, rsa, "ricerca", "search_articolo", Session("sch_articolo"), 12, false, "Infoschede/ArticoliSeleziona.asp?TYPE=M&") %>
									</td>
								</tr>
								
								<% if assegnata then %>
									<% sql = " SELECT riv_id, IDElencoIndirizzi, NomeElencoIndirizzi, CognomeElencoIndirizzi, isSocieta, NomeOrganizzazioneElencoIndirizzi " & _
											 " FROM gv_rivenditori WHERE (riv_profilo_id IN ("&COSTRUTTORI&")) ORDER BY ModoRegistra" 
									%>
									<tr><th colspan="2" <%= Search_Bg("sch_costruttore") %>>COSTRUTTORE</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL dropDown(conn, sql, "riv_id", "NomeOrganizzazioneElencoIndirizzi", "search_costruttore", session("sch_costruttore"), false, "style=""width: 100%;""", Session("LINGUA")) %>
										</td>
									</tr>
								<% end if %>
								
								<tr><th colspan="2" <%= Search_Bg("sch_data_ricevimento_from;sch_data_ricevimento_to") %>>DATA RICEVIMENTO SCHEDA</th></tr>
								<tr><th class="L2" colspan="2" <%= Search_Bg("sch_data_ricevimento_from") %>>a partire dal:</th></tr>
								<tr>
									<td class="content" colspan="2">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_ricevimento_from", Session("sch_data_ricevimento_from"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr><th class="L2" colspan="2" <%= Search_Bg("sch_data_ricevimento_to") %>>fino al:</th></tr>
								<tr>
									<td class="content" colspan="2">
									<% CALL WriteDataPicker_Input("ricerca", "search_data_ricevimento_to", Session("sch_data_ricevimento_to"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<% if assegnata then %>
									<tr><th colspan="2" <%= Search_Bg("sch_data_fine_from;sch_data_fine_to") %>>DATA FINE LAVORO</th></tr>
									<tr><th class="L2" colspan="2" <%= Search_Bg("sch_data_fine_from") %>>a partire dal:</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_fine_from", Session("sch_data_fine_from"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><th class="L2" colspan="2" <%= Search_Bg("sch_data_fine_to") %>>fino al:</th></tr>
									<tr>
										<td class="content" colspan="2">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_fine_to", Session("sch_data_fine_to"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
								
									<tr><th colspan="2" <%= Search_Bg("sch_numero_DDT_presa;sch_data_presa_from;sch_data_presa_to") %>>DDT di presa</th></tr>
									<tr><th class="L2" colspan="2" <%= Search_Bg("sch_numero_DDT_presa") %>>numero:</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_numero_DDT_presa" value="<%= TextEncode(session("sch_numero_DDT_presa")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th class="L2" colspan="2" <%= Search_Bg("sch_data_presa_from") %>>a partire dal:</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_presa_from", Session("sch_data_presa_from"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><th class="L2" colspan="2" <%= Search_Bg("sch_data_presa_to") %>>fino al:</th></tr>
									<tr>
										<td class="content" colspan="2">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_presa_to", Session("sch_data_presa_to"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									
									<tr><th colspan="2" <%= Search_Bg("sch_numero_DDT_riconsegna;sch_data_riconsegna_from;sch_data_riconsegna_to") %>>DDT di riconsegna</th></tr>
									<tr><th class="L2" colspan="2" <%= Search_Bg("sch_numero_DDT_riconsegna") %>>numero:</th></tr>
									<tr>
										<td class="content" colspan="2">
											<input type="text" name="search_numero_DDT_riconsegna" value="<%= TextEncode(session("sch_numero_DDT_riconsegna")) %>" style="width:100%;">
										</td>
									</tr>
									<tr><th class="L2" colspan="2" <%= Search_Bg("sch_data_riconsegna_from") %>>a partire dal:</th></tr>
									<tr>
										<td class="content" colspan="2">
											<% CALL WriteDataPicker_Input("ricerca", "search_data_riconsegna_from", Session("sch_data_riconsegna_from"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
									<tr><th class="L2" colspan="2" <%= Search_Bg("sch_data_riconsegna_to") %>>fino al:</th></tr>
									<tr>
										<td class="content" colspan="2">
										<% CALL WriteDataPicker_Input("ricerca", "search_data_riconsegna_to", Session("sch_data_riconsegna_to"), "", "/", true, true, LINGUA_ITALIANO) %>
										</td>
									</tr>
								<% end if %>
								<tr>
									<td colspan="2" class="footer">
										<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;">
										<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;">
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><td style="font-size:8px;">&nbsp;</td></tr>
					<tr>
						<td>
							<table cellspacing="1" cellpadding="0" class="tabella_madre">
								<caption class="border">Export dati</caption>
								<tr>
									<td class="content_right">
										<a style="width:100%; text-align:center; line-height:12px;" class="button"
											title="" 
											onclick="OpenAutoPositionedScrollWindow('SchedeExport_perClienti.asp', 'export', 800, 600, true);" href="javascript:void(0);">
											EXPORT PER CLIENTI
										</a>
									</td>
								</tr>
								<tr>
									<td class="content_right">
										<a style="width:100%; text-align:center; line-height:12px;" class="button"
											title="" 
											onclick="OpenAutoPositionedScrollWindow('SchedeExport_perCostruttori.asp', 'export', 800, 600, true);" href="javascript:void(0);">
											EXPORT PER COSTRUTTORI
										</a>
									</td>
								</tr>
								<%					
								sql = " SELECT sc_numero AS [NUMERO], CAST( CONVERT(varchar(10),sc_data_ricevimento,112) AS datetime) AS [DATA RICEVIMENTO], " + vbCrLF + _
									  " (SELECT admin_cognome FROM gv_agenti WHERE ag_id = sgtb_schede.sc_centro_assistenza_id ) AS [CENTRO ASSISTENZA], " + vbCrLF + _
									  " (SELECT (CASE issocieta " + vbCrLF + _
									  " 			WHEN 0 THEN NomeElencoIndirizzi + ' ' + CognomeElencoIndirizzi " + vbCrLF + _
									  " 			ELSE NomeOrganizzazioneElencoIndirizzi " + vbCrLF + _
									  " 		END) FROM gv_rivenditori WHERE riv_id = sgtb_schede.sc_cliente_id) AS [CLIENTE], " + vbCrLF + _
									  " (SELECT StatoProvElencoIndirizzi FROM gv_rivenditori WHERE riv_id = sgtb_schede.sc_cliente_id) AS [PROVINCIA CLIENTE], " + vbCrLF + _
									  " ISNULL(sc_rif_cliente, '') AS [RIFERIMENTO CLIENTE], " + vbCrLF + _
									  " (SELECT art_nome_it FROM gv_articoli WHERE rel_id = sgtb_schede.sc_modello_id) AS [MODELLO], " + vbCrLF + _
									  " (SELECT rel_cod_int FROM grel_art_valori WHERE rel_id = sgtb_schede.sc_modello_id) AS [CODICE MODELLO], " + vbCrLF + _
									  " sc_matricola AS [MATRICOLA MODELLO], sc_negozio_acquisto AS [NEGOZIO DI ACQUISTO], CAST(CONVERT(varchar(10),sc_data_acquisto,112) AS datetime) AS [DATA ACQUISTO], " + vbCrLF + _
									  " sc_numero_scontrino AS [NUMERO SCONTRINO], " + vbCrLF + _
									  " (CASE sc_in_garanzia WHEN 0 THEN 'NO' ELSE 'SI' END) AS [IN GARANZIA], " + vbCrLF + _
									  " (SELECT prb_nome_it FROM sgtb_problemi WHERE prb_id = sgtb_schede.sc_guasto_segnalato_id) AS [GUASTO SEGNALATO], " + vbCrLF + _
									  " (SELECT prb_nome_it FROM sgtb_problemi WHERE prb_id = sgtb_schede.sc_guasto_riscontrato_id) AS [GUASTO RISCONTRATO], " + vbCrLF + _
									  " (SELECT esi_nome_it FROM sgtb_esiti WHERE esi_id = sgtb_schede.sc_esito_intervento_id) AS [ESITO INTERVENTO], " + vbCrLF + _
									  " CAST(CONVERT(varchar(10),sc_data_fine_lavoro,112) AS datetime) AS [DATA FINE LAVORO], " + vbCrLF + _
									  " sc_ora_manodopera_intervento AS [ORE MANODOPERA], " + vbCrLF + _
									  " sc_prezzo_manodopera AS [PREZZO MANODOPERA], " + vbCrLF + _
									  " sc_note_chiusura AS [NOTE DI CHIUSURA], " + vbCrLF + _
									  " sc_numero_DDT_di_carico AS [NUMERO DDT DI CARICO], " + vbCrLF + _
									  " (SELECT ddt_numero FROM sgtb_ddt WHERE ddt_id = sgtb_schede.sc_rif_DDT_di_resa_id) AS [NUMERO DDT DI RICONSEGNA], " + vbCrLF + _
									  " (SELECT CAST(CONVERT(varchar(10),ddt_data,112) AS datetime) FROM sgtb_ddt WHERE ddt_id = sgtb_schede.sc_rif_DDT_di_resa_id) AS [DATA DDT DI RICONSEGNA], " + vbCrLF + _
									  " (SELECT COUNT(*) FROM sgtb_dettagli_schede WHERE dts_scheda_id=sgtb_schede.sc_id) AS [NUMERO RICAMBI UTILIZZATI], " + vbCrLF + _
									  " (SELECT TOP 1 ISNULL(dts_ricambio_codice,'') FROM sgtb_dettagli_schede WHERE dts_scheda_id=sgtb_schede.sc_id) AS [RICAMBIO 1 - CODICE], " + vbCrLF + _
									  " (SELECT TOP 1 ISNULL(dts_ricambio_nome,'') FROM sgtb_dettagli_schede WHERE dts_scheda_id=sgtb_schede.sc_id) AS [RICAMBIO 1 - NOME], " + vbCrLF + _
									  " (SELECT TOP 1 ISNULL(dts_ricambio_codice,'') FROM sgtb_dettagli_schede WHERE dts_scheda_id=sgtb_schede.sc_id  " + vbCrLF + _
									  "			 AND dts_id NOT IN (SELECT TOP 1 dts_id FROM sgtb_dettagli_schede WHERE dts_scheda_id=sgtb_schede.sc_id )) AS [RICAMBIO 2 - CODICE], " + vbCrLF + _
									  " (SELECT TOP 1 ISNULL(dts_ricambio_nome,'') FROM sgtb_dettagli_schede WHERE dts_scheda_id=sgtb_schede.sc_id " + vbCrLF + _
									  "			 AND dts_id NOT IN (SELECT TOP 1 dts_id FROM sgtb_dettagli_schede WHERE dts_scheda_id=sgtb_schede.sc_id )) AS [RICAMBIO 2 - NOME], " + vbCrLF + _
									  " (SELECT TOP 1 ISNULL(dts_ricambio_codice,'') FROM sgtb_dettagli_schede WHERE dts_scheda_id=sgtb_schede.sc_id  " + vbCrLF + _
									  "			 AND dts_id NOT IN (SELECT TOP 2 dts_id FROM sgtb_dettagli_schede WHERE dts_scheda_id=sgtb_schede.sc_id )) AS [RICAMBIO 3 - CODICE], " + vbCrLF + _
									  " (SELECT TOP 1 ISNULL(dts_ricambio_nome,'') FROM sgtb_dettagli_schede WHERE dts_scheda_id=sgtb_schede.sc_id " + vbCrLF + _
									  "			 AND dts_id NOT IN (SELECT TOP 2 dts_id FROM sgtb_dettagli_schede WHERE dts_scheda_id=sgtb_schede.sc_id )) AS [RICAMBIO 3 - NOME] " + vbCrLF + _
									  " FROM sgtb_schede " + vbCrLF + _						  
									  " WHERE sgtb_schede.sc_id IN (" & sql_export & ") " + vbCrLF + _
									  " ORDER BY sc_numero DESC "

								Session("INFOSCHEDE_EXP_PER_SCHEDE") = sql	  
								%>
								<tr>
									<td class="content_center">
										<%
										CALL WRITE_EXPORT_LINK("EXPORT PER SCHEDE", "DATA_ConnectionString", "INFOSCHEDE_EXP_PER_SCHEDE", FORMAT_EXCEL_XML, false)
										%>
									</td>
								</tr>
								<%					
								sql = " SELECT sc_numero AS [NUMERO], CAST( CONVERT(varchar(10),sc_data_ricevimento,112) AS datetime) AS [DATA RICEVIMENTO], " + vbCrLF + _
									  " (SELECT art_nome_it FROM gv_articoli WHERE rel_id = s.sc_modello_id) AS [MODELLO], " + vbCrLF + _
									  " (SELECT rel_cod_int FROM grel_art_valori WHERE rel_id = s.sc_modello_id) AS [CODICE MODELLO], " + vbCrLF + _
									  " sc_matricola AS [MATRICOLA MODELLO], dts_ricambio_codice as [CODICE RICAMBIO], dts_ricambio_nome AS [NOME RICAMBIO] " + vbCrLF + _
									  " FROM sgtb_schede s INNER JOIN sgtb_dettagli_schede d ON s.sc_id = d.dts_scheda_id " + vbCrLF + _						  
									  " WHERE s.sc_id IN (" & sql_export & ") " + vbCrLF + _
									  " ORDER BY sc_numero DESC "

								Session("INFOSCHEDE_EXP_PER_RICAMBI") = sql	  
								%>
								<tr>
									<td class="content_center">
									
										<%
										CALL WRITE_EXPORT_LINK("EXPORT PER RICAMBI", "DATA_ConnectionString", "INFOSCHEDE_EXP_PER_RICAMBI", FORMAT_EXCEL_XML, false)
										%>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					</form>
				</table>
			</td>
			<td width="1%">&nbsp;</td>
			<td valign="top">
<!-- BLOCCO RISULTATI -->
                <table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						Elenco schede
					</caption>
					<% if not rs.eof then %>
						<tr><th>Trovate n&ordm; <%= Pager.recordcount %> schede in n&ordm; <%= Pager.PageCount %> pagine</th></tr>
						<%rs.AbsolutePage = Pager.PageNo
						while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
                            <tr>
								<td class="body">
									<table width="100%" border="0" cellspacing="1" cellpadding="0">
										<tr>
											<td class="<%=IIF(rs("sc_in_garanzia"), "header ", "header_disabled ")%>" colspan="6">
												<table border="0" cellspacing="0" cellpadding="0" align="right">
                                                    <tr>
                                                        <td style="font-size: 1px;">
															<% if id_centro_assistenza > 0 then %>
																<a class="button" href="SchedeMod.asp?ID=<%= rs("sc_id") %>&ID_CENTRO_ASSISTENZA=<%=id_centro_assistenza%>">MODIFICA</a>
															<% else %>
																<a class="button" href="SchedeMod.asp?ID=<%= rs("sc_id") %>">MODIFICA</a>
															<% end if %>
                                                            &nbsp;
                                                            <% if cIntero(rs("sc_rif_DDT_di_resa_id")) > 0 then %>
                        										<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare la scheda. Merce restituita.">
                        											CANCELLA
                        										</a>
															<% elseif cIntero(rs("sc_external_id")) > 0 then %>
																<a class="button_disabled" href="javascript:void(0);" title="Impossibili cancellare la scheda perch&egrave; proveniente da un import.">
                        											CANCELLA
                        										</a>
                        									<% else %>
                        										<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('SCHEDE','<%= rs("sc_id") %>');" >
                        											CANCELLA
                        										</a>
                        									<% end if %>
														</td>
													</tr>
												</table>
												<%="n. " & rs("sc_numero") & " del " & rs("sc_data_ricevimento")%>
											</td>
										</tr>
										<% if not assegnata then %>
											<tr>
												<td class="header" align="right" colspan="6">
													<a href="javascript:void(0)" class="button_L2"
														onclick="OpenAutoPositionedScrollWindow('SchedeAssegnaCentroAssistenza.asp?ID_SCHEDA=<%= rs("sc_id") %>', 'SelezioneCentroAssistenza', 450, 480, true)" 
														title="Click per aprire la finestra per la selezione del centro assistenza">
														ASSEGNA A CENTRO ASSISTENZA
													</a>
												</td>
											</tr>
										<% end if %>
										<% if cIntero(rs("sc_external_id")) > 0 then %>
											<tr>
												<td class="header OrdConfermato" colspan="4" style="font-weight:normal !important; font-size:10px;">
													scheda importata da db Access
												</td>
												<td class="header OrdConfermato" align="right" colspan="2" style="font-size:10px;">
													ext. id: &nbsp; <%=cIntero(rs("sc_external_id"))%>
												</td>
											</tr>
										<% end if %>
                                        <tr>
                                            <td class="label" style="width:20%;">stato:</td>
											<% sql = "SELECT sts_nome_it FROM sgtb_stati_schede WHERE sts_id = " & rs("sc_stato_id") %>
                                            <td class="content" style="width:35%;"><%= GetValueList(conn, NULL, sql)%></td>
											<td class="label" style="width:13%;">in garanzia:</td>
                                            <td class="content">
												<input disabled type="checkbox" class="checkbox" <%= chk(rs("sc_in_garanzia"))%>>
											</td>
											<td class="content_right" colspan="2">
												<% if rs("sc_richiesta_garanzia") then %>
													richiesta garanzia
												<% else %>
													&nbsp;
												<% end if %>
											</td>
										</tr>
										<% if assegnata then %>
											<tr>
												<td class="label_no_width">centro assistenza:</td>
												<% sql = "SELECT * FROM gv_agenti WHERE ag_id = " & rs("sc_centro_assistenza_id") 
												rsa.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText %>
												<td class="content" colspan="3"><%= ContactFullName(rsa)%></td>
												<td class="content_right" colspan="2">
													<% if (rs("sc_stato_id") <> StatoSchedaConclusa) then %>
														<a class="button_l2" href="Schede.asp?ARCHIVIA_SC_ID=<%= rs("sc_id") %>">ARCHIVIA SCHEDA</a>
													<% else %>
														&nbsp;
													<% end if %>
												</td>
												<% rsa.close %>
											</tr>
										<% end if %>
										<tr>
                                            <td class="label_no_width">cliente:</td>
											<% sql = "SELECT * FROM gv_rivenditori WHERE riv_id = " & rs("sc_cliente_id") 
											rsa.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText %>
                                            <td class="content" colspan="5"><%= ContactFullName(rsa)%></td>
											<% rsa.close %>
										</tr>
										<tr>
                                            <td class="label_no_width">modello:</td>
											<% sql = "SELECT art_nome_it FROM gv_articoli WHERE rel_id = " & rs("sc_modello_id") %>
                                            <td class="content" colspan="5"><%= GetValueList(conn, NULL, sql)%></td>
										</tr>
									</table>
								</td>
							</tr>
							<% rs.moveNext
						wend%>
						<tr>
							<td class="footer" style="text-align:left;" colspan="8">
								<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
							</td>
						</tr>
					<%else%>
						<tr><td class="noRecords">Nessun record trovato</th></tr>
					<% end if %>
				</table>
			</td> 
		</tr>
		<tr><td>&nbsp;</td></tr>
	</table>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set rsa = nothing
set conn = nothing
%>
