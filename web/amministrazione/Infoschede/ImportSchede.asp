<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.buffer = false %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1073741824 %>

<!--#INCLUDE FILE="Intestazione.asp" -->
<!--#INCLUDE FILE="../nextCom/Imports/Tools_Import.asp" -->

<% 

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Import schede"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Tabelle.asp"
dicitura.scrivi_con_sottosez()  


dim conn, rs, rsd, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.RecordSet")

%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Import dati delle schede</caption>
        <tr><th colspan="3">PARAMETRI DI IMPORT</th></tr>
		
		<% if not (request("importa")<>"" AND request("file_import")<>"") then %>
            <form action="" method="post" id="form1" name="form1">
			<tr>
				<td class="label" style="width:22%;">file da importare:</td>
				<td class="content" colspan="2">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "file_import", request("file_import") , "width:400px;", true) %>
                    <span class="note">Selezionare il file ACCESS dal quale verranno importate le schede.</span>
				</td>
			</tr>
			<% Session("ERRORE") = "" %>
			<tr>
				<td class="footer" colspan="3">
					(*) Campi obbligatori.
					<input style="width:20%;" type="submit" class="button" name="importa" value="AVANTI &gt;&gt;">
				</td>
			</tr>
			</form>
        <% else %>
            <tr>
				<td class="label" style="width:18%;">file da importare:</td>
				<td class="content">
					<%= request("file_import") %>
				</td>
			</tr>
            <% dim FilePath, ConnectionString, Field, errore
            dim Sconn, Srs
            
            'costruzione stringa di connessione al database
            FilePath = replace(Application("IMAGE_PATH") & Application("AZ_ID") & "\images\" & request("file_import"), "\\", "\")
            select case uCase(right(trim(request("file_import")), 3))
                case "MDB"
                    ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + _
                                       "Data Source=" & FilePath & ";"
                case "XLS"
                    ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath &_
					                   ";Extended Properties=""Excel 8.0;HDR=YES;"";Jet OLEDB:Engine Type=35;"
                case else 
					Session("ERRORE") = "FORMATO DEL FILE NON RICONOSCIUTO"
					%>
					<script>
						window.location.href = window.location.pathname;
					</script>
					<% response.end 
            end select 

            'APERTURA CONNESSIONE
            set Sconn = Server.CreateObject("ADODB.Connection")
            Sconn.open ConnectionString

            %>
            <tr>
				<td class="label" style="width:18%;">stringa di connessione:</td>
				<td class="content"><%= ConnectionString %></td>
			</tr>
            <% if request("tabella_import")="" then %> 
                <tr><th colspan="3">SELEZIONE TABELLA DI IMPORT</th></tr>
                <form action="" method="post" id="form1" name="form1">
                    <% for each field in request.form %>
                        <input type="hidden" name="<%= field %>" value="<%= request.form(field) %>">
                    <% next %>
                    <tr>
        				<td class="label"><%= IIF(instr(1, FilePath, "mdb", vbTextCompare)>0, "tabella", "foglio") %> sorgente:</td>
		        		<td class="content">
                            <% set rs = Sconn.OpenSchema(adSchemaTables)
                            CALL DropDownRecordset(rs, "table_name", "table_name", "tabella_import", "", true, "", LINGUA_ITALIANO) %>
							Scegliere tabella "Scheda Cliente Distribuzione"
                        </td>
                    </tr>
                    <tr>
        				<td class="footer" colspan="3">
        					(*) Campi obbligatori.
        					<input style="width:20%;" type="submit" class="button" name="importa" value="IMPORTA SCHEDE">
        				</td>
        			</tr>
                </form>
                
            <% else %>
                <tr><th colspan="3">ESECUZIONE IMPORT</th></tr>
				<tr>
					<td colspan="3">
						url per il lancio automatico: 
						?importa=<%=Server.urlEncode(request("importa"))%>&file_import=<%=server.urlEncode(request("file_import"))%>&tabella_import=<%=server.UrlEncode(request("tabella_import"))%>
					</td>
				</tr>
                <% 	
				
				dim ultima_inserita
				sql = "SELECT TOP 1 sc_external_id FROM sgtb_schede ORDER BY sc_external_id DESC "
				ultima_inserita = cIntero(GetValueList(Conn, NULL, sql))
				
				sql = " SELECT top 110 * FROM [" & ParseSQL(request("tabella_import"), adChar) & "] " + _
					  " WHERE id_scheda > " & ultima_inserita & _
					  " ORDER BY ID_SCHEDA"
                set Srs = Server.CreateObject("ADODB.Recordset")
                Srs.open sql, Sconn, adOpenStatic, adLockOptimistic, adCmdText %>
                <tr>
    				<td class="label" style="width:18%;">n. modelli:</td>
    				<td class="content"><%= Srs.recordcount %></td>
    			</tr>
			</table>
						<%
						conn.begintrans

						dim  id_cliente, id_modello, modello_altro, art_from_access, cod_from_access, prezzo, id_ric, sconto, sc_id, i
			
						while not Srs.eof AND sRs.AbsolutePosition < 101
							id_cliente = CIntero(Trim(Cstring(SourceField(Srs, "Id_Cliente", true))))
							id_modello = CIntero(Trim(Cstring(SourceField(Srs, "id_Modello", true))))
							errore = ""%>
							<!-- 
							Id_Scheda = <%= srs("Id_Scheda") %>
							id_cliente = <%= id_cliente %> 
							id_cliente = <%= id_cliente %> 
							id_modello = <%= id_modello %> 
							-->
							<%
							if id_cliente > 0 AND id_modello > 0  then
								
								sql = "SELECT riv_id FROM gv_rivenditori WHERE CONVERT(int, ISNULL(PraticaPrefisso, 0)) = " & id_cliente
								id_cliente = cIntero(GetValueList(conn, NULL, sql))
								if id_cliente = 0 then
									errore = Srs("Numero")&" - CLIENTE NON TROVATO: " & Trim(Cstring(SourceField(Srs, "Id_Cliente", true)))
									'cliente non trovato: la assegna ad infoservice
									id_cliente = 13151
								end if
								
								'query su db_access
								sql = "SELECT Codice FROM Modelli WHERE Id_Modello = " & Srs("Id_modello") 'cerco per codice
								cod_from_access = GetValueList(Sconn, NULL, sql)
								sql = " SELECT rel_id FROM gtb_articoli INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " & _
									  " WHERE art_cod_int LIKE '"& ParseSQL(cod_from_access, adChar) &"'" 
								id_modello = cIntero(GetValueList(conn, NULL, sql))
								if id_modello = 0 then
									sql = "SELECT Descrizione FROM Modelli WHERE Id_Modello = " & Srs("Id_modello") 'cerco per nome
									art_from_access = GetValueList(Sconn, NULL, sql)
									sql = " SELECT rel_id FROM gtb_articoli INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " & _
										  " WHERE RTRIM(LTRIM(art_nome_it)) LIKE '"& ParseSQL(Trim(art_from_access), adChar) &"'" 
									id_modello = cIntero(GetValueList(conn, NULL, sql))
									if id_modello = 0 then
									
										'registro come "altro modello"
										id_modello = 78584
										modello_altro = cod_from_access + " " + art_from_access
										errore = errore & "<br>"&Srs("Numero")&" - MODELLO NON TROVATO: " & Trim(Cstring(SourceField(Srs, "id_Modello", true)))
									else
										modello_altro = ""
									end if
								else
									modello_altro = ""
								end if
								
								if errore <> "" then
									%>
									<table>
										<tr>
											<td class="content_b" colspan="3">
												id scheda: <%= srs("id_scheda") %><br>
												<%=errore%>
											</td>
										</tr>
									</table>
									<%
									errore = ""
								end if 
								if errore = "" then
									
									sql = "SELECT * FROM sgtb_schede"
									rsd.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
									
									rsd.AddNew
									rsd("sc_external_id") = Trim(Cstring(SourceField(Srs, "Id_Scheda", true)))
									rsd("sc_stato_id") = 1 'fatturata
									rsd("sc_numero") = Trim(Cstring(SourceField(Srs, "Numero", true)))
									if IsDate(Trim(Cstring(SourceField(Srs, "DataRicevimento", true)))) then
										rsd("sc_data_ricevimento") = Trim(Cstring(SourceField(Srs, "DataRicevimento", true)))
									else
										rsd("sc_data_ricevimento") = NULL
									end if
									rsd("sc_cliente_id") = id_cliente
									rsd("sc_centro_assistenza_id") = 12718	'ag_id corrispondente al centro assistenza hidroservice									
									rsd("sc_modello_id") = id_modello				
									rsd("sc_modello_altro") = modello_altro
									rsd("sc_matricola") = Trim(Cstring(SourceField(Srs, "Matricola", true)))
									if IsDate(Trim(Cstring(SourceField(Srs, "DataAcquisto", true)))) then
										rsd("sc_data_acquisto") = Trim(Cstring(SourceField(Srs, "DataAcquisto", true)))
									else
										rsd("sc_data_acquisto") = NULL
									end if
									rsd("sc_numero_scontrino") = Trim(Cstring(SourceField(Srs, "NumeroScontrino", true)))
									rsd("sc_in_garanzia") = cBoolean(SourceField(Srs, "Garanzia", true), false)
									if IsDate(Trim(Cstring(SourceField(Srs, "DataConsegna", true)))) then
										rsd("sc_data_fine_lavoro") = Trim(Cstring(SourceField(Srs, "DataConsegna", true)))
									else
										rsd("sc_data_fine_lavoro") = NULL
									end if
									rsd("sc_note_cliente") = Trim(Cstring(SourceField(Srs, "Note", true)))
									rsd("sc_numero_DDT_di_carico") = Trim(Cstring(SourceField(Srs, "DDT", true)))

									if IsDate(Trim(Cstring(SourceField(Srs, "DataDDT", true)))) then
										rsd("sc_data_DDT_di_carico") = Trim(Cstring(SourceField(Srs, "DataDDT", true)))
									else
										rsd("sc_data_DDT_di_carico") = NULL
									end if
									rsd("sc_guasto_segnalato_altro") = Trim(Cstring(SourceField(Srs, "GuastoSegnalato", true)))
									rsd("sc_guasto_riscontrato_altro") = Trim(Cstring(SourceField(Srs, "GuastoRiscontrato", true)))
									
									rsd("sc_accessori_presenti_altro") = Trim(Cstring(SourceField(Srs, "AccessorioPresente1", true))) & IIF(Trim(Cstring(SourceField(Srs, "AccessorioPresente3", true)))<>""," - "&Trim(Cstring(SourceField(Srs, "AccessorioPresente3", true))),"")
									rsd("sc_note_chiusura") = Trim(Cstring(SourceField(Srs, "Esito Intervento", true)))
									
									prezzo = Cstring(SourceField(Srs, "PrezzoManodopera", true))
									prezzo = Trim(Replace(prezzo, "€", ""))
									rsd("sc_prezzo_manodopera") = prezzo
									rsd("sc_ora_manodopera_intervento") = CIntero(SourceField(Srs, "OreManodopera", true))
									
									prezzo = Cstring(SourceField(Srs, "CostoPresa", true))
									prezzo = Trim(Replace(prezzo, "€", ""))
									rsd("sc_costo_presa") = cReal(prezzo)
									
									prezzo = Cstring(SourceField(Srs, "CostoRiconsegna", true))
									prezzo = Trim(Replace(prezzo, "€", ""))
									rsd("sc_costo_riconsegna") = cReal(prezzo)
									
									rsd("sc_insData") = Now()
									rsd("sc_insAdmin_id") = cIntero(Session("ID_ADMIN"))
									rsd("sc_modData") = Now()
									rsd("sc_modAdmin_id") = cIntero(Session("ID_ADMIN"))
									
									'CALL ListRecordset(rsd, false)

									'On Error Resume Next
									rsd.Update
									'if Err.Number <> 0 then 
									'	response.write rsd("sc_numero") & "<br>"
									'end if
									'On error goto 0
									
									sc_id = rsd("sc_id")
									rsd.close
									
									
									'importo i RICAMBI	(dettagli scheda)

									for i = 1 to 4
										sql = "SELECT * FROM sgtb_dettagli_schede "
										rsd.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
										if Trim(Cstring(SourceField(Srs, "CodRicambio" & i, true)))<>"" OR Trim(Cstring(SourceField(Srs, "Ricambio" & i, true)))<>"" then
											sql = " SELECT rel_id FROM gtb_articoli INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " & _
													  " WHERE art_cod_int LIKE '"&ParseSQL(Trim(Cstring(SourceField(Srs, "CodRicambio" & i, true))), adChar)&"' OR " & _
													  "	 RTRIM(LTRIM(art_nome_it)) LIKE '"&ParseSQL(Trim(Cstring(SourceField(Srs, "Ricambio" & i, true))), adChar) &"'"
											id_ric = cIntero(GetValueList(conn, NULL, sql))
																			
											rsd.AddNew
											
											rsd("dts_ricambio_id") = id_ric
											rsd("dts_ricambio_codice") = Trim(Cstring(SourceField(Srs, "CodRicambio" & i, true)))
											rsd("dts_ricambio_nome") = Trim(Cstring(SourceField(Srs, "Ricambio" & i, true)))
											prezzo = Trim(Cstring(SourceField(Srs, "PrezzoRicambio" & i, true)))
											prezzo = Trim(Replace(prezzo, "€", ""))
											rsd("dts_ricambio_prezzo") = cReal(prezzo)
											rsd("dts_ricambio_qta") = cIntero(Trim(Cstring(SourceField(Srs, "QtRicambio" & i, true))))
											sconto = Trim(Cstring(SourceField(Srs, "ScontoRicambio" & i, true)))
											sconto = Trim(Replace(sconto, "%", ""))
											rsd("dts_ricambio_sconto") = sconto
											rsd("dts_scheda_id") = sc_id
											prezzo = cReal(prezzo) * cIntero(Trim(Cstring(SourceField(Srs, "QtRicambio" & i, true))))
											rsd("dts_prezzo_totale") = (prezzo - ((prezzo*cReal(sconto))/100))
											
											rsd.Update
										end if
										rsd.close
									next
									
								end if
							else
								%>
									<table>
										<tr>
											<td class="content_b" colspan="3">
												Scheda non corretta: <%= srs("id_Scheda") %>
											</td>
										</tr>
									</table>
									<%
							end if
							Srs.movenext
						wend
						
						if not srs.eof _
						   AND request.servervariables("REQUEST_METHOD") <> "POST" then %>
							<script language="JavaScript" type="text/javascript">
								document.location.reload(true);
							</script>
						<% end if
						
						Srs.close

						'chiusura transazione di import
						conn.committrans 

						%>
			<table cellspacing="1" cellpadding="0" class="tabella_madre">
                <tr>
                    <td class="content_b" colspan="3">IMPORT DATI COMPLETATO</td>
                </tr>
        		<tr>
        			<td class="footer" colspan="6" style="border-bottom:1px solid #999999;">
        				<a class="button" href="default.asp">FINE</a>
        			</td>
        		</tr>
            <% end if
            
            Sconn.close
            set Sconn = nothing
        end if %>
	</table>
</div>

<%
conn.close
set rs = nothing
set rsd = nothing
set conn = nothing
%>