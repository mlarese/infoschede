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
dicitura.sezione = "Import DDT"
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
		<caption>Import dati dei DDT</caption>
        <tr><th colspan="3">PARAMETRI DI IMPORT</th></tr>
		
		<% if not (request("importa")<>"" AND request("file_import")<>"") then %>
            <form action="" method="post" id="form1" name="form1">
			<tr>
				<td class="label" style="width:22%;">file da importare:</td>
				<td class="content" colspan="2">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "file_import", request("file_import") , "width:400px;", true) %>
                    <span class="note">Selezionare il file ACCESS dal quale verranno importati i DDT.</span>
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
            dim Sconn, Srs, Srs_d
            
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
							Scegliere tabella "DDT"
                        </td>
                    </tr>
                    <tr>
        				<td class="footer" colspan="3">
        					(*) Campi obbligatori.
        					<input style="width:20%;" type="submit" class="button" name="importa" value="IMPORTA DDT">
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
				
				dim ultimo_inserito
				' sql = "SELECT TOP 1 ddt_external_id FROM sgtb_ddt ORDER BY ddt_external_id DESC "
				' ultimo_inserito = cIntero(GetValueList(Conn, NULL, sql))
ultimo_inserito = 0 
				sql = " SELECT * FROM [" & ParseSQL(request("tabella_import"), adChar) & "] " + _
					  " WHERE Id_DDT > " & ultimo_inserito & _
					  " ORDER BY Id_DDT"
                set Srs = Server.CreateObject("ADODB.Recordset")
                Srs.open sql, Sconn, adOpenStatic, adLockOptimistic, adCmdText %>
                <tr>
    				<td class="label" style="width:18%;">n. ddt:</td>
    				<td class="content"><%= Srs.recordcount %></td>
    			</tr>
			</table>
						<%
						conn.begintrans

						dim  id_cliente, id_destinazione, id_trasportatore, id_causale, i, ddt_id, sc_id, stop_insert
			
						while not Srs.eof
							id_cliente = CIntero(Trim(Cstring(SourceField(Srs, "Id_Cliente1", true))))
							id_destinazione = CIntero(Trim(Cstring(SourceField(Srs, "Id_Cliente2", true))))
							id_causale = CIntero(Trim(Cstring(SourceField(Srs, "Id_Causale", true))))
							errore = ""
							stop_insert = false

							if id_cliente > 0 AND id_destinazione > 0  then
								
								sql = "SELECT riv_id FROM gv_rivenditori WHERE CONVERT(int, ISNULL(PraticaPrefisso, 0)) = " & id_cliente
								id_cliente = cIntero(GetValueList(conn, NULL, sql))
								if id_cliente = 0 then
									errore = Srs("Id_DDT")&" - CLIENTE NON TROVATO: " & Trim(Cstring(SourceField(Srs, "Id_Cliente1", true)))
									stop_insert = true
									'cliente non trovato: la assegna ad infoservice
									'id_cliente = 13151
								end if
								
								sql = "SELECT IDElencoIndirizzi FROM tb_indirizzario WHERE CONVERT(int, ISNULL(PraticaPrefisso, 0)) = " & id_destinazione
								id_destinazione = cIntero(GetValueList(conn, NULL, sql))
								if id_destinazione = 0 then
									errore = Srs("Id_DDT")&" - DESTINAZIONE NON TROVATA: " & Trim(Cstring(SourceField(Srs, "Id_Cliente2", true)))
									'destinazione non trovato: assegno l'indirizzo di infoservice
									id_destinazione = 56883
								end if
								
								'trasportatore di default
								id_trasportatore = 14415
								
								if errore <> "" then
									%>
									<table>
										<tr>
											<td class="content_b" colspan="3">
												id DDT: <%= srs("Id_DDT") %><br>
												<%=errore%>
											</td>
										</tr>
									</table>
									<%
									errore = ""
								end if 
								if errore = "" AND not stop_insert then
									
									sql = "SELECT * FROM sgtb_ddt"
									rsd.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
									
									rsd.AddNew
									rsd("ddt_external_id") = Trim(Cstring(SourceField(Srs, "Id_DDT", true)))
									rsd("ddt_categoria_id") = 1 'categoria DDT
									rsd("ddt_numero") = CIntero(SourceField(Srs, "Id_DDT", true))
									rsd("ddt_causale_id") = id_causale
									rsd("ddt_trasportatore_id") = id_trasportatore
									rsd("ddt_cliente_id") = id_cliente
									rsd("ddt_destinazione_id") = id_destinazione
									if IsDate(Trim(Cstring(SourceField(Srs, "Data", true)))) then
										rsd("ddt_data") = Trim(Cstring(SourceField(Srs, "Data", true)))
									else
										rsd("ddt_data") = NULL
									end if
									rsd("ddt_peso") = "0"
									rsd("ddt_volume") = "0"
									rsd("ddt_numero_colli") = "0"
									'rsd("ddt_porto_id") = null
									'rsd("ddt_trasporto_id") = null
									
									rsd.Update
								
									ddt_id = rsd("ddt_id")
									rsd.close
									
									
									' schede collegate al DDT che sto importando
									
									sql = "SELECT * FROM [RifDDT] WHERE Id_DDT_Dis = " & Trim(Cstring(SourceField(Srs, "Id_DDT", true)))
									set Srs_d = Server.CreateObject("ADODB.Recordset")
									Srs_d.open sql, Sconn, adOpenStatic, adLockOptimistic, adCmdText
									while not Srs_d.eof
										sql = " SELECT sc_id FROM sgtb_schede WHERE sc_external_id = " & cIntero(Srs_d("Id_Scheda"))
										sc_id = cIntero(GetValueList(conn, NULL, sql))
										if sc_id > 0 then
											sql = "UPDATE sgtb_schede SET sc_rif_ddt_di_resa_id = " & ddt_id & " WHERE sc_id = " & sc_id
											conn.execute(sql)
										else
											response.write "<br>ERRORACCIO: scheda non trovata: "&cIntero(Srs_d("Id_Scheda"))&";   "
										end if
										Srs_d.moveNext
									wend

								end if
							else
								%>
									<table>
										<tr>
											<td class="content_b" colspan="3">
												DDT non corretto: <%= srs("Id_DDT") %>
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