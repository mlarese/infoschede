<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1073741824 %>

<!--#INCLUDE FILE="Intestazione.asp" -->
<!--#INCLUDE FILE="../nextCom/Imports/Tools_Import.asp" -->

<% 

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Import modelli"
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
		<caption>Import dati dei modelli</caption>
        <tr><th colspan="3">PARAMETRI DI IMPORT</th></tr>
		
		<% if not (request("importa")<>"" AND request("file_import")<>"") then %>
            <form action="" method="post" id="form1" name="form1">
			<tr>
				<td class="label" style="width:22%;">file da importare:</td>
				<td class="content" colspan="2">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "file_import", request("file_import") , "width:400px;", true) %>
                    <span class="note">Selezionare il file EXCEL (FORMATO EXCEL 2003) dal quale verranno importati i ricambi.</span>
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
                        </td>
                    </tr>
                    <tr>
        				<td class="footer" colspan="3">
        					(*) Campi obbligatori.
        					<input style="width:20%;" type="submit" class="button" name="importa" value="IMPORTA MODELLI">
        				</td>
        			</tr>
                </form>
                
            <% else %>
                <tr><th colspan="3">ESECUZIONE IMPORT</th></tr>
                <% sql = "SELECT * FROM [" & ParseSQL(request("tabella_import"), adChar) & "] ORDER BY Codice "
                set Srs = Server.CreateObject("ADODB.Recordset")
                Srs.open sql, Sconn, adOpenStatic, adLockOptimistic, adCmdText %>
                <tr>
    				<td class="label" style="width:18%;">n. modelli:</td>
    				<td class="content"><%= Srs.recordcount %></td>
    			</tr>
				<tr>
					<td colspan="3">
						<%
						conn.begintrans

						'CALL conn.execute(sql, , adExecuteNoRecords)
						dim  codice, descr, categoriaId, marcaId, art_id
			
						while not Srs.eof
							codice = Trim(Cstring(SourceField(Srs, "Codice", true)))
							descr = Trim(Cstring(SourceField(Srs, "Descrizione", true)))
							errore = ""
							if descr <> "" then
								
								sql = "SELECT tip_id FROM gtb_tipologie WHERE tip_codice LIKE '"&Trim(Cstring(SourceField(Srs, "CATEGORIE", true)))&"'" 
								categoriaId = cIntero(GetValueList(conn, NULL, sql))
								if categoriaId = 0 then
									errore = "CATEGORIA NON TROVATA per articolo: cod. " & codice & ", " & descr
								end if
								
								sql = "SELECT mar_id FROM gtb_marche WHERE mar_codice LIKE '"&Trim(Cstring(SourceField(Srs, "Id_Costruttore", true)))&"'" 
								marcaId = cIntero(GetValueList(conn, NULL, sql))
								if marcaId = 0 then
									errore = errore & "<br>MARCA NON TROVATA per articolo: cod. " & codice & ", " & descr
								end if
								
								if errore <> "" then
									%>
									<table>
										<tr>
											<td class="content_b" colspan="3">
												<%=errore%>
											</td>
										</tr>
									</table>
									<%
									errore = ""
								else
									sql = "SELECT * FROM gtb_articoli "
									rsd.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
									
									rsd.AddNew
									rsd("art_nome_it") = descr
									rsd("art_cod_int") = codice
									rsd("art_prezzo_base") = 0
									rsd("art_tipologia_id") = categoriaId
									rsd("art_marca_id") = marcaId
									rsd("art_iva_id") = 1
									
									rsd("art_giacenza_min") = 1
									rsd("art_lotto_riordino") = 1
									rsd("art_qta_min_ord") = 1
									rsd("art_NovenSingola") = false
									rsd("art_se_accessorio") = false
									rsd("art_ha_accessori") = false
									rsd("art_se_bundle") = false
									rsd("art_se_confezione") = false
									rsd("art_in_bundle") = false
									rsd("art_in_confezione") = false
									rsd("art_varianti") = false
									rsd("art_disabilitato") = false
									rsd("art_insData") = Now()
									rsd("art_insAdmin_id") = cIntero(Session("ID_ADMIN"))
									rsd("art_modData") = Now()
									rsd("art_modAdmin_id") = cIntero(Session("ID_ADMIN"))
									rsd("art_applicativo_id") = Session("ID_SITO")
									rsd("art_unico") = false
									rsd("art_spedizione_id") = 1
									rsd("art_ordine") = Trim(Cstring(SourceField(Srs, "Id_Modello", true)))
									rsd("art_qta_max_ord") = 1
									
									rsd.Update
									art_id = rsd("art_id")
									rsd.close
									
									'GREL_ART_VALORI
									sql = "SELECT * FROM grel_art_valori "
									rsd.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
									
									rsd.AddNew
									
									rsd("rel_art_id") = art_id
									rsd("rel_prezzo") = 0
									rsd("rel_giacenza_min") = 1
									rsd("rel_lotto_riordino") = 1
									rsd("rel_qta_min_ord") = 1
									rsd("rel_cod_int") = codice
									rsd("rel_disabilitato") = false
									rsd("rel_insData") = Now()
									rsd("rel_insAdmin_id") = cIntero(Session("ID_ADMIN"))
									rsd("rel_modData") = Now()
									rsd("rel_modAdmin_id") = cIntero(Session("ID_ADMIN"))
									
									rsd.Update
									rsd.close
									
								end if
							end if
							Srs.movenext
						wend
						
						Srs.close

						'chiusura transazione di import
						conn.committrans 

						%>
					</td>
				</tr>
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