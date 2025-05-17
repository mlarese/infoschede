<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<% Server.ScriptTimeout = 1073741824 %>
<% Titolo_sezione = "Import dati dei contatti da file in formato NEXT-com"
Action = "INDIETRO"
href = "default.asp"%>
<!--#include file="Intestazione.asp"-->
<% 
dim conn, rs, rsr, rsDes, sql, categoria, rubrica, rubricheAgg, id_rubrica
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")
set rsDes = Server.CreateObject("ADODB.RecordSet")
rsDes.CursorLocation = adUseClient
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>Import dati dei contatti in formato NEXT-com</caption>
        <tr><th colspan="3">PARAMETRI DI IMPORT</th></tr>
		<% if request("importa")="" then %>
            <form action="" method="post" id="form1" name="form1">
			<tr>
				<td class="label" style="width:18%;">file da importare:</td>
				<td class="content" colspan="2">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "file_import", request("file_import") , "width:400px;", true) %>
                    <span class="note">Selezionare il file (EXCEL O ACCESS) dal quale vengono importati i contatti.</span>
				</td>
			</tr>
        </table>
			<% CALL FORM_SelezioneRubrica(conn) %>
            
        <table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
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
            <% dim FilePath, ConnectionString, RubricaId, CntId, ListaRecapiti, Recapito, Recapiti, rec, note, Field, Value
            dim Sconn, Srs, CntObj
            
            'costruzione stringa di connessione al database
            FilePath = replace(Application("IMAGE_PATH") & Application("AZ_ID") & "\images\" & request("file_import"), "\\", "\")
            select case uCase(right(trim(request("file_import")), 3))
                case "MDB"
                    ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + _
                                       "Data Source=" & FilePath & ";"
                case "XLS"
                    ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath &_
					                   ";Extended Properties=""Excel 8.0;HDR=YES;"";Jet OLEDB:Engine Type=35;"
                case else %>
                <tr>
                    <td class="errore" colspan="3">FORMATO DEL FILE NON RICONOSCIUTO</td>
                </tr>
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
        					<input style="width:20%;" type="submit" class="button" name="importa" value="IMPORTA CONTATTI">
        				</td>
        			</tr>
                </form>
                
            <% else %>
                <tr><th colspan="3">ESECUZIONE IMPORT</th></tr>
                <% sql = "SELECT * FROM [" & request("tabella_import") & "]"
                set Srs = Server.CreateObject("ADODB.Recordset")
                Srs.open sql, Sconn, adOpenStatic, adLockOptimistic, adCmdText 
				
				'call ListRecordset(Srs, true)
				Srs.moveFirst
					
				%>
                <tr>
    				<td class="label" style="width:18%;">contatti:</td>
    				<td class="content"><%= Srs.recordcount %></td>
    			</tr>
                <% 'apertura transazione di import
                conn.begintrans
                
                'recupera elenco recapiti
                sql = "SELECT * FROM tb_tipNumeri"
                rsr.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
                
                'gestione rubrica
                RubricaId = GestioneRubrica(conn, FilePath, request("tabella_import")) %>
                
               <tr>
    				<td class="label">rubrica di destinazione:</td>
    				<td class="content">
                        <% sql = "SELECT nome_rubrica FROM tb_rubriche WHERE id_rubrica=" & RubricaId%>
    	    			<%= GetValueList(conn, rs, sql) %>
    				</td>
    			</tr>
                <% set CntObj = new IndirizzarioLock
                set CntObj.conn = conn
                
				sql = "SELECT ict_id, ict_codice FROM tb_indirizzario_carattech WHERE ict_codice <> '' "
				rsDes.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
				Set rsDes.ActiveConnection = Nothing
                
                while not Srs.eof
                    CntObj.RemoveAll
		            
                    if (SourceField(Srs, "ente-organizzazione", true) & "" & SourceField(Srs, "cognome", true)) <> "" then
                        
    		            CntObj("rubrica") = RubricaId
                        
                        'dati principali anagrafica
                        
                        if cString(SourceField(Srs, "ente-organizzazione", true))<>"" then
                            CntObj("IsSocieta") = true
                        end if
						
						if SourceField(Srs, "categoria", true)<>"" then
							'verifica esistenza categoria
							sql = "SELECT icat_id FROM tb_indirizzario_categorie WHERE icat_codice LIKE '" + ParseSQL(SourceField(Srs, "categoria", true),adChar) + "'"
							categoria = cIntero(GetValueList(conn, rs, sql))
							if categoria > 0 then
								CntObj("cnt_categoria_id") = categoria
							else
								sql = "SELECT icat_id FROM tb_indirizzario_categorie WHERE icat_codice LIKE 'default'"
								categoria = cIntero(GetValueList(conn, rs, sql))
								if categoria > 0 then
									CntObj("cnt_categoria_id") = categoria
								end if
							end if
						end if
                        
                        CntObj("TitoloElencoIndirizzi") = SourceField(Srs, "titolo", true)
                        CntObj("NomeElencoIndirizzi") = SourceField(Srs, "nome", true)
						CntObj("SecondoNomeElencoIndirizzi") = SourceField(Srs, "secondonome", true)
                        CntObj("CognomeElencoIndirizzi") = SourceField(Srs, "cognome", true)
                        CntObj("NomeOrganizzazioneElencoIndirizzi") = Left(Trim(SourceField(Srs, "ente-organizzazione", true) & IIF(FieldExists(Srs, "fg"), " ", "") & SourceField(Srs, "fg", true)), 255)
                        CntObj("IndirizzoElencoIndirizzi") = SourceField(Srs, "indirizzo", true) & IIF(FieldExists(Srs, "indirizzo_2"), "; ", "") &  SourceField(Srs, "indirizzo_2", true)
                        CntObj("CAPElencoIndirizzi") = SourceField(Srs, "cap", true)
                        CntObj("LocalitaElencoIndirizzi") = SourceField(Srs, "Localita", true)
                        CntObj("CittaElencoIndirizzi") = Left(SourceField(Srs, "citta", true), 50)
                        CntObj("statoProvElencoIndirizzi") = SourceField(Srs, "provincia", true)
                        CntObj("CountryElencoIndirizzi") = SourceField(Srs, "Nazione", true)
                        CntObj("CF") = SourceField(Srs, "Codice fiscale", true)
						CntObj("Partita_IVA") = SourceField(Srs, "Partita IVA", true)
                        CntObj("QualificaElencoIndirizzi") = SourceField(Srs, "ruolo / qualifica", true)
						CntObj("ZonaElencoIndirizzi") = SourceField(Srs, "Zona", true)
						
						note = SourceField(Srs, "note", true)
						if SourceField(Srs, "note2", true)<>"" then
							note = note + vbCrlf + SourceField(Srs, "note2", true)
						end if
						if SourceField(Srs, "Riferimento", true)<>"" then
							note = note + vbCrlf + "Riferimento:" + SourceField(Srs, "Riferimento", true)
						elseif SourceField(Srs, "Riferimento2", true)<>"" then
							note = note + vbCrlf + "Riferimento:" + SourceField(Srs, "Riferimento2", true)
						end if
						CntObj("NoteElencoIndirizzi") = note
						CntObj("PraticaPrefisso") = SourceField(Srs, "codice", true)
                        CntId = CntObj.InsertIntoDB()
                        
                        rsr.movefirst
                        while not rsr.eof 
                            Field = cString(rsr("nome_tipoNumero")) %>
                            <!-- <%= field %> -->
                            <%Value = cString(SourceField(Srs, Field, true))
                            if value <> ""  then
                                ListaRecapiti = split(value, ",")
                                for each Recapito in ListaRecapiti
                                    if Trim(Recapito)<>"" then
										if rsr("id_tipoNumero")=VAL_EMAIL then
											recapito = replace(replace(recapito, " " , ""), ";", ",")
											recapiti = split(recapito, ",")
											for each rec in recapiti
												CALL CntObj.AddValoreNumero(CntId, rsr("id_tipoNumero"), rsr("id_tipoNumero")=VAL_EMAIL, Trim(rec))
											next
										else
											CALL CntObj.AddValoreNumero(CntId, rsr("id_tipoNumero"), rsr("id_tipoNumero")=VAL_EMAIL, Trim(Recapito))
										end if
                                    end if
                                next
                            end if
                            rsr.movenext
                        wend
						
						'rubriche aggiuntive
						if Trim(SourceField(Srs, "Rubriche", true)) <> "" then
							rubricheAgg = Split(Trim(SourceField(Srs, "Rubriche", true)), ";")
							for each rubrica in rubricheAgg
								if Trim(rubrica) <> "" then
									id_rubrica = 0
									id_rubrica = CntObj.GetRubricaByName(Trim(rubrica), true)
									if id_rubrica > 0 then
										CALL CntObj.AddToRubrica(CntId, id_rubrica)
									end if
								end if
							next
							'response.write SourceField(Srs, "Rubriche", true)
							'response.end
						end if
					
					else %>
						<tr>
							<td class="content" colspan="3">manca cognome o organizzazione</td>
						</tr>
                    <% end if
					
					'descrittori
					if not rsDes.eof then
						rsDes.Movefirst
						while not rsDes.eof
							field = rsDes("ict_codice")
							'response.write "<br>campo:" & field
							'response.write " - valore:" & SourceField(Srs, field, true)
							sql = "SELECT * FROM rel_cnt_ctech WHERE ric_cnt_id =" & CntId & " AND ric_ctech_id=" & rsDes("ict_id")
							rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
							if rs.eof then
								rs.addnew
							end if
							rs("ric_cnt_id") = cntid
							rs("ric_ctech_id") = rsDes("ict_id")
							rs("ric_valore_it") = SourceField(Srs, field, true)
							rs.update
							rs.close
							
							rsDes.moveNext
						wend
					end if
				'response.end

                    Srs.movenext
                wend
				
                
                Srs.close
                
                rsr.close
                
                'chiusura transazione di import
                conn.committrans %>
                <tr>
                    <td class="content_b" colspan="3">IMPORT DATI COMPLETATO</td>
                </tr>
        		<tr>
        			<td class="footer" colspan="6">
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
set rsr = nothing
set rsDes = nothing
set conn = nothing
%>