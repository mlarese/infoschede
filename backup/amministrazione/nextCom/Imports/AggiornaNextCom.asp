<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1073741824 %>
<% response.buffer = false %>
<% Titolo_sezione = "Import dati dei contatti da file in formato NEXT-com"
Action = "INDIETRO"
href = "default.asp"%>
<!--#include file="Intestazione.asp"-->
<% 
dim conn, rs, rsr, rsv, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")
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
            <% dim FilePath, ConnectionString, RubricaId, CntId, ListaRecapiti, Recapito, Field, Value, valore, newInsert
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
                <% sql = "SELECT * FROM [" & ParseSQL(request("tabella_import"), adChar) & "]"
                set Srs = Server.CreateObject("ADODB.Recordset")
				response.write sql
                Srs.open sql, Sconn, adOpenStatic, adLockOptimistic, adCmdText %>
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
                RubricaId = GestioneRubrica(conn, FilePath, ParseSQL(request("tabella_import"), adChar)) %>
                
               <tr>
    				<td class="label">rubrica di destinazione:</td>
    				<td class="content">
                        <% sql = "SELECT nome_rubrica FROM tb_rubriche WHERE id_rubrica=" & RubricaId%>
    	    			<%= GetValueList(conn, rs, sql) %>
    				</td>
    			</tr>
                <% set CntObj = new IndirizzarioLock
                set CntObj.conn = conn
                
                
                while not Srs.eof
                    CntObj.RemoveAll
		            
                    if SourceField(Srs, "ente-organizzazione", true) <> "" then
						newInsert = true
						'la partita iva è già presente, aggiorno il relativo contatto
                        sql = " SELECT IDElencoIndirizzi FROM tb_Indirizzario WHERE partita_iva LIKE '" & Trim(SourceField(Srs, "Partita IVA", true)) & "'"
						if cIntero(GetValueList(conn, NULL, sql))>0 then
							CntObj.LoadFromDB(cIntero(GetValueList(conn, NULL, sql)))
							newInsert = false
						end if
						
    		            CntObj("rubrica") = RubricaId
                        
                        'dati principali anagrafica
                        
                        if cString(SourceField(Srs, "ente-organizzazione", true))<>"" then
                            CntObj("IsSocieta") = true
                        end if
              
                        CntObj("TitoloElencoIndirizzi") = SourceField(Srs, "titolo", true)
                        CntObj("NomeElencoIndirizzi") = SourceField(Srs, "nome", true)
                        CntObj("CognomeElencoIndirizzi") = SourceField(Srs, "cognome", true)
                        CntObj("NomeOrganizzazioneElencoIndirizzi") = SourceField(Srs, "ente-organizzazione", true) & IIF(FieldExists(Srs, "fg"), " ", "") & SourceField(Srs, "fg", true)
                        CntObj("IndirizzoElencoIndirizzi") = SourceField(Srs, "indirizzo", true) & IIF(FieldExists(Srs, "indirizzo_2"), "; ", "") &  SourceField(Srs, "indirizzo_2", true)
                        CntObj("LocalitaElencoIndirizzi") = SourceField(Srs, "Localita", true)
						CntObj("CAPElencoIndirizzi") = SourceField(Srs, "cap", true)
						CntObj("CittaElencoIndirizzi") = SourceField(Srs, "citta", true)
						'valore = Trim(SourceField(Srs, "citta", true))
						'response.write "valore=" & valore & "<br>"
						'if valore <> "" then
					'		CntObj("CittaElencoIndirizzi") = Right(valore, Len(valore)-6)
					'		CntObj("CAPElencoIndirizzi") = Left(valore, 5)
				'		end if
						valore = ""
                        CntObj("statoProvElencoIndirizzi") = SourceField(Srs, "provincia", true)
                        CntObj("CountryElencoIndirizzi") = SourceField(Srs, "Nazione", true)
                        CntObj("CF") = SourceField(Srs, "Codice fiscale", true)
						CntObj("partita_iva") = SourceField(Srs, "Partita IVA", true)
                        CntObj("QualificaElencoIndirizzi") = SourceField(Srs, "Qualifica", true)
						CntObj("ZonaElencoIndirizzi") = SourceField(Srs, "Zona", true)
						CntObj("NoteElencoIndirizzi") = SourceField(Srs, "note", true)
						
						if newInsert then
							CntId = CntObj.InsertIntoDB()
						else
							CntObj.UpdateDB()
						end if
                        
                        rsr.movefirst
                        while not rsr.eof 
                            Field = cString(rsr("nome_tipoNumero")) %>
                            <!-- <%= field %> -->
                            <%Value = cString(SourceField(Srs, Field, true))
                            if value <> ""  then
                                ListaRecapiti = split(value, ", ")
                                for each Recapito in ListaRecapiti
                                    if Trim(Recapito)<>"" then
										valore = Trim(Recapito)
										if rsr("id_tipoNumero")=VAL_TELEFONO OR rsr("id_tipoNumero")=VAL_FAX then
											valore = Replace(valore, ".", "")
											valore = Replace(valore, "/", "")
											if cIntero(left(valore, 1)) <> 0 then
												valore = "041" & valore
											end if
										end if
                                        CALL CntObj.AddValoreNumero(CntId, rsr("id_tipoNumero"), rsr("id_tipoNumero")=VAL_EMAIL, valore)
                                    end if
                                next
                            end if
                            rsr.movenext
                        wend
                    end if
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
set rsv = nothing
set conn = nothing
%>