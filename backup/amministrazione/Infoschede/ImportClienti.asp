<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1073741824 %>

<!--#INCLUDE FILE="Intestazione.asp" -->
<!--#INCLUDE FILE="../nextCom/Imports/Tools_Import.asp" -->
<!--#include file="../library/ClassIndirizzarioLock.asp"-->

<% 

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Import clienti (per le schede)"
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
		<caption>Import dati clienti</caption>
        <tr><th colspan="3">PARAMETRI DI IMPORT</th></tr>
		
		<% if not (request("importa")<>"" AND request("file_import")<>"") then %>
            <form action="" method="post" id="form1" name="form1">
			<tr>
				<td class="label" style="width:22%;">file da importare:</td>
				<td class="content" colspan="2">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "file_import", request("file_import") , "width:400px;", true) %>
                    <span class="note">Selezionare il file ACCESS dal quale verranno importati i clienti.</span>
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
            dim Sconn, Srs, CntObj, CntId
            
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
							scegliere "Clienti Distribuzioneold"
                        </td>
                    </tr>
                    <tr>
        				<td class="footer" colspan="3">
        					(*) Campi obbligatori.
        					<input style="width:20%;" type="submit" class="button" name="importa" value="IMPORTA CLIENTI">
        				</td>
        			</tr>
                </form>
                
            <% else %>
                <tr><th colspan="3">ESECUZIONE IMPORT</th></tr>
                <% sql = "SELECT * FROM [" & ParseSQL(request("tabella_import"), adChar) & "] "
                set Srs = Server.CreateObject("ADODB.Recordset")
                Srs.open sql, Sconn, adOpenStatic, adLockOptimistic, adCmdText %>
                <tr>
    				<td class="label" style="width:18%;">n. clienti:</td>
    				<td class="content"><%= Srs.recordcount %></td>
    			</tr>
				<tr>
					<td colspan="3">
						<%
						conn.begintrans
						
						dim  rubricaId, ragSoc, nomCogn, part, i, ut_id, abilitazioni, profilo, email
						
						'gestione rubrica
						'RubricaId = GestioneRubrica(conn, FilePath, request("tabella_import")) 
						%>

						<% 
						set CntObj = new IndirizzarioLock
						set CntObj.conn = conn
						
			
						while not Srs.eof
							ragSoc = Trim(Cstring(SourceField(Srs, "RagioneSociale", true)))
							errore = ""
							if ragSoc <> "" then
								'CntObj("rubrica") = RubricaId
								if instr(uCase(Replace(ragSoc,".","")), "SRL")>0 OR instr(uCase(Replace(ragSoc,".","")), "SNC")>0 _
									OR instr(uCase(Replace(ragSoc,".","")), "SPA")>0 OR instr(uCase(Replace(ragSoc,".","")), "SAS")>0 OR _
										instr(uCase(ragSoc), "SOC.")>0 OR Trim(Cstring(Srs("Ragione Sociale 2")))<>"" then
									CntObj("IsSocieta") = true
									CntObj("NomeOrganizzazioneElencoIndirizzi") = ragSoc & IIF(Trim(Cstring(Srs("Ragione Sociale 2")))<>"", " - "&Trim(Cstring(Srs("Ragione Sociale 2"))), "")
									CntObj("CognomeElencoIndirizzi") = ""
									CntObj("NomeElencoIndirizzi") = ""
								else
									CntObj("IsSocieta") = false
									nomCogn = Split(ragSoc, " ")
									for i =lBound(nomCogn) to ubound(nomCogn)
										if i = 0 then
											CntObj("CognomeElencoIndirizzi") = Trim(nomCogn(i))
										elseif i = 1 then
											CntObj("NomeElencoIndirizzi") = Trim(nomCogn(i))
										else
											CntObj("NomeElencoIndirizzi") = CntObj("NomeElencoIndirizzi") & " " &  Trim(nomCogn(i))
										end if
									next
									CntObj("NomeOrganizzazioneElencoIndirizzi") = ""
								end if
								
								'mi copio l'id del db access, per il successivo import delle schede
								CntObj("PraticaPrefisso") = Trim(Cstring(Srs("id_clientiDistribuzione")))
								
								CntObj("IndirizzoElencoIndirizzi") = Trim(Cstring(SourceField(Srs, "Via", true)))
								CntObj("LocalitaElencoIndirizzi") = Trim(Cstring(SourceField(Srs, "Localita", true)))
								CntObj("cittaElencoIndirizzi") = Trim(Cstring(SourceField(Srs, "Città", true)))
								CntObj("CAPElencoIndirizzi") = Trim(Cstring(SourceField(Srs, "CAP", true)))
								CntObj("StatoProvElencoIndirizzi") = Trim(Cstring(SourceField(Srs, "Provincia", true)))
								CntObj("ZonaElencoIndirizzi") = Trim(Cstring(SourceField(Srs, "Regione", true)))
								CntObj("partita_iva") = Trim(Cstring(SourceField(Srs, "PIVA", true)))
								CntObj("NoteElencoIndirizzi") = Trim(Cstring(SourceField(Srs, "note", true)))
								'recapiti
								if Trim(Cstring(SourceField(Srs, "Telefono", true)))<>"" then
									CntObj("telefono") = Trim(Cstring(SourceField(Srs, "Telefono", true)))
								else
									CntObj("telefono") = ""
								end if
								if Trim(Cstring(SourceField(Srs, "Email", true)))<>"" then
									email = Trim(Cstring(SourceField(Srs, "Email", true)))
									if inStrRev(email,"#") > 0 then
										if inStr(email,"@") < inStr(email,"#") then
											email = Left(email, inStrRev(email,"#") - 1)
										else
											email = Replace(email,"#","")
										end if
									end if
									email = Replace(email,"mailto:","")
									CntObj("email") = email
								else
									CntObj("email") = ""
								end if
								if Trim(Cstring(SourceField(Srs, "Fax", true)))<>"" then
									CntObj("fax") = Trim(Cstring(SourceField(Srs, "Fax", true)))
								else
									CntObj("fax") = ""
								end if

								
								CntId = CntObj.InsertIntoDB()
								
								if cBoolean(Srs("FERRAMENTA"),false) then
									CntObj.AddToRubrica CntId, 95 
								end if
								
								if cBoolean(Srs("G-DISTRIBUZIONE"),false) then
									CntObj.AddToRubrica CntId, 96
								end if
								
								if cBoolean(Srs("RIPARATORE"),false) then
									CntObj.AddToRubrica CntId, 97
								end if
								
								if cBoolean(Srs("CLIENTE PROFESSIONALE"),false) then
									CntObj.AddToRubrica CntId, 98
								end if
								
								if cBoolean(Srs("CLIENTE ACQUISITO"),false) then
									CntObj.AddToRubrica CntId, 99
								end if
								
								'aggiungo utente
								CntObj("login") = RANDOM_LOGIN_E_PASSWORD
								CntObj("password") = RANDOM_LOGIN_E_PASSWORD
								CntObj("abilitato") = false
								if cBoolean(Srs("FERRAMENTA"),false) OR cBoolean(Srs("G-DISTRIBUZIONE"),false) OR cBoolean(Srs("RIPARATORE"),false) then
									abilitazioni = "USER_RIVENDITORE, B2B_I_CLIENTE"
									profilo = 3
								elseif cBoolean(Srs("CLIENTE PROFESSIONALE"),false) then
									abilitazioni = "USER_CLIENTE_PROFESSIONALE, B2B_I_CLIENTE"
									profilo = 6
								else
									abilitazioni = "B2B_I_CLIENTE"
									profilo = 2
								end if
								ut_id = CntObj.UserFromContact(CntId, abilitazioni)
								
								
								sql = "SELECT * FROM gtb_rivenditori "
								rsd.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
								rsd.AddNew
								rsd("riv_id") = ut_id
								rsd("riv_listino_id") = 141 'listino base
								rsd("riv_valuta_id") = 1
								rsd("riv_codice") = Trim(Cstring(Srs("id_clientiDistribuzione")))
								rsd("riv_modopagamento_id") = 0
								rsd("riv_profilo_id") = profilo
								rsd.Update
								
								rsd.close
							
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
								end if
							else
								%>
								<table>
									<tr>
										<td class="content_b" colspan="3">
											saltato
										</td>
									</tr>
								</table>
								<%
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