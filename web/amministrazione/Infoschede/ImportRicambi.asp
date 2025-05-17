<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1073741824 %>
<% response.Buffer = false %>

<!--#INCLUDE FILE="Intestazione.asp" -->
<!--#INCLUDE FILE="../nextCom/Imports/Tools_Import.asp" -->
<!--#INCLUDE FILE="../nextB2B/Tools4Save_B2B.asp" -->

<% 

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Import ricambi"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Tabelle.asp"
dicitura.scrivi_con_sottosez()  


dim conn, rs, d_rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set d_rs = Server.CreateObject("ADODB.RecordSet")

dim objVariante 
set objVariante = new GestioneVariante
set objVariante.conn = conn

%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Import dati dei ricambi</caption>
        <tr><th colspan="3">PARAMETRI DI IMPORT</th></tr>
		
		<% if not (request("importa")<>"" AND request("file_import")<>"" AND request("categoria_id")<>"" AND request("marca_id")<>"") then %>
            <form action="" method="post" id="form1" name="form1">
			<tr>
				<td class="label" style="width:22%;">file da importare:</td>
				<td class="content" colspan="2">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "file_import", request("file_import") , "width:400px;", true) %>
                    <span class="note">Selezionare il file EXCEL (FORMATO EXCEL 2003) dal quale verranno importati i ricambi.</span>
				</td>
			</tr>
			<% sql = " SELECT * FROM gtb_tipologie WHERE tip_padre_id IN (SELECT tip_id FROM gtb_tipologie " & _
					 " WHERE tip_codice LIKE '"&CODICE_CAT_RICAMBI&"') " & _
					 " ORDER BY tip_nome_it "
			%>
			<tr>
				<td class="label">categoria da associare:<br>(scelta tra tutte le sottocategorie dei ricambi)</td>
				<td class="content">
					<% CALL DropDown(conn, sql, "tip_id", "tip_nome_it", "categoria_id", request("categoria_id"), true, "style=""width:auto;""", LINGUA_ITALIANO) %>
					(*)
				</td>
			</tr>
			<% sql = " SELECT * FROM gtb_marche ORDER BY mar_nome_it "
			%>
			<tr>
				<td class="label">marca dei ricambi:</td>
				<td class="content">
					<% CALL DropDown(conn, sql, "mar_id", "mar_nome_it", "marca_id", request("marca_id"), true, "style=""width:auto;""", LINGUA_ITALIANO) %>
					(*)
				</td>
			</tr>
			<tr>
				<td class="label">colonne obbligatorie nel file excel:</td>
				<td class="content">
					Codice, Descrizione, Prezzo
				</td>
			</tr>
			<tr>
				<td class="label">colonne opzionali nel file excel:</td>
				<td class="content">
					CodAlt<br>
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
            <% dim FilePath, ConnectionString, CategoriaId, MarcaId, Field, errore, prezzo, art_id, tot
            dim Sconn, S_rs
            
            'costruzione stringa di connessione al database
            FilePath = replace(Application("IMAGE_PATH") & Application("AZ_ID") & "\images\" & replace(request("file_import"),"/", "\"), "\\", "\")
			'FilePath = "C:\frameworks\infoschede.it\database\Listini2021.mdb"
            select case uCase(right(trim(request("file_import")), 3))
                case "MDB"
                    ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + _
                                       "Data Source=" & FilePath & ";"
                case "XLS", "XLSX", "LSX"
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
			response.write ConnectionString
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
        					<input style="width:20%;" type="submit" class="button" name="importa" value="IMPORTA RICAMBI">
        				</td>
        			</tr>
                </form>
                
            <% else %>
                <tr><th colspan="3">ESECUZIONE IMPORT</th></tr>
                <% sql = "SELECT  * FROM [" & ParseSQL(request("tabella_import"), adChar) & "] "
                set S_rs = Server.CreateObject("ADODB.Recordset")
                S_rs.open sql, Sconn, adOpenStatic, adLockOptimistic, adCmdText 
				%>
                <tr>
    				<td class="label" style="width:18%;">n. ricambi:</td>
    				<td class="content"><%= S_rs.recordcount %></td>
    			</tr>
                <%
				tot = S_rs.recordcount
				
                conn.begintrans
				
				CategoriaId = cIntero(request("categoria_id"))
				MarcaId = cIntero(request("marca_id"))
				
				%>
               <tr>
    				<td class="label">categoria scelta:</td>
    				<td class="content">
                        <% sql = "SELECT tip_nome_it FROM gtb_tipologie WHERE tip_id=" & CategoriaId%>
    	    			<%= GetValueList(conn, rs, sql) %>
    				</td>
    			</tr>
				<tr>
					<td colspan="2">
						<%
						dim conta, codice, descr
						conta = 1
						
						if CategoriaId>0 and MarcaId>0 then
							sql = "DELETE FROM gtb_articoli where art_tipologia_id=" & CategoriaId & " AND art_marca_id=" & MarcaId
							call conn.execute(sql)
						end if					
						while not S_rs.eof%>
							<table	 cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:1px">
							<%
								codice = Trim(Cstring(SourceField(S_rs, "Codice", true)))
								descr = Trim(Cstring(SourceField(S_rs, "Descrizione", true)))
								if descr <> "" AND CategoriaId > 0 AND MarcaId > 0 then
								
									sql = "SELECT * FROM gtb_articoli WHERE art_tipologia_id = " & CategoriaId & " AND art_cod_int LIKE '"&codice&"'"
									d_rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
			
									if d_rs.eof then
										'nuovo inserimento
										d_rs.addNew
										d_rs("art_NoVenSingola") = false
										d_rs("art_se_accessorio") = false
										d_rs("art_ha_accessori") = false
										d_rs("art_insData") = NOW()
										d_rs("art_modData") = NOW()
										d_rs("art_in_confezione") = false
										d_rs("art_se_bundle") = false
										d_rs("art_in_bundle") = false
										d_rs("art_se_confezione") = false
										d_rs("art_varianti") = false
										d_rs("art_cod_int") = codice
										d_rs("art_nome_it") = descr
										d_rs("art_disabilitato") = false
										d_rs("art_unico") = false

										d_rs("art_prezzo_base") = cReal(Trim(s_rs("Prezzo")))
										
										d_rs("art_spedizione_id") = 1
										d_rs("art_applicativo_id") = 38

										d_rs("art_iva_id") = 8
										d_rs("art_giacenza_min") = 1
										'd_rs("art_qta_min_ord") = cIntero(Trim(s_rs("Qta min")))
										d_rs("art_lotto_riordino") = 1
										d_rs("art_qta_max_ord") = 1
										
										d_rs("art_marca_id") = MarcaId
										d_rs("art_tipologia_id") = CategoriaId
										d_rs.update
										
										
										'inserisce dati variante di default
										objVariante.InsertUpdate d_rs("art_id"), 0, "", _
															  d_rs("art_cod_int"), d_rs("art_cod_pro"), d_rs("art_cod_alt"), _
															  d_rs("art_prezzo_base"), false, 0, 0, "", _
															  false, 1, 1, 1
										%>
										<tr>
											<td width="5%" class="content">INSERIMENTO</td>
										<%
									else
										'modifica articolo già esistente
										d_rs("art_prezzo_base") = cReal(Trim(s_rs("Prezzo")))
										
										CALL AggiornaPrezziVarianti(conn, rs_guest, d_rs("art_id"))
										
										d_rs.update
										
										%>
										<tr>
											<td width="5%" class="content">MODIFICA</td>
										<%
									end if
									d_rs.close
															
									%>
									<td width="15%" class="content"><%= Trim(s_rs("Codice")) %></td>
									<td width="20%" class="content"><%= cReal(Trim(s_rs("Prezzo"))) %></td>
									<td class="content"><%= Trim(s_rs("Descrizione")) %></td>
									<td width="18%" class="content_b" style="text-align:right;"><%= s_rs.absoluteposition %> / <%= s_rs.recordcount %></td>
								</tr>
								<% end if %>
							</table>
							<% 
							s_rs.movenext
						wend
						s_rs.close

						'chiusura transazione di import
						conn.ROLLBACKTRANS 

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
set d_rs = nothing
set conn = nothing
%>