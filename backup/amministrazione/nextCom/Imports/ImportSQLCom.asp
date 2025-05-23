<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1073741824 %>
<% Titolo_sezione = "Import dati dei contatti da file in formato NEXT-com"
Action = "INDIETRO"
href = "default.asp"%>
<!--#include file="Intestazione.asp"-->
<% 
dim conn, rs, rsr, rsv, sql,TableName
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
				<td class="label" style="width:18%;">tabella da importare:</td>
				<td class="content" colspan="2">
					<input type="text" name="tabella_import" value="" />
                    <span class="note">(*) scrivi il nome della tabella da importare.</span>
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
					<%= request("tabella_import") %>
				</td>
			</tr>
            <% dim FilePath, ConnectionString, RubricaId, CntId, ListaRecapiti, Recapito, Field, Value
            dim Sconn, Srs, CntObj
            
			ConnectionString = "Provider=SQLOLEDB.1;" &_
										   "User ID=sa;" &_
										   "Password=an739NA;" &_
										   "Initial Catalog=ideal-lux;" &_
										   "Data Source=next-server;"
			
            'costruzione stringa di connessione al database
            TableName = request("tabella_import")
            

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
        				<td class="label">tabella sorgente:</td>
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
                <% sql = "SELECT * FROM " & ParseSQL(request("tabella_import"), adChar) & ""
				response.write sql
                set Srs = Server.CreateObject("ADODB.Recordset")
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
                <% 
				dim pref_int,pref_loc
				set CntObj = new IndirizzarioLock
                set CntObj.conn = conn
                
                
                while not Srs.eof
                    CntObj.RemoveAll
		            
                    if (SourceField(Srs, "ente-organizzazione", true) & "" & SourceField(Srs, "cognome", true)) <> "" then
                        
    		            CntObj("rubrica") = RubricaId
                        
                        'dati principali anagrafica
                        
                        if cString(SourceField(Srs, "ente-organizzazione", true))<>"" then
                            CntObj("IsSocieta") = true
                        end if
                        
                        CntObj("TitoloElencoIndirizzi") = SourceField(Srs, "titolo", true)
                        CntObj("NomeElencoIndirizzi") = SourceField(Srs, "nome", true)
                        CntObj("CognomeElencoIndirizzi") = SourceField(Srs, "cognome", true)
                        CntObj("NomeOrganizzazioneElencoIndirizzi") = SourceField(Srs, "ente-organizzazione", true) & IIF(FieldExists(Srs, "fg"), " ", "") &  SourceField(Srs, "fg", true)
                        CntObj("IndirizzoElencoIndirizzi") = SourceField(Srs, "indirizzo", true) & IIF(FieldExists(Srs, "indirizzo_2"), "; ", "") &  SourceField(Srs, "indirizzo_2", true)
                        CntObj("CAPElencoIndirizzi") = SourceField(Srs, "cap", true)
                        CntObj("LocalitaElencoIndirizzi") = SourceField(Srs, "Localita", true)
                        CntObj("CittaElencoIndirizzi") = SourceField(Srs, "citta", true)
                        CntObj("statoProvElencoIndirizzi") = SourceField(Srs, "provincia", true)
                        CntObj("CountryElencoIndirizzi") = SourceField(Srs, "Nazione", true)
                        CntObj("CF") = SourceField(Srs, "Codice fiscale", true)
						CntObj("partita_iva") = SourceField(Srs, "Partita IVA", true)
                        CntObj("QualificaElencoIndirizzi") = SourceField(Srs, "Qualifica", true)
						CntObj("ZonaElencoIndirizzi") = SourceField(Srs, "Zona", true)
						CntObj("NoteElencoIndirizzi") = SourceField(Srs, "note", true)
                        CntId = CntObj.InsertIntoDB()
                        
						'Prefissi
						pref_int = SourceField(Srs, "prefint", true)
						if pref_int<>"" then
							pref_int = replace(pref_int,"00","+")
						end if
						pref_loc = SourceField(Srs, "prefloc", true)
                        rsr.movefirst
                        while not rsr.eof 
                            Field = cString(rsr("nome_tipoNumero")) %>
                            <!-- <%= field %> -->
                            <%Value = cString(SourceField(Srs, Field, true))
                            if value <> ""  then
                                ListaRecapiti = split(value, ", ")
                                for each Recapito in ListaRecapiti
                                    if Trim(Recapito)<>"" then
										if not instr(1,Recapito,"@",vbTextCompare) then
											 Recapito = pref_int & pref_loc & Recapito
										end if
										CALL CntObj.AddValoreNumero(CntId, rsr("id_tipoNumero"), rsr("id_tipoNumero")=VAL_EMAIL, Trim(Recapito))
                                    end if
                                next
                            end if
                            rsr.movenext
                        wend
                    end if
					
					'inserisco i dati nel nextinfo
					CALL NextInfoInsert( conn,CntId ,CntObj("CountryElencoIndirizzi"))
					
					
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

' al limite rimane vuoto
sub NextInfoInsert( conn,CntId, country )
	dim nuovo_inserimento
	dim rs,rstemp, sql, id_area
	set rs = Server.CreateObject("ADODB.RecordSet")
	set rstemp = Server.CreateObject("ADODB.RecordSet")
	sql = "SELECT * FROM itb_anagrafiche WHERE ana_id = " & CntId
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	nuovo_inserimento = false
	if rs.eof then
		rs.AddNew
		nuovo_inserimento = true
	end if
	
	rs("ana_id") = CntId
	rs("ana_codice") = ""
	rs("ana_tipo_id") = GetValueList(conn, rstemp, "SELECT ant_id FROM itb_anagrafiche_tipi WHERE ant_nome_it LIKE 'Rivenditore'")
	if country<>"" then
		id_area = GetValueList(conn, rstemp, "SELECT are_id FROM itb_aree WHERE are_nome_it LIKE '"+replace(country,"'","''")+"'")
		if id_area > 0 then
			rs("ana_area_id") = id_area
		else
			rs("ana_area_id") = 60 'da definire
		end if
	else
		rs("ana_area_id") = 60 'da definire
	end if
	'rs("ana_descr_it") = ""
	'rs("ana_descr_en") = ""
	'rs("ana_descr_fr") = ""
	'rs("ana_descr_es") = ""
	'rs("ana_descr_de") = ""
	rs("ana_ranking") = 150
	rs("ana_censurato") = false 
	rs("ana_web_click") = 0
	'rs("ana_web_reset") = Date
	rs("ana_visibile") =  true
	rs("ana_link_attivi") = true
	rs("ana_classificazione") = ""
	'rs("ana_censurato_perche") = ""
	'rs("") =
	CALL SetUpdateParamsRS(rs, "ana_", nuovo_inserimento)
	rs.update
end sub

%>