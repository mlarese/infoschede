<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#include file="Intestazione.asp"-->
<%

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Import storico dell'indice da access o excel"
dicitura.scrivi_con_sottosez()
dim conn, rs, rsr, rsv, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")
%>

<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>Import storico dell'indice da access o excel</caption>
        <tr><th colspan="3">PARAMETRI DI IMPORT</th></tr>
		<% if request("importa")="" then %>
		
            <form action="" method="post" id="form1" name="form1">
			<tr>
				<td class="label" style="width:18%;">file da importare:</td>
				<td class="content" colspan="2">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "file_import", request("file_import") , "width:400px;", true) %>
                    <span class="note">Selezionare il file (EXCEL O ACCESS) dal quale viene importato lo storico dell'indice.</span>
				</td>
			</tr>
        </table>
		
        <table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
			<tr>
				<td class="footer" colspan="3">
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
            <% dim FilePath, ConnectionString, Field, Value
            dim Sconn, Srs, CntObj
            
            'costruzione stringa di connessione al database
            FilePath = replace(Application("IMAGE_PATH") & Application("AZ_ID") & "\images\" & request("file_import"), "\\", "\")
			FilePath = replace(FilePath, "/", "\")
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
        					<input style="width:20%;" type="submit" class="button" name="importa" value="IMPORTA STORICO INDICE">
        				</td>
        			</tr>
                </form>
                
            <% else %>
                <tr><th colspan="3">ESECUZIONE IMPORT</th></tr>
                <% sql = "SELECT * FROM [" & ParseSQL(request("tabella_import"), adChar) & "]"
                set Srs = Server.CreateObject("ADODB.Recordset")
                Srs.open sql, Sconn, adOpenStatic, adLockOptimistic, adCmdText 
				ListRecordset Srs, true
				%>
                <tr>
    				<td class="label" style="width:18%;">record indice:</td>
    				<td class="content"><%= Srs.recordcount %></td>
    			</tr>
				<tr>
    				<td class="label" style="width:100%;"><br></td>
    			</tr>
                <% 'apertura transazione di import
                conn.begintrans
				
				Srs.movefirst
				
				while not Srs.eof
					sql = "SELECT TOP 1 idx_id" &_
						   " FROM v_indice"
					if FieldExists(Srs, "co_F_table_id") AND FieldExists(Srs, "co_F_key_id") then
						sql = sql &  " WHERE co_F_table_id = " & Srs("co_F_table_id") &_
						    " AND co_F_key_id = " & Srs("co_F_key_id")
					else
						sql = sql &  " WHERE idx_id = " & Srs("riu_idx_id")
					end if
					rsv.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
					
					if rsv.recordcount > 0 then
					
						sql = "SELECT *" &_
							   " FROM rel_index_url_redirect" &_
							  " WHERE riu_url = '" & Srs("riu_url") & "'" &_
								" AND riu_lingua = '" & Srs("riu_lingua") & "'"
						rsr.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
						
						if rsr.recordcount = 0 then
						
							rsr.addNew
							
							rsr("riu_idx_id") = rsv("idx_id")
							rsr("riu_url") = Srs("riu_url")
							rsr("riu_lingua") = Srs("riu_lingua")
							rsr("riu_insData") = Now
							rsr("riu_insAdmin_id") = Session("ID_ADMIN")
							rsr("riu_modData") = Now
							rsr("riu_modAdmin_id") = Session("ID_ADMIN")
							
							rsr.update
						
						elseif FieldExists(Srs, "co_F_table_id") AND FieldExists(Srs, "co_F_key_id") then%>
							<tr>
								<td class="label" style="width:18%;">coppia ID TABELLA e CHIAVE ESTERNA gia presente nell'indice:</td>
								<td class="content"><%= Srs("co_F_table_id") %> - <%= Srs("co_F_key_id") %></td>
							</tr>
						
						<% end if
						
						rsr.close
						
						if FieldExists(Srs, "co_F_table_id") AND FieldExists(Srs, "co_F_key_id") then%>
						
							<tr>
								<td class="label" style="width:18%;">coppia ID TABELLA e CHIAVE ESTERNA ok:</td>
								<td class="content"><%= Srs("co_F_table_id") %> - <%= Srs("co_F_key_id") %></td>
							</tr>
						
					<%  end if

					elseif FieldExists(Srs, "co_F_table_id") AND FieldExists(Srs, "co_F_key_id") then %>
						<tr>
							<td class="label" style="width:18%;">coppia ID TABELLA e CHIAVE ESTERNA non trovata nell'indice:</td>
							<td class="content"><%= Srs("co_F_table_id") %> - <%= Srs("co_F_key_id") %></td>
						</tr>
						
					<% end if

					rsv.close
					Srs.movenext
					
                wend
                
                Srs.close

                'chiusura transazione di import
                conn.committrans %>
                <tr>
                    <td class="content_b" colspan="3">IMPORT DATI STORICO INDICE COMPLETATO</td>
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