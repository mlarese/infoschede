<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_IndexContent.asp" -->


<%
'check dei permessi dell'utente
if NOT index.ChkPrm(prm_indice_accesso, 0) then %>
<script type="text/javascript">
	window.close()
</script>
<%
end if

' <!--#INCLUDE FILE="../Tools4Admin.asp" -->
' <!--#INCLUDE FILE="../class_testata.asp" -->
' <!--#INCLUDE FILE="../ToolsDescrittori.asp" -->
' <!--#INCLUDE FILE="../Tools4Color.asp" -->


'--------------------------------------------------------
sezione_testata = "import indirizzi alternativi" %>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim conn, rs, rsr, rsv, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")
%>

<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Import indirizzi alternativi da file excel di webmaster tools</caption>
        <tr><th colspan="3">PARAMETRI DI IMPORT</th></tr>
		<% if request("importa")="" then %>
		
            <form action="" method="post" id="form1" name="form1">
			<tr>
				<td class="label" style="width:18%;">file da importare:</td>
				<td class="content" colspan="2">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "file_import", request("file_import") , "width:400px;", true) %>
                    <span class="note">Selezionare il file (EXCEL) dal quale devono essere importati gli URL.</span>
				</td>
			</tr>
			<tr>
				<td class="label" style="width:18%;">&nbsp;</td>
				<td class="note" colspan="2">ATTENZIONE! Il file excel deve avere una colonna denominata "url". <br>
					Una riga puù avere il seguente formato: "http://www.agenziarallo.it/dynalay.asp?PAGINA=329,404 (Non trovato),4 pagine,15/09/11". <br>
					Si occuperà lo script di estrarre la parte di url interessata, in questo esempio: "dynalay.asp?PAGINA=329"</td>
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
        					<input style="width:20%;" type="submit" class="button" name="importa" value="IMPORTA URL ALTERNATIVI">
        				</td>
        			</tr>
                </form>
                
            <% else %>
                <tr><th colspan="3">ESECUZIONE IMPORT</th></tr>
                <% sql = "SELECT * FROM [" & ParseSQL(request("tabella_import"), adChar) & "]"
                set Srs = Server.CreateObject("ADODB.Recordset")
                Srs.open sql, Sconn, adOpenStatic, adLockOptimistic, adCmdText 
				
				'ListRecordset Srs, true
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
				
				dim url_array, url, idx, web_id, urlBase, lingua, lan
				idx = CIntero(request("IDX"))
				web_id = cIntero(GetValueList(conn, NULL, "SELECT idx_webs_id FROM tb_contents_index WHERE idx_id = " & idx))
				urlBase = GetSiteUrl(conn, web_id, NULL)
				
				while not Srs.eof
					
					url_array = Split(Srs("url"),",")
					url = Trim(url_array(0))
					url = Trim(Replace(url, urlBase & "/", ""))
					
					sql = "SELECT * FROM rel_index_url_redirect WHERE riu_idx_id="&idx&" AND riu_url LIKE '"&ParseSQL(url, adChar)&"'"
					rsr.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
						
					if rsr.recordcount = 0 then
						lingua = ""
						for each lan in Application("LINGUE")
							if inStr("/"&url, "/"&lan&"/")>0 then
								lingua = lan
							end if
						next
						if lingua = "" then
							lingua = "it"
						end if
						
						rsr.addNew
						
						rsr("riu_idx_id") = idx
						rsr("riu_url") = url
						rsr("riu_lingua") = lingua
						rsr("riu_insData") = Now
						rsr("riu_insAdmin_id") = Session("ID_ADMIN")
						rsr("riu_modData") = Now
						rsr("riu_modAdmin_id") = Session("ID_ADMIN")
							
						rsr.update
						
						%>
							<tr>
								<td class="label" style="width:10%;"><%= lingua%></td>
								<td class="content"><%= url %></td>
							</tr>
						
						<% 
					end if	
					rsr.close

					Srs.movenext
                wend
                Srs.close

                'chiusura transazione di import
                conn.committrans 
				%>
                <tr>
                    <td class="content_b" colspan="3">IMPORT URL COMPLETATO</td>
                </tr>
        		<tr>
        			<td class="footer" colspan="6">
        				<a class="button" href="" onclick="window.close()">FINE</a>
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

<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>