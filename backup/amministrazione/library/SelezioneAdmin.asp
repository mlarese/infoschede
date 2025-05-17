<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="ClassPageNavigator.asp" -->
<%'--------------------------------------------------------
sezione_testata = "Selezione amministratori" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

'imposta parametri di caricamento e funzionamento
dim FieldName, FormName, MultipleSelection, Prefix, ContattiSqlCondition

FieldName = cString(request.querystring("FieldName"))
FormName = cString(request.querystring("FormName"))
MultipleSelection = ( cIntero(request.querystring("MultipleSelection"))<>0 )
Prefix = FormName & "_" & FieldName & "_contatti_"
ContattiSqlCondition = session("CONDIZIONE_SELEZIONE_ADMIN_" & FormName & "_" & FieldName)
'ContattiSqlCondition = session("CONDIZIONE_SELEZIONE_CONTATTI_" & FormName & "_" & FieldName)

dim listType
ListType = UCase(request.querystring("ListType"))

dim conn, rs, rsE, sql, Pager
dim ListaValori, Cells, Selectable
set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsE = Server.CreateObject("ADODB.RecordSet")

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset(Prefix)
	if not(request("tutti")<>"") then
		CALL SearchSession_Set(Prefix)
	end if
end if

if Session(Prefix & "denominazione")<>"" then
     sql = sql & " AND " & SQL_FullTextSearch(Session(Prefix & "denominazione"), "admin_cognome;admin_nome")
end if

if Session(Prefix & "email")<>"" then
     sql = sql & " AND admin_email LIKE '%"& ParseSQL(Session(Prefix & "email"), adChar) &"%'"
end if

if Session(Prefix & "cell")<>"" then
     sql = sql & " AND admin_cell LIKE '%"& ParseSQL(Session(Prefix & "cell"), adChar) &"%'"
end if

if Session(Prefix & "fax")<>"" then
     sql = sql & " AND admin_fax LIKE '%"& ParseSQL(Session(Prefix & "fax"), adChar) &"%'"
end if

'composizione query (ATTENZIONE: la composizione viene fatta al contrario, partendo dalle condizioni)
sql = " WHERE " + IIF(ContattiSqlCondition<>"", ContattiSqlCondition, " (1=1) ") + _
      sql + _
      " ORDER BY admin_cognome"

sql = " SELECT * FROM tb_admin " + sql

'response.write sql
'response.end
CALL Pager.OpenSmartRecordset(conn, rs, sql, 20)
%>

<script language="JavaScript" type="text/javascript">
    function seleziona_contatto(sender, nome){
        opener.<%= FieldName %>_selezione_contatto(sender, sender.value, nome);
        <% if not MultipleSelection then %>
            window.close();
        <% end if %>
    }
	
	function Tutti() {
		for(var i=0; i < form1.elements.length; i++)
			if (form1.elements[i].id.substring(0, 10) == "selezione_" && !form1.elements[i].checked)
				form1.elements[i].click()
	}
	
	function Reset() {
		for(var i=0; i < form1.elements.length; i++)
			if (form1.elements[i].id.substring(0, 10) == "selezione_" && form1.elements[i].checked)
				form1.elements[i].click()
	}
</script>
<div id="content_ridotto">
<form action="" method="post" name="form1">
    <table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:8px;">
        <caption>
            <span style="float:right; padding-top:1px;">
                <input type="submit" name="cerca" value="CERCA" class="button">
				<input type="submit" name="tutti" value="VEDI TUTTI" class="button">
            </span>
            Opzioni di ricerca
        </caption>
        <tr>
            <th>NOME / DENOMINAZIONE</th>
            <% if instr(1, ListType, "EMAIL", vbTextCompare)>0 then %>
                <th width="20%">EMAIL</th>
            <% end if %>
			<% if instr(1, ListType, "CELL", vbTextCompare)>0 then %>
                <th width="20%">Cellulare</th>
            <% end if %>
			<% if instr(1, ListType, "FAX", vbTextCompare)>0 then %>
                <th width="20%">FAX</th>
            <% end if %>
        </tr>
        <tr>
            <td class="content_center">
                <input type="text" name="search_denominazione" value="<%= replace(session(Prefix & "denominazione"), """", "&quot;") %>" style="width:100%;">
            </td>
            <%if instr(1, ListType, "EMAIL", vbTextCompare)>0 then %>
                <td class="content_center">
                    <input type="text" name="search_email" value="<%= replace(session(Prefix & "email"), """", "&quot;") %>" style="width:100%;">
                </td>
            <% end if %>
			<% if instr(1, ListType, "CELL", vbTextCompare)>0 then %>
                <td class="content_center">
                    <input type="text" name="search_cell" value="<%= replace(session(Prefix & "cell"), """", "&quot;") %>" style="width:100%;">
                </td>
            <% end if %>
			<% if instr(1, ListType, "FAX", vbTextCompare)>0 then %>
                <td class="content_center">
                    <input type="text" name="search_fax" value="<%= replace(session(Prefix & "fax"), """", "&quot;") %>" style="width:100%;">
                </td>
            <% end if %>
        </tr>
    </table>
    <table cellspacing="1" cellpadding="0" class="tabella_madre">
	    <caption>
            <% if MultipleSelection then %>
                <span style="float:right;">
                    <a id="tutti" class="button_L2" href="javascript:void(0);" onclick="Tutti()">
					    SELEZIONA TUTTI
					</a>
					&nbsp;
					<a id="nessuno" class="button_L2" href="javascript:void(0);" onclick="Reset()">
					    DESELEZIONA TUTTI
					</a>
                </span>
            <% end if %>
		    Trovati n&ordm; <%= Pager.recordcount %> amministratori in n&ordm; <%= Pager.PageCount %> pagine
        </caption>
        <% if not rs.eof then %>
		    <tr>
			    <th class="center" style="width:5%;">SCEGLI</th>
                <th colspan="2">AMMINISTRATORE</th>
			</tr>
       	<% 	rs.AbsolutePage = Pager.PageNo
		   	dim type_emailMandatory, type_faxMandatory, type_cellMandatory
			type_emailMandatory = instr(1, ListType, "EMAILMANDATORY", vbTextCompare)>0
			type_cellMandatory = instr(1, ListType, "CELLMANDATORY", vbTextCompare)>0
			type_faxMandatory = instr(1, ListType, "FAXMANDATORY", vbTextCompare)>0
			
			while not rs.eof and rs.AbsolutePage = Pager.PageNo
                ListaValori = ""
                if type_emailMandatory then
                    if IsEmail(rs("admin_email")) then
						ListaValori = ListaValori & rs("admin_email")
					end if
					
					CALL WriteContact(rs, ListaValori, VAL_EMAIL)
                
				elseif type_cellMandatory then
                    if IsPhoneNumber(rs("admin_cell")) then
						ListaValori = ListaValori & rs("admin_cell")
					end if
					
					CALL WriteContact(rs, ListaValori,VAL_CELLULARE)
				
				elseif type_faxMandatory then
                    if IsPhoneNumber(rs("admin_fax")) then
						ListaValori = ListaValori & rs("admin_fax")
					end if
					
					CALL WriteContact(rs, ListaValori, VAL_FAX)
                
				else
					CALL WriteContact(rs, ListaValori, VAL_NONE)
				end if
				
				rs.moveNext
			wend%>
			<tr>
			    <td colspan="4" class="footer" style="text-align:left;">
			        <% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
		        </td>
			</tr>
        <% else%>
		    <tr><td class="noRecords">Nessun record trovato</th></tr>
		<% end if %>
	</table>
</form>
</body>
</html>

<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>

<%
rs.close
conn.close
set rs = nothing
set rsE = nothing
set conn = nothing


'restituisce true se il contatto e selezionabile listType dovrà essere uno dei tipi di numero possibili
Function IsSelectable(value, listType)
	value = CString(value)
	if listType=VAL_EMAIL then
		IsSelectable = NOT type_emailMandatory OR value <> ""
	elseif listType=VAL_CELLULARE then
		IsSelectable = NOT type_cellMandatory OR value <> ""
	elseif listType=VAL_FAX then
		IsSelectable = NOT type_faxMandatory OR value <> ""
	else
		IsSelectable = true
	end if
End Function



'scrive le colonne per la scelta del contatto
Sub WriteContact(rs, value, typeOfVal)
	dim selectable, ID, nome
	ID = rs("id_admin")
	nome = rs("admin_cognome") &" "& rs("admin_nome")
	selectable = IsSelectable(value, typeOfVal) %>
<tr>
	<td class="content_center">
		<input type="hidden" name="nome_<%= ID %>" id="nome_<%= ID %>" value="<%= nome %>">
	    <% if Selectable then %>
	        <input class="checkbox" name="selezione" id="selezione_<%= ID %>"
				   type="<%= IIF(MultipleSelection, "checkbox", "radio") %>"
				   value="<%= ID %>"
	               onclick="seleziona_contatto(this, '<%= JSReplacerEncode(nome) %>')">
	        <script type="text/javascript">
	            document.getElementById("selezione_<%= ID %>").checked = opener.<%= FieldName %>_is_selected_contatto(<%= ID %>);
	        </script>
	    <% else %>
	        <input type="<%= IIF(MultipleSelection, "checkbox", "radio") %>" class="checkbox_disabled" disabled>
	    <% end if %>
	</td>
	<td class="<%= IIF(Selectable, "content", "content_disabled") %>">
	    <%= nome %>
	    <% if ((typeOfVal = VAL_EMAIL) or (typeOfVal = VAL_CELLULARE) or (typeOfVal = VAL_FAX))  and value<>"" then %>
	            <span class="note">( <%= value %> )</span>
	    <% elseif typeOfVal <> VAL_NONE then  %>
				<span class="note warning">( <%=IIF((typeOfVal = VAL_CELLULARE) or (typeOfVal = VAL_FAX),"Numero","E-Mail") %> non presente ) </span>
		<% end if %>
	</td>
</tr>
<%
End Sub
%>
