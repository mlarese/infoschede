<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<% if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) AND cString(request("contatti_email_newsletter")) = "" then %>
	<script language="JavaScript" type="text/javascript">
		var url = self.location.href
		if (window.opener.document.getElementById("contatti_email_newsletter_true").checked)
		{
			url += '&contatti_email_newsletter=true'
		}
		else
		{
			url += '&contatti_email_newsletter=false'
		}
		window.location.href = url;
	</script>
<% end if %>

<!--#INCLUDE FILE="ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../NEXTcom/Tools_Contatti.asp" -->

<!--#INCLUDE FILE="categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="indexContent/ClassContent.asp" -->
<!--#INCLUDE FILE="indexContent/ClassIndex.asp" -->
<%'--------------------------------------------------------
sezione_testata = "Selezione contatti" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 


'imposta parametri di caricamento e funzionamento
dim FieldName, FormName, MultipleSelection, Prefix, ContattiSqlCondition, RubricheSqlCondition

FieldName = cString(request.querystring("FieldName"))
FormName = cString(request.querystring("FormName"))
MultipleSelection = ( cIntero(request.querystring("MultipleSelection"))<>0 )
Prefix = FormName & "_" & FieldName & "_contatti_"
ContattiSqlCondition = session("CONDIZIONE_SELEZIONE_CONTATTI_" & FormName & "_" & FieldName)
RubricheSqlCondition = session("CONDIZIONE_SELEZIONE_RUBRICHE_" & FormName & "_" & FieldName)

dim listType, typeLoginId
ListType = UCase(request.querystring("ListType"))
typeLoginId = InStr(listType, "LOGINID") > 0

dim conn, rs, rsE, sql, Pager, rubriche_visibili
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

'recupera rubriche visibili all'utente (se ambiente di sessione del NEXT-com caricato)
rubriche_visibili = GetList_Rubriche(conn, rs)
'imposta filtro di visibilita' se non c'e' gia' una ricerca.
if Session(Prefix & "rubriche") = "" then
    Session(Prefix & "rubriche") = rubriche_visibili
end if

'aggiunge filtro per rubrica
if DB_Type(conn) = DB_Access then
	sql = "IIf(IsNull(cntRel),IDElencoIndirizzi,cntRel)"
else
	sql = "CASE ISNULL(cntRel, 0) WHEN 0 THEN IDElencoIndirizzi ELSE cntRel END"
end if 
sql = " AND IdElencoIndirizzi IN ( SELECT "&sql&" " + _
								 " FROM (rel_rub_ind INNER JOIN tb_rubriche ON rel_rub_ind.id_rubrica = tb_rubriche.id_rubrica) " + _
								 " INNER JOIN tb_indirizzario ON rel_rub_ind.id_indirizzo = tb_indirizzario.IDElencoIndirizzi " + _
                                 " WHERE tb_rubriche.id_rubrica IN (" & Session(Prefix & "rubriche") & ") " + _
                                         IIF(RubricheSqlCondition<>"", " AND " + RubricheSqlCondition, "") + ") "

if Session(Prefix & "denominazione")<>"" then
     sql = sql & " AND " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session(Prefix & "denominazione"))
end if

if Session(Prefix & "login")<>"" then
     sql = sql & " AND " & SQL_FullTextSearch(Session(Prefix & "login"), "ut_login")
end if

if Session(Prefix & "email")<>"" then
     sql = sql & " AND EXISTS (SELECT * FROM tb_valoriNumeri"& _
	 		     "			   WHERE id_indirizzario = idElencoIndirizzi AND " & SQL_IsTrue(conn, "email_default") & " AND valoreNumero LIKE '%"& ParseSQL(Session(Prefix & "email"), adChar) &"%')"
end if

if Session(Prefix & "cell")<>"" then
     sql = sql & " AND EXISTS (SELECT * FROM tb_valoriNumeri"& _
	 		     "			   WHERE id_indirizzario = idElencoIndirizzi AND " & SQL_IsTrue(conn, "email_default") & " AND valoreNumero LIKE '%"& ParseSQL(Session(Prefix & "cell"), adChar) &"%')"
end if

if Session(Prefix & "fax")<>"" then
     sql = sql & " AND EXISTS (SELECT * FROM tb_valoriNumeri"& _
	 		     "			   WHERE id_indirizzario = idElencoIndirizzi AND " & SQL_IsTrue(conn, "email_default") & " AND valoreNumero LIKE '%"& ParseSQL(Session(Prefix & "fax"), adChar) &"%')"
end if

if Session(Prefix & "address")<>"" then
     sql = sql & " AND (" & SQL_FullTextSearch(Session(Prefix & "address"), "IndirizzoElencoIndirizzi") & " OR " & _
							SQL_FullTextSearch(Session(Prefix & "address"), "CittaElencoIndirizzi") & " OR " & _
							SQL_FullTextSearch(Session(Prefix & "address"), "StatoProvElencoIndirizzi") & ")"
end if 

if cIntero(Session(Prefix & "categorie"))>0 then
	sql = sql & " AND cnt_categoria_id = " & cIntero(Session(Prefix & "categorie"))
end if

if cIntero(Session(Prefix & "campagne"))>0 then
	sql = sql & " AND IDElencoIndirizzi IN (SELECT rcc_cnt_id FROM rel_cnt_campagne WHERE rcc_campagna_id = " & cIntero(Session(Prefix & "campagne")) & " ) "
end if


'composizione query (ATTENZIONE: la composizione viene fatta al contrario, partendo dalle condizioni)
sql = " WHERE " + IIF(ContattiSqlCondition<>"", ContattiSqlCondition, " (1=1) ") + _
	  IIF(DB_Type(conn) = DB_Access, " AND ISNULL(cntRel) ", " AND ISNULL(cntRel, 0) = 0 ") + _
      sql + _
      " ORDER BY ModoRegistra"

if instr(1, ListType, "LOGIN", vbTextCompare)>0 then
    sql = IIF(instr(1, ListType, "LOGINMANDATORY", vbTextCompare)>0, "INNER", "LEFT ") + _
          " JOIN tb_utenti ON tb_indirizzario.idelencoindirizzi = tb_utenti.ut_nextcom_id " + sql
end if

sql = " IDElencoIndirizzi,NomeElencoIndirizzi,CognomeElencoIndirizzi,NomeOrganizzazioneElencoIndirizzi," & _
	  " IndirizzoElencoIndirizzi, CittaElencoIndirizzi, StatoProvElencoIndirizzi, ZonaElencoIndirizzi, CAPElencoIndirizzi, CountryElencoIndirizzi, localitaElencoIndirizzi, " & _
		"isSocieta,ModoRegistra,DataIscrizione, tb_cnt_lingue.*,lingua " &_
		" FROM (tb_indirizzario INNER JOIN tb_cnt_lingue ON tb_indirizzario.lingua = tb_cnt_lingue.lingua_codice) " + sql
if instr(1, ListType, "LOGIN", vbTextCompare)>0 then
	sql = " tb_utenti.*, " + sql
end if

sql = "SELECT " + sql

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
                <input type="submit" name="cerca" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "CERCA", "SEARCH", "", "", "", "", "", "")%>" class="button">
				<input type="submit" name="tutti" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "VEDI TUTTI", "VIEW ALL", "", "", "", "", "", "")%>" class="button">
            </span>
            <%= ChooseValueByAllLanguages(Session("LINGUA"), "Opzioni di ricerca", "Search Options", "", "", "", "", "", "")%>
        </caption>
        <tr>
            <th width="35%"><%= ChooseValueByAllLanguages(Session("LINGUA"), "RUBRICA", "ADDRESS BOOK", "", "", "", "", "", "")%></th>
            <th><%= ChooseValueByAllLanguages(Session("LINGUA"), "NOME / DENOMINAZIONE", "NAME / DENOMINATION", "", "", "", "", "", "")%></th>
            <% if instr(1, ListType, "LOGIN", vbTextCompare)>0 then %>
                <th width="15%">LOGIN</th>
            <% end if
            if instr(1, ListType, "EMAIL", vbTextCompare)>0 then %>
                <th width="20%">EMAIL</th>
            <% end if %>
			<% if instr(1, ListType, "CELL", vbTextCompare)>0 then %>
                <th width="20%"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Cellulare", "Mobile", "", "", "", "", "", "")%></th>
            <% end if %>
			<% if instr(1, ListType, "FAX", vbTextCompare)>0 then %>
                <th width="20%">FAX</th>
            <% end if %>
        </tr>
        <tr>
            <td class="content_center">
                <% sql = " SELECT id_rubrica, nome_rubrica FROM tb_rubriche " &_
                         " WHERE id_rubrica IN (" & rubriche_visibili & ") " &_
                                 IIF(RubricheSqlCondition<>"", " AND " + RubricheSqlCondition, "") & _
                         " ORDER BY nome_rubrica"
                CALL dropDown(conn, sql, "id_rubrica", "nome_rubrica", "search_rubriche", Session(Prefix & "rubriche"), false, "style=""width:99%;""", Session("LINGUA"))%>	  
            </td>
            <td class="content_center">
                <input type="text" name="search_denominazione" value="<%= replace(session(Prefix & "denominazione"), """", "&quot;") %>" style="width:99%;">
            </td>
            <% if instr(1, ListType, "LOGIN", vbTextCompare)>0 then %>
                 <td class="content_center">
                    <input type="text" name="search_login" value="<%= replace(session(Prefix & "login"), """", "&quot;") %>" style="width:99%;">
                </td>
            <% end if
            if instr(1, ListType, "EMAIL", vbTextCompare)>0 then %>
                <td class="content_center">
                    <input type="text" name="search_email" value="<%= replace(session(Prefix & "email"), """", "&quot;") %>" style="width:99%;">
                </td>
            <% end if %>
			<% if instr(1, ListType, "CELL", vbTextCompare)>0 then %>
                <td class="content_center">
                    <input type="text" name="search_cell" value="<%= replace(session(Prefix & "cell"), """", "&quot;") %>" style="width:99%;">
                </td>
            <% end if %>
			<% if instr(1, ListType, "FAX", vbTextCompare)>0 then %>
                <td class="content_center">
                    <input type="text" name="search_fax" value="<%= replace(session(Prefix & "fax"), """", "&quot;") %>" style="width:99%;">
                </td>
            <% end if %>
        </tr>
		<%
		dim colspan_address
		colspan_address = 2
		if instr(1, ListType, "LOGIN", vbTextCompare)>0 then colspan_address = colspan_address + 1 end if
		if instr(1, ListType, "EMAIL", vbTextCompare)>0 then colspan_address = colspan_address + 1 end if
		if instr(1, ListType, "CELL", vbTextCompare)>0 then colspan_address = colspan_address + 1 end if
		if instr(1, ListType, "FAX", vbTextCompare)>0 then colspan_address = colspan_address + 1 end if
		%>
		<tr>
			<td colspan="<%=colspan_address%>">
				<table cellspacing="0" cellpadding="0" style="width:100%;">
					<tr>
						<th>INDIRIZZO</th>
						<% if Session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE") then %><th>CATEGORIA</th><% end if %>
						<% if Session("NEXTCOM_ATTIVA_GESTIONE_ATTIVITA") then %><th>CAMPAGNA MARKETING</th><% end if %>
					</tr>
					<tr>
						<td class="content_center" style="width:50%;">
							<input type="text" name="search_address" value="<%= replace(session(Prefix & "address"), """", "&quot;") %>" style="width:99%;">
						</td>
						<% if Session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE") then %>
							<td class="content_center">
								<% sql = " SELECT icat_id, icat_nome_it FROM tb_indirizzario_categorie " &_
										 " ORDER BY icat_nome_it "
								CALL dropDown(conn, sql, "icat_id", "icat_nome_it", "search_categorie", Session(Prefix & "categorie"), false, "style=""width:99%;""", Session("LINGUA"))%>	  
							</td>
						<% end if %>
						<% if Session("NEXTCOM_ATTIVA_GESTIONE_ATTIVITA") then %>
							<td class="content_center">
								<% sql = " SELECT inc_id, inc_nome FROM tb_indirizzario_campagne " &_
										 " ORDER BY inc_nome "
								CALL dropDown(conn, sql, "inc_id", "inc_nome", "search_campagne", Session(Prefix & "campagne"), false, "style=""width:99%;""", Session("LINGUA"))%>	  
							</td>
						<% end if %>
					</tr>
				</table>
			</td>
		</tr>
    </table>
    <table cellspacing="1" cellpadding="0" class="tabella_madre">
	    <caption>
            <% if MultipleSelection then %>
                <span style="float:right;">
                    <a id="tutti" class="button_L2" href="javascript:void(0);" onclick="Tutti()">
					    <%= ChooseValueByAllLanguages(Session("LINGUA"), "SELEZIONA TUTTI", "SELECT ALL", "", "", "", "", "", "")%>
					</a>
					&nbsp;
					<a id="nessuno" class="button_L2" href="javascript:void(0);" onclick="Reset()">
					    <%= ChooseValueByAllLanguages(Session("LINGUA"), "DESELEZIONA TUTTI", "DESELECT ALL", "", "", "", "", "", "")%>
					</a>
                </span>
            <% end if %>
		    <%= ChooseValueByAllLanguages(Session("LINGUA"), "Trovati n&ordm; " & Pager.recordcount & " record in n&ordm; " & Pager.PageCount & " pagine ", Pager.recordcount & " records found in " & Pager.PageCount & " pages ", "", "", "", "", "", "")%>
        </caption>
        <% if not rs.eof then %>
		    <tr>
			    <th class="center" style="width:5%;"><%= ChooseValueByAllLanguages(Session("LINGUA"), "SCEGLI", "CHOOSE", "", "", "", "", "", "")%></th>
                <% if instr(1, ListType, "LOGIN", vbTextCompare)>0 then %>
                    <th><%= ChooseValueByAllLanguages(Session("LINGUA"), "CONTATTO", "CONTACT", "", "", "", "", "", "")%></th>
                    <th colspan="1">LOGIN</th>
                <% else %>
    				<th colspan="1"><%= ChooseValueByAllLanguages(Session("LINGUA"), "CONTATTO", "CONTACT", "", "", "", "", "", "")%></th>
                <% end if %>
				<th colspan="">INDIRIZZO</th>
				<th>&nbsp;</th>
			</tr>
       	<% 	rs.AbsolutePage = Pager.PageNo
		   	dim type_cntRel, type_emailMandatory, type_login, type_email, type_cell
			dim type_faxMandatory,type_fax,type_cellMandatory
			type_cntRel = instr(1, ListType, "CNTREL", vbTextCompare) > 0
			type_emailMandatory = instr(1, ListType, "EMAILMANDATORY", vbTextCompare)>0
			type_cellMandatory = instr(1, ListType, "CELLMANDATORY", vbTextCompare)>0
			type_faxMandatory = instr(1, ListType, "FAXMANDATORY", vbTextCompare)>0
			type_login = instr(1, ListType, "LOGIN", vbTextCompare)>0
			type_email = instr(1, ListType, "EMAIL", vbTextCompare)>0
			type_cell = instr(1, ListType, "CELL", vbTextCompare)>0
			type_fax = instr(1, ListType, "FAX", vbTextCompare)>0
			while not rs.eof and rs.AbsolutePage = Pager.PageNo
                ListaValori = ""
                if type_email then
					if cString(request("contatti_email_newsletter")) = "true" AND cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then
						sql = SQL_isTrue(conn, "email_newsletter")
					else
						sql = SQL_isTrue(conn, "email_Default")
					end if
                    'recupera email del contatto
                    sql = " SELECT ValoreNumero, email_newsletter FROM tb_ValoriNumeri " &_
					      " WHERE id_TipoNumero=" & VAL_EMAIL & " AND " + sql + _
						  "	AND NOT " + SQL_IsNull(conn, "ValoreNumero") + " AND ValoreNumero<>'' " + _
						  " AND id_Indirizzario=" & rs("IdElencoIndirizzi")
                    rsE.open sql , conn, adOpenStatic, adLockOptimistic, adCmdText
                            
                    while not rsE.eof
					    if IsEmail(cString(rsE("ValoreNumero"))) then
							ListaValori = ListaValori & rsE("ValoreNumero")
							if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) AND cBoolean(rsE("email_newsletter"), false) then
								ListaValori = ListaValori & "&nbsp;" + get_icona_newsletter(true)
							end if
						else
						    ListaValori = ListaValori & "<span class=""alert"">" & rsE("ValoreNumero") & " NON VALIDA!</span>"
						end if
					    rsE.movenext
					    if not rsE.eof then ListaValori = ListaValori & ", "
                    wend
					rsE.close
					CALL WriteContact(rs, ListaValori,VAL_EMAIL, false)
                end if
				if type_cell then
                    'recupera cellulare del contatto
                    sql = " SELECT ValoreNumero FROM tb_ValoriNumeri " &_
					      " WHERE id_TipoNumero=" & VAL_CELLULARE & " AND " + SQL_isTrue(conn, "email_Default") + _
						  "	AND NOT " + SQL_IsNull(conn, "ValoreNumero") + " AND ValoreNumero<>'' " + _
						  " AND id_Indirizzario=" & rs("IdElencoIndirizzi")
                    rsE.open sql , conn, adOpenStatic, adLockOptimistic, adCmdText
                            
                    while not rsE.eof
					    if IsPhoneNumber(rsE("ValoreNumero")) then
							ListaValori = ListaValori & "<span title=""" & FormatMobilePhone(rsE("ValoreNumero")) & """>" & rsE("ValoreNumero") & "</span>"
						else
						    ListaValori = ListaValori & "<span class=""alert"">" & rsE("ValoreNumero") & " NON VALIDO!</span>"
						end if
					    rsE.movenext
					    if not rsE.eof then ListaValori = ListaValori & ", "
                    wend
					rsE.close
					CALL WriteContact(rs, ListaValori,VAL_CELLULARE, false)
                end if
				
				if type_fax then
                    'recupera fax del contatto
                    sql = " SELECT ValoreNumero FROM tb_ValoriNumeri " &_
					      " WHERE id_TipoNumero=" & VAL_FAX & " AND " + SQL_isTrue(conn, "email_Default") + _
						  "	AND NOT " + SQL_IsNull(conn, "ValoreNumero") + " AND ValoreNumero<>'' " + _
						  " AND id_Indirizzario=" & rs("IdElencoIndirizzi")
                    rsE.open sql , conn, adOpenStatic, adLockOptimistic, adCmdText
                            
                    while not rsE.eof
					    if IsPhoneNumber(rsE("ValoreNumero")) then
							ListaValori = ListaValori & rsE("ValoreNumero")
						else
						    ListaValori = ListaValori & "<span class=""alert"">" & rsE("ValoreNumero") & " NON VALIDO!</span>"
						end if
					    rsE.movenext
					    if not rsE.eof then ListaValori = ListaValori & ", "
                    wend
					rsE.close
					CALL WriteContact(rs, ListaValori,VAL_FAX, false)
                end if
				
				if not (type_email OR type_cell OR type_fax) then
					CALL WriteContact(rs, ListaValori,VAL_NONE, false)
				end if
				
		 		if type_cntRel AND NOT type_login then
					if cString(request("contatti_email_newsletter")) = "true" AND cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then
						sql = SQL_isTrue(conn, "email_newsletter")
					else
						sql = SQL_isTrue(conn, "email_Default")
					end if
					sql = " SELECT * FROM (tb_indirizzario i"& _
						  " INNER JOIN tb_cnt_lingue l ON i.lingua = l.lingua_codice)"& _
						  " LEFT JOIN tb_valoriNumeri v ON i.idElencoIndirizzi = v.id_Indirizzario"& _
						  " WHERE cntRel = "& rs("idElencoIndirizzi") & " AND " & sql
					if type_cell then
						sql = sql &	 " AND id_TipoNumero=" & VAL_CELLULARE & " "
					end if
					if type_email then
						sql = sql &	 " AND id_TipoNumero=" & VAL_EMAIL & " "
					end if
					if type_fax then
						sql = sql &	 " AND id_TipoNumero=" & VAL_FAX & " "
					end if
					sql = sql &	  " ORDER BY modoRegistra"
						  
					rse.open sql, conn, adOpenStatic, adLockOptimistic
					if not rse.eof then %>
						<tr>
							<td class="content_center" rowspan="2">&nbsp;</td>
							<th class="l2" class="content" colspan="<%= IIF(type_login, "4", "3") %>"><%= ChooseValueByAllLanguages(Session("LINGUA"), "contatti interni:", "inboard contacts:", "", "", "", "", "", "")%></th>
						</tr>
						<tr>
							<td colspan="<%= IIF(type_login, "4", "3") %>">
								<table cellpadding="0" cellspacing="1" style="width: 100%;">
				<%				while not rse.eof
									if instr(1, ListType, "EMAIL", vbTextCompare)>0 then
										if cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) AND cBoolean(rse("email_newsletter"), false) then
											CALL WriteContactExt(rse, rse("valoreNumero"),VAL_EMAIL, true, true)
										else
											CALL WriteContact(rse, rse("valoreNumero"),VAL_EMAIL, true)
										end if
									elseif instr(1, ListType, "CELL", vbTextCompare)>0 then
										CALL WriteContact(rse, rse("valoreNumero"),VAL_CELLULARE, true)
									elseif instr(1, ListType, "FAX", vbTextCompare)>0 then
										CALL WriteContact(rse, rse("valoreNumero"),VAL_FAX, true)
									end if
									rse.movenext
								wend %>
								</table>
							</td>
						</tr>
		<%			end if
					rse.close
				end if %>
        <% 		rs.moveNext
			wend%>
			<tr>
			    <td colspan="5" class="footer" style="text-align:left;">
			        <% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
		        </td>
			</tr>
        <% else%>
		    <tr><td class="noRecords"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Nessun record trovato", "No record found", "", "", "", "", "", "")%></th></tr>
		<% end if %>
	</table>
</form>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>
</body>
</html>
<%
rs.close
conn.close
set rs = nothing
set rsE = nothing
set conn = nothing


'restituisce true se il contatto e selezionabile listType dovrà essere uno dei tipi di numero possibili
Function IsSelectable(value,listType)
	'response.write "IsSelectable(" & value & "," & listType & ")"
	if listType=VAL_EMAIL then
		IsSelectable = NOT type_email OR NOT type_emailMandatory OR value <> ""
	elseif listType=VAL_CELLULARE then
		IsSelectable = NOT type_cell OR NOT type_cellMandatory OR value <> ""
	elseif listType=VAL_FAX then
		IsSelectable = NOT type_fax OR NOT type_faxMandatory OR value <> ""
	else
		IsSelectable = true
	end if
End Function


Sub WriteContact(rs, value, typeOfVal, interni)
	CALL WriteContactExt(rs, value, typeOfVal, interni, false)
end sub

'scrive le colonne per la scelta del contatto
Sub WriteContactExt(rs, value, typeOfVal, interni, is_email_newsletter)
	dim selectable, ID
	'ID in output
	if typeLoginId then
		ID = rs("ut_id")
	else
		ID = rs("IdElencoIndirizzi")
	end if
	
	selectable = IsSelectable(value,typeOfVal) %>
	<tr>
		<td class="content_center">
			<input type="hidden" name="nome_<%= rs("IdElencoIndirizzi") %>" id="nome_<%= rs("IdElencoIndirizzi") %>" value="<%= ContactFullName(rs) %>">
			<% if Selectable then %>
				<input class="checkbox" name="selezione" id="selezione_<%= rs("IdElencoIndirizzi") %>"
					   type="<%= IIF(MultipleSelection, "checkbox", "radio") %>"
					   value="<%= ID %>"
					   onclick="seleziona_contatto(this, '<%= JSReplacerEncode(ContactFullName(rs)) %>')">
				<script type="text/javascript">
					document.getElementById("selezione_<%= rs("IdElencoIndirizzi") %>").checked = opener.<%= FieldName %>_is_selected_contatto(<%= ID %>);
				</script>
			<% else %>
				<input type="<%= IIF(MultipleSelection, "checkbox", "radio") %>" class="checkbox_disabled" disabled>
			<% end if %>
		</td>
		<td class="<%= IIF(Selectable, "content", "content_disabled") %>">
			<%= ContactFullName(rs) %>
			<% if ((typeOfVal = VAL_EMAIL) or (typeOfVal = VAL_CELLULARE) or (typeOfVal = VAL_FAX))  and value<>"" then %>
					<span class="note">( <%= value %>&nbsp;<%CALL write_icona_newsletter(is_email_newsletter)%>)</span>
			<% elseif typeOfVal <> VAL_NONE then  %>
					<span class="note warning">( <%=IIF((typeOfVal = VAL_CELLULARE) or (typeOfVal = VAL_FAX),"Numero ","E-Mail") %><%=IIF(cBoolean(cString(request("contatti_email_newsletter")), false)," per NEWSLETTER", "")%> non presente ) </span>
			<% end if %>
		</td>
		<% if type_login then %>
		<td class="<%= IIF(Selectable, "content", "content_disabled") %>"><%= uCase(rs("ut_login")) %></td>
		<% end if %>
		<td class="<%= IIF(Selectable, "content", "content_disabled") %>"><%= ContactAddress(rs) %></td>
		<td class="content_center" style="width:<%= IIF(interni, "27px;padding-left: 1px", "6%") %>;"><img src="../grafica/flag_mini_<%= rs("lingua") %>.jpg" alt="Lingua: <%= rs("lingua_nome_IT") %>"></td>
	</tr>
	<%
End Sub


%>