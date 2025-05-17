<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
dim Pager, destinazione_mode
set Pager = new PageNavigator

if request("field_nome") = "destinazione" then
	destinazione_mode = true
else
	destinazione_mode = false
end if

'--------------------------------------------------------
if destinazione_mode then
	sezione_testata = "Selezione dell'indirizzo" 
else 
	sezione_testata = "Selezione dell'anagrafica" 
end if	
%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede_Const.asp" -->
<%'----------------------------------------------------- 
dim conn, sql, rs, rsr, sql_filtri, admin_id
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("cli_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("cli_")
	end if
end if

if cString(Session("LAST_FIELD_NOME")) <> request("field_nome") then
	Pager.Reset()
	CALL SearchSession_Reset("cli_")
	Session("LAST_FIELD_NOME") = request("field_nome")
end if

if request("CENTRO_ASSISTENZA_ID") <> "" then
	'se l'utente è centro assistenza (od officina) mostro solo i clienti inseriti dell'utente stesso o i clienti associati alle schede collegate all'utente.
	sql = " AND ((cnt_insAdmin_id = "&CIntero(Session("ID_ADMIN"))&" OR cnt_modAdmin_id = "&CIntero(Session("ID_ADMIN"))&") " & _
		  " OR (cnt_insAdmin_id IN (SELECT ag_admin_id FROM gtb_agenti WHERE ag_id = "&cIntero(request("CENTRO_ASSISTENZA_ID"))&") " & _
		  " OR IDElencoIndirizzi IN (SELECT ut_NextCom_id FROM tb_utenti WHERE ut_id IN " & _
		  " 							(SELECT sc_cliente_id FROM sgtb_schede WHERE sc_centro_assistenza_id = "&cIntero(request("CENTRO_ASSISTENZA_ID"))&")))) "

end if



'filtra per nome
if Session("cli_nome")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("cli_nome"))
end if

'filtra per login
if Session("cli_login")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("cli_login"), "ut_login")
end if

'filtra per indirizzo
if Session("cli_indirizzo")<>"" AND destinazione_mode then
	sql = sql & " AND " & SQL_FullTextSearch_Contatto_Indirizzo(conn, Session("cli_indirizzo"))
end if

sql_filtri = sql

if cString(Trim(request.querystring("filtro_profilo")))<>"" then
	sql = " AND riv_profilo_id "
	if cBoolean(request.querystring("filtro_exclude"), false) then
		sql = sql & " NOT "
	end if
	sql = sql & " IN (" & ParseSQL(request.querystring("filtro_profilo"), adChar) & ") " & sql_filtri
end if

sql = " SELECT * FROM "&IIF(destinazione_mode,"tb_indirizzario","gv_rivenditori") & _
	  " WHERE (1=1) "& sql & _
	  " ORDER BY ModoRegistra"
	  
CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)
%>

<script language="JavaScript" type="text/javascript">
	<% if destinazione_mode then %>
		function SelezioneDestinazione(ObjId, ObjDestinazione){
			opener.document.form1.<%= request.querystring("field_id") %>.value = ObjId.value;
			opener.document.form1.<%= request.querystring("field_nome") %>.value = ObjDestinazione.value;
			window.close();
		}
	<% else %>
		function Selezione(ObjId, ObjNome, objDestId, ObjDestinazione){
			// ObjId			id rivenditore
			// ObjNome			nome contatto
			// objDestId		id del contatto corrispondente (IDElencoIndirizzi)
			// ObjDestinazione	indirizzo
			opener.document.form1.<%= request.querystring("field_id") %>.value = ObjId.value;
			opener.document.form1.<%= request.querystring("field_nome") %>.value = ObjNome.value;
			<% if request.querystring("field_destinazione_id") <> "" then %>
				opener.document.form1.<%= request.querystring("field_destinazione_id") %>.value = objDestId.value;
				opener.document.form1.<%= request.querystring("field_destinazione_nome") %>.value = ObjDestinazione.value;
				opener.document.form1.submit();
			<% elseif request.querystring("AFTER")="onchange" then %>
				opener.document.form1.<%= request.querystring("field_id") %>.onchange();
			<% elseif request.querystring("AFTER")="submit" then %>
				opener.document.form1.submit();
			<% end if %>
			window.close();
		}
	<% end if %>
</script>
<div id="content_ridotto">
<form action="" method="post" id="ricerca" name="ricerca">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption>
		<table border="0" cellspacing="0" cellpadding="1" align="right">
			<tr>
				<td style="font-size: 1px; padding-right:1px;" nowrap>
					<input type="submit" name="cerca" value="CERCA" class="button">
					&nbsp;
					<input type="submit" name="tutti" value="VEDI TUTTI" class="button">
				</td>
			</tr>
		</table>
		Opzioni di ricerca
	</caption>
	<tr>
		<th>NOME CONTATTO</th>
		<th>LOGIN CONTATTO</th>
	</tr>
	<tr>
		<td class="content">
			<input type="text" name="search_nome" value="<%= TextEncode(session("cli_nome")) %>" style="width:100%;">
		</td>
		<td class="content">
			<input type="text" name="search_login" value="<%= TextEncode(session("cli_login")) %>" style="width:100%;">
		</td>
	</tr>
	<% if destinazione_mode then %>
		<tr><th colspan="2">INDIRIZZO</th></tr>
		<tr>
			<td class="content" colspan="2">
				<input type="text" name="search_indirizzo" value="<%= replace(session("cli_indirizzo"), """", "&quot;") %>" style="width:100%;">
			</td>
		</tr>
	<% end if %>
</table>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption class="border">Elenco <%=IIF(destinazione_mode,"destinazioni", "anagrafiche")%></caption>
	<tr>
		<td>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
				<tr>
					<td class="label_no_width" colspan="3">
						<% if rs.eof then %>
							Nessuna cliente trovato.
						<% else %>
							Trovati n&ordm; <%= Pager.recordcount %> clienti in n&ordm; <%= Pager.PageCount %> pagine
						<% end if %>
					</td>
				</tr>
				<% if not rs.eof then %>
					<tr>
						<th class="L2">SEL.</th>
						<% if not destinazione_mode then %>
							<th class="L2">CONTATTO</th>
						<% end if %>
						<th class="L2">INDIRIZZO</th>
					</tr>
					<%rs.AbsolutePage = Pager.PageNo
					while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
						<tr>
							<% if destinazione_mode then %>
								<td width="4%" class="content_center">
									<input type="hidden" name="DESTINAZIONE_<%= rs("IDElencoIndirizzi") %>" value="<%= ContactAddress(rs) %>">
									<input type="radio" name="seleziona" class="checkbox" value="<%= rs("IDElencoIndirizzi") %>" <%= Chk(CInteger(request.querystring("selected")) = rs("IDElencoIndirizzi")) %>
										   title="Click per selezionare il cliente"	
										   onclick="SelezioneDestinazione(this, ricerca.DESTINAZIONE_<%= rs("IDElencoIndirizzi") %>)">
								</td>
							<% else %>
								<td width="4%" class="content_center">
									<input type="hidden" name="NAME_<%= rs("IDElencoIndirizzi") %>" value="<%= ContactFullName(rs) %>">
									<input type="hidden" name="DEST_ID_<%= rs("IDElencoIndirizzi") %>" value="<%= rs("IDElencoIndirizzi") %>">
									<input type="hidden" name="DESTINAZIONE_<%= rs("IDElencoIndirizzi") %>" value="<%= ContactAddress(rs) %>">
									<input type="radio" name="seleziona" class="checkbox" value="<%= rs("riv_id") %>" <%= Chk(CInteger(request.querystring("selected")) = rs("riv_id")) %>
										   title="Click per selezionare il cliente"	
										   onclick="Selezione(this, ricerca.NAME_<%=rs("IDElencoIndirizzi")%>, ricerca.DEST_ID_<%= rs("IDElencoIndirizzi") %>, ricerca.DESTINAZIONE_<%= rs("IDElencoIndirizzi") %>)">
								</td>
							<% end if %>
							<% if not destinazione_mode then %>
								<td class="content">
									<a href="javascript:void(0);" title="apri scheda del cliente" <%= ACTIVE_STATUS %>
										onclick="OpenAutoPositionedScrollWindow('ClientiGestione.asp?ID=<%= rs("IDElencoIndirizzi") %>', 'cliente', 760, 400, true);">
										<%= ContactFullName(rs) %>
									</a>
								</td>
							<% end if %>
							<td class="content"><%= ContactAddress(rs)%></td>
						</tr>
						<% rs.MoveNext
					wend%>
					<tr>
						<td colspan="3" class="footer">
							<table width="100%" cellpadding="0" cellspacing="0">
								<tr>
									<td><% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%></td>
									<td align="right">
										<a class="button" href="javascript:window.close();" title="chiudi la finestra" <%= ACTIVE_STATUS %>>
											CHIUDI</a>
									</td>
								</tr>
							</table>
							
						</td>
					</tr>
				<% end if %>
			</table>
		</td>
	</tr>
</table>
</form>
</div>
<% if request("BUTTONS_ADD") = "true" then %>
	<div id="pulsanti" style="position:absolute; top:580px; left:4px; width:99%;">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Crea nuova anagrafica</caption>
			<tr><th>SCEGLI LA CATEGORIA</th></tr>
			<tr>
				<td class="content" style="padding:2px;">
					<a nowrap target="SelezioneCliente" class="button_L2" href="ClientiGestione.asp?PROFILO=anagrafiche_clienti&IDPROFILO=<%=CLIENTI_PRIVATI%>&STANDALONE=true&field_id=<%=request("field_id")%>" style="margin-left:4px;">
						CLIENTE PRIVATO
					</a>
					<a nowrap target="SelezioneCliente" class="button_L2" href="ClientiGestione.asp?PROFILO=anagrafiche_clienti&IDPROFILO=<%=CLIENTI_PROFESSIONALI%>&STANDALONE=true&field_id=<%=request("field_id")%>" style="margin-left:4px;">
						CLIENTE PROFESSIONALE
					</a>
					<a nowrap target="SelezioneCliente" class="button_L2" href="ClientiGestione.asp?PROFILO=anagrafiche_clienti&IDPROFILO=<%=RIVENDITORI%>&STANDALONE=true&field_id=<%=request("field_id")%>" style="margin-left:4px;">
						RIVENDITORE
					</a>
					<a nowrap target="SelezioneCliente" class="button_L2" href="ClientiGestione.asp?PROFILO=anagrafiche_clienti&IDPROFILO=<%=SUPERVISORI_NEGOZI%>&STANDALONE=true&field_id=<%=request("field_id")%>" style="margin-left:4px;">
						SUPERVISORE NEGOZI
					</a>
				</td>

			</tr>
		</table>
	</div>
<% end if %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>

</body>
</html>
<% 
rs.close
conn.close
set rs = nothing
set rsr = nothing
set conn = nothing
%>