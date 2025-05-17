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
		if (window.opener.document.getElementById("invia_anche_contatti_interni").checked)
		{
			url += '&anche_contatti_interni=true'
		}
		window.location.href = url;
	</script>
<% end if %>


<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<%'--------------------------------------------------------
sezione_testata = "Selezione rubriche" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim conn, rs, rsc, sql, Pager
dim rubriche_visibili, messageType
dim nContatti, nContattiValidi, Title

dim input_vis, input_hid
input_vis = "visRubriche"		'nome del campo visibile nell'opener
input_hid = "rubriche"			'nome del campo nascosto nell'opener
if cString(request("input_hid")) <> "" then
	input_hid = cString(request("input_hid"))
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

messageType = request("MessageType")

'recupera rubriche visibili all'utente
rubriche_visibili = GetList_Rubriche(conn, rs)

set Pager = new PageNavigator

sql = "SELECT * FROM tb_rubriche WHERE id_rubrica IN (" & rubriche_visibili & ") ORDER BY nome_rubrica"
CALL Pager.OpenSmartRecordset(conn, rs, sql, 30)
%>
<script language="JavaScript" type="text/javascript">
	function Clikka(selezionato, id, stringa) {
		if (selezionato) {
			window.opener.document.form1.<%= input_vis %>.value += stringa +';';			
			window.opener.document.form1.<%= input_hid %>.value += ' '+ id +';';
		} else {
			//var re = eval('/ ' + stringa + ';/g');
			var re = stringa + ';';
			window.opener.document.form1.<%= input_vis %>.value = window.opener.document.form1.<%= input_vis %>.value.replace(re, '');
			//re = eval('/ ' + id + ';/g');
			re = ' ' + id + ';';
			window.opener.document.form1.<%= input_hid %>.value = window.opener.document.form1.<%= input_hid %>.value.replace(re, '');
		}
	}
	
	function ClikkaTutti(){
		for(var i=0; i < form1.elements.length; i++)
		{
			if (document.form1.elements[i].id.substring(0, 4) == "chk_"){
				if (!document.form1.elements[i].checked){
					document.form1.elements[i].checked = true;
					document.form1.elements[i].onclick();
				}
			}
		}
	}
	
	function ClikkaNessuno(){
		for(var i=0; i < form1.elements.length; i++)
		{
			if (document.form1.elements[i].id.substring(0, 4) == "chk_"){
				if (document.form1.elements[i].checked){
					document.form1.elements[i].checked = false;
					document.form1.elements[i].onclick();
				}
			}
		}
	}
	
</script>
<div id="content_ridotto">
<form action="" method="post" name="form1" id="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption class="border">
			Trovate n&ordm; <%= Pager.recordcount %> rubriche in n&ordm; <%= Pager.PageCount %> pagine
		</caption>
		<% if not rs.eof then %>
			<tr>
				<td class="content_right" colspan="4">
					<a class="button_L2" href="javascript:void(0)" onclick="ClikkaTutti()">
						SELEZIONA TUTTE
					</a>
					&nbsp;&nbsp;
					<a class="button_L2" href="javascript:void(0)" onclick="ClikkaNessuno()">
						DESELEZIONA TUTTE
					</a>
				</td>
			</tr>
			<tr>
				<th class="center" style="width:8%;" rowspan="2">SCEGLI</th>
				<th rowspan="2">RUBRICA</th>
				<th class="center" colspan="2" style="border-bottom:0px;">CONTATTI</th>
			</tr>
			<tr>
				<th class="center">ASSOCIATI</th>
				<th class="center"><%= Comunicazioni_LabelByType(messageType, "con email", "con fax", "con cellulare") %>
				<%
				if cString(request("contatti_email_newsletter")) = "true" AND cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then
					response.write "<br>per NEWSLETTER&nbsp;<img src=""../grafica/i.p.new.gif""/>"
				end if
				%>
				</th>
			</tr>
			<%	rs.AbsolutePage = Pager.PageNo
				while not rs.eof and rs.AbsolutePage = Pager.PageNo 
					sql = "SELECT COUNT(*) FROM rel_rub_ind WHERE id_rubrica=" & rs("id_rubrica")
					nContatti = cInteger(GetValueList(conn, rsc, sql))
					
					dim sql_where
					if cString(request("contatti_email_newsletter")) = "true" AND cBoolean(Session("ATTIVA_RECAPITI_NEWSLETTER"), false) then
						sql_where = SQL_isTrue(conn, "email_newsletter")
					else
						sql_where = SQL_isTrue(conn, "email_Default")
					end if
					'sql = " SELECT COUNT(*) FROM rel_rub_ind INNER JOIN tb_ValoriNumeri ON rel_rub_ind.id_indirizzo = tb_ValoriNumeri.id_Indirizzario " + _
					'	  " WHERE id_rubrica=" & rs("id_rubrica") & " AND tb_ValoriNumeri.id_TipoNumero=" & Comunicazioni_LabelByType(messageType, VAL_EMAIL, VAL_FAX, VAL_CELLULARE) & _
					'	  " AND " & sql_where
					sql = " SELECT COUNT(*) FROM tb_Indirizzario INNER JOIN " & _
						  " tb_ValoriNumeri ON tb_Indirizzario.IDElencoIndirizzi = tb_ValoriNumeri.id_Indirizzario " & _
						  " WHERE (IDElencoIndirizzi IN (SELECT id_indirizzo FROM rel_rub_ind WHERE id_rubrica=" & rs("id_rubrica") & ") "
						  if cString(request("anche_contatti_interni")) = "true" then
							sql = sql & " OR cntRel IN (SELECT id_indirizzo FROM rel_rub_ind WHERE id_rubrica=" & rs("id_rubrica") & ") "
						  end if
						  sql = sql & ") AND tb_ValoriNumeri.id_TipoNumero=" & Comunicazioni_LabelByType(messageType, VAL_EMAIL, VAL_FAX, VAL_CELLULARE) & _
									  " AND " & sql_where
					
					nContattiValidi = cInteger(GetValueList(conn, rsc, sql))
					
					if nContatti > 0 AND nContattiValidi>0 then
						Title = ""
					else
						Title = "Rubrica non selezionabile perch&egrave; " + _
								IIF(nContatti = 0, "non ha contatti associati", _
									"non ha alcun contatto con un " + Comunicazioni_LabelByType(messageType, "indirizzo email", "numero di fax", "numero di cellulare")) + _
								"valido."
					end if %>
				<tr>
					<td class="content_center">
						<% if nContatti > 0 AND nContattiValidi>0 then %>
							<input onblur="this.checked = ((window.opener.document.form1.<%= input_hid %>.value.indexOf(' <%= rs("id_rubrica") %>;')) >= 0);" type="Checkbox" name="chk_<%= rs("id_rubrica") %>" id="chk_<%= rs("id_rubrica") %>" onclick="Clikka(this.checked, '<%= rs("id_rubrica") %>', ' <%= JSReplacerEncode(rs("nome_Rubrica")) %>')" class="checkbox">
						<% else %>
							<input type="Checkbox" class="checkbox" disabled title="<%= Title %>">
						<% end if %>
					</td>
					<td class="content<%= IIF(nContatti > 0 AND nContattiValidi>0, "", "_disabled") %>" title="<%= Title %>"><%= rs("nome_rubrica") %></td>
					<td class="content_center<%= IIF(nContatti > 0 AND nContattiValidi>0, "", "_disabled") %>"><%= nContatti %></td>
					<td class="content_center<%= IIF(nContatti > 0 AND nContattiValidi>0, "", "_disabled") %>"><%= nContattiValidi %></td>
				</tr>
				<% rs.moveNext
			wend%>
			<tr>
				<td colspan="4" class="footer" style="text-align:left;">
					<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
				</td>
			</tr>
		<%else%>
			<tr><td class="noRecords">Nessun record trovato</th></tr>
		<% end if %>
	</table>
</form>
</div>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
	
	for(var i=0; i < form1.elements.length; i++)
	{
		if (document.form1.elements[i].id.substring(0, 4) == "chk_")
			document.form1.elements[i].onblur();
	}
</script>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set rsc = nothing
set conn = nothing%>