<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<%
dim conn, rs, rsE, sql, Pager, checked, visualizza, emails
dim nome, field_id, field_nome, input_vis, input_hid, rubriche_visibili

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsE = Server.CreateObject("ADODB.RecordSet")

'recupera rubriche visibili all'utente
rubriche_visibili = GetList_Rubriche(conn, rs)
set Pager = new PageNavigator

'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	if request("tutti")<>"" then
		Session("cnt_rubriche") = ""
		Session("cnt_denominazione") = ""
	elseif request("cerca")<>"" then
		Session("cnt_rubriche") = request("search_rubriche")
		Session("cnt_denominazione") = request("search_denominazione")
	end if
end if

sql = " SELECT * FROM tb_indirizzario WHERE " & _
	  " IDElencoIndirizzi IN (SELECT id_indirizzo FROM rel_rub_ind WHERE id_rubrica IN ("

'filtra sulle rubriche
if Session("cnt_rubriche")<>"" then
	sql = sql & Session("cnt_rubriche")
else
	sql = sql & rubriche_visibili
end if
sql = sql & ")) "

if Session("cnt_denominazione")<>"" then
	sql = sql & " AND " + SQL_FullTextSearch_Contatto_Nominativo(conn, Session("cnt_denominazione"))
end if


sql = sql & " ORDER BY ModoRegistra"

input_vis = "visContatti"					'nome del campo visibile nell'opener
input_hid = "contatti"						'nome del campo nascosto nell'opener

CALL Pager.OpenSmartRecordset(conn, rs, sql, 20)
%>
<html>
	<head>
		<title>Selezione contatti</title>
		<link rel="stylesheet" type="text/css" href="../library/stili.css">
		<meta name="robots" content="noindex,nofollow" />
		<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
	</head>
<script language="JavaScript" type="text/javascript">
	function Clikka(selezionato, id, stringa) {
	<% If request.querystring("submit") = "" then 	'se vengo dalla posta %>
		if (selezionato) {
			opener.form1.<%= input_vis %>.value += stringa + ";"
			opener.form1.<%= input_hid %>.value += id + ";"
		} else {
			var re = eval('/' + stringa + ';/g');
			opener.form1.<%= input_vis %>.value = opener.form1.<%= input_vis %>.value.replace(re, '')
			re = eval('/' + id + ';/g');
			opener.form1.<%= input_hid %>.value = opener.form1.<%= input_hid %>.value.replace(re, '')
		}
	<% Else											'vengo dalle attivita %>
		opener.form1.hid_cliente_id.value = id
		opener.form1.submit()
		window.close()
	<% End If %>
	}
	
	function Tutti() {
		for(var i=0; i < form1.elements.length; i++)
			if (form1.elements(i).id.substring(0, 4) == "chk_" && !form1.elements(i).checked)
				form1.elements(i).click()
	}
	
	function Reset() {
		for(var i=0; i < form1.elements.length; i++)
			if (form1.elements(i).id.substring(0, 4) == "chk_" && form1.elements(i).checked)
				form1.elements(i).click()
	}
</script>
<body topmargin="9" onload="window.focus()">

<form action="" method="post" name="form1">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td>
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						<table width="100%" border="0" cellspacing="0">
							<tr>
								<td class="caption">Opzioni di ricerca</td>
								<td align="right" style="padding-right:5px;"><a class="button" href="javascript:void(0);" onclick="window.close();" >CHIUDI</a></td>
							</tr>
						</table>
					</caption>
					<tr>
						<th>RUBRICA</th>
						<th colspan="3">NOME / DENOMINAZIONE</th>
					</tr>
					<tr>
						<td class="content_center" width="34%">
							<% sql = " SELECT id_rubrica, nome_rubrica FROM tb_rubriche " &_
								 	 " WHERE id_rubrica IN (" & rubriche_visibili & ") " &_
									" ORDER BY nome_rubrica"
							CALL dropDown(conn, sql, "id_rubrica", "nome_rubrica", "search_rubriche", Session("cnt_rubriche"), false, _
										  "style=""width:100%;""", LINGUA_ITALIANO)%>	  
						</td>
						<td class="content_center" width="40%">
							<input type="text" name="search_denominazione" value="<%= replace(session("cnt_denominazione"), """", "&quot;") %>" style="width:100%;">
						</td>
						<td class="content_center" style="vertical-align:middle;">
							<input type="submit" name="cerca" value="CERCA" class="button" style="width:100%; height:18px;">
						</td>
						<td class="content_center" style="vertical-align:middle;">
							<input type="submit" name="tutti" value="TUTTI" class="button" style="width:100%; height:18px;">
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr><td style="font-size:5px;">&nbsp;</td></tr>
		<tr>
			<td>
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						<table cellpadding="0" cellspacing="0" width="100%">
						<tr>
							<td class="caption">Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</td>
							<td align="right">
								<a id="tutti" class="button_L2" href="javascript:void(0);" onclick="Tutti()">
									SELEZIONA TUTTI
								</a>
								&nbsp;
								<a id="nessuno" class="button_L2" href="javascript:void(0);" onclick="Reset()">
									DESELEZIONA TUTTI
								</a>
							</td>
						</tr>
						</table>
					</caption>
					<% if not rs.eof then %>
						<tr>
							<th class="center" style="width:5%;">SCEGLI</th>
							<th>CONTATTO</th>				
						</tr>
						<%	sql = "SELECT ValoreNumero FROM tb_ValoriNumeri " &_
								  "WHERE id_TipoNumero=6 AND id_Indirizzario="
							rs.AbsolutePage = Pager.PageNo
							while not rs.eof and rs.AbsolutePage = Pager.PageNo
								if instr(";"& request("ELENCO"), ";"& rs("IDElencoIndirizzi") & ";") > 0 then
									checked = "checked"
								else
									checked = ""
								end if
								if rs("isSocieta") then
									visualizza = rs("NomeOrganizzazioneElencoIndirizzi")
								else
									visualizza = rs("CognomeElencoIndirizzi") &" "& rs("NomeElencoIndirizzi")
								end if
								visualizza = replace(visualizza, "^", "")
								visualizza = replace(visualizza, "'", "")
								visualizza = replace(visualizza, """", "")
								visualizza = replace(visualizza, "(", "")
								visualizza = replace(visualizza, ")", "")
						%>
							<tr>
								<td class="content_center">
								<% if request.querystring("submit") = "" then %>
									<input <%= checked %> type="Checkbox" id="chk_<%= rs("IDElencoIndirizzi") %>" name="chk_<%= rs("IDElencoIndirizzi") %>" onclick="Clikka(this.checked, '<%= rs("IDElencoIndirizzi") %>', '<%= visualizza %>')" class="checkbox">
								<% Else %>
									<input type="Radio" id="chk_<%= rs("IDElencoIndirizzi") %>" name="chk_<%= rs("IDElencoIndirizzi") %>" onclick="Clikka(this.checked, '<%= rs("IDElencoIndirizzi") %>', '<%= visualizza %>')" class="noborder">
								<% End If %>
								</td>
								<% Emails = GetValueList(conn, rsE, sql & rs("IDElencoIndirizzi"))%>
								<td class="content">
									<% If rs("IsSocieta") then %>
										<%= rs("NomeOrganizzazioneElencoIndirizzi") %>
									<% Else  %>
										<%= rs("CognomeElencoIndirizzi") &" "& rs("NomeElencoIndirizzi") %>
									<% End If 
									if Emails<>"" then%>
										<span class="note">(<%= Emails %>)</span>
									<%end if%>
								</td>
							</tr>
							<% rs.moveNext
						wend%>
						<tr>
							<td colspan="5" class="footer" style="text-align:left;">
								<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
							</td>
						</tr>
					<%else%>
						<tr><td class="noRecords">Nessun record trovato</th></tr>
					<% end if %>
				</table>
			</td>
		</tr>
	</table>
</form>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set rsE = nothing
set conn = nothing%>
