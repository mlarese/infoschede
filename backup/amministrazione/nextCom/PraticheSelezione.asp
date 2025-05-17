<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<%
dim Pager
set Pager = new PageNavigator
if request("FIELD")<>"" then
	Pager.Reset()
	Session("PraSel_SELECTED") = cInteger(request("selected"))
	Session("PraSel_FIELD") = request("FIELD")
	Session("PraSel_aperte") = "A"
	Session("PraSel_chiuse") = ""
	response.redirect "PraticheSelezione.asp"
end if

'--------------------------------------------------------
sezione_testata = "Selezione della pratica" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
dim conn, sql, rs, rsr
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")


if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	if request("tutti")<>"" then
		Session("PraSel_rubrica") = ""
		Session("PraSel_contatto") = ""
		Session("PraSel_pratica") = ""
		Session("PraSel_aperte") = ""
		Session("PraSel_chiuse") = ""
	elseif request("cerca")<>"" then
		Session("PraSel_rubrica") = request("search_rubrica")
		Session("PraSel_contatto") = request("search_contatto")
		Session("PraSel_pratica") = request("search_pratica")
		Session("PraSel_aperte") = request("search_aperte")
		Session("PraSel_chiuse") = request("search_chiuse")
	end if
end if


sql = "SELECT IDElencoIndirizzi, IsSocieta, NomeOrganizzazioneElencoIndirizzi, NomeELencoIndirizzi, CognomeElencoIndirizzi, " & _
	  " ModoRegistra, pra_id, pra_nome, pra_archiviata " & _
	  " FROM tb_pratiche INNER JOIN tb_Indirizzario ON tb_pratiche.pra_cliente_id = tb_indirizzario.IDElencoIndirizzi " & _
	  " WHERE  (pra_creatore_id = "& Session("ID_ADMIN") &" OR "& AL_query(conn, AL_PRATICHE) &") "

if Session("PraSel_rubrica")<>"" then
	sql = sql & " AND IDElencoIndirizzi IN (SELECT id_indirizzo FROM rel_rub_ind WHERE id_rubrica =" & Session("PraSel_Rubrica") & ")"
end if	  	  

if Session("PraSel_aperte")<>"" XOR Session("PraSel_chiuse")<>"" then
	if Session("PraSel_aperte") <> "" then
		sql = sql & " AND NOT " & SQL_IsTrue(conn, "pra_archiviata")
	end if
	if Session("PraSel_chiuse") <> "" then
		sql = sql & " AND " & SQL_IsTrue(conn, "pra_archiviata")
	end if
end if

if Session("PraSel_pratica")<>"" then
	sql = sql & " AND " + SQL_FullTextSearch(Session("PraSel_pratica"), "pra_nome")
end if

if Session("PraSel_contatto")<>"" then
	sql = sql & " AND " + SQL_FullTextSearch_Contatto_Nominativo(conn, Session("PraSel_contatto"))
end if

sql = sql & " ORDER BY pra_nome, ModoRegistra"

CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)
%>
<script language="JavaScript" type="text/javascript">
	function Selezione(objPraId, objPraNome, objCnt){
		opener.form1.<%= Session("PraSel_FIELD") %>.value = objPraId.value;
		opener.form1.contatto.value = objCnt.value;
		opener.form1.pratica.value = objPraNome.value;
		window.close();
	}
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
		<th>RUBRICA CONTATTO</th>
		<th>NOME CONTATTO</th>
		<th>NOME PRATICA</th>
		<th colspan="2">STATO PRATICHE</th>
	</tr>
	<tr>
		<td class="content" width="32%">
			<% sql = " SELECT id_rubrica, nome_rubrica FROM tb_rubriche " &_
					 " WHERE id_rubrica IN (" & GetList_Rubriche(conn, rsr) & ") " &_
					 " ORDER BY nome_rubrica"
			CALL dropDown(conn, sql, "id_rubrica", "nome_rubrica", "search_rubrica", Session("PraSel_rubrica"), false, _
						  "style=""width:100%;""", LINGUA_ITALIANO)%>	  
		</td>
		<td class="content" width="21%">
			<input type="text" name="search_contatto" value="<%= replace(session("PraSel_contatto"), """", "&quot;") %>" style="width:100%;">
		</td>
		<td class="content" width="21%">
			<input type="text" name="search_pratica" value="<%= replace(session("PraSel_pratica"), """", "&quot;") %>" style="width:100%;">
		</td>
		<td class="content pratiche">
			<input type="Checkbox" name="search_aperte" class="checkbox" value="A" <%= IIF(Session("PraSel_aperte")<>"", " checked", "") %>>
			<strong>aperte</strong> 
		</td>
		<td class="content">
			<input type="Checkbox" name="search_chiuse" class="checkbox" value="A" <%= IIF(Session("PraSel_chiuse")<>"", " checked", "") %>>
			archiviate 
		</td>
	</tr>
</table>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption class="border">Elenco pratiche</caption>
	<tr>
		<td>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
				<tr>
					<td class="label_no_width" colspan="3">
						<% if rs.eof then %>
							Nessuna pratica trovata.
						<% else %>
							Trovate n&ordm; <%= Pager.recordcount %> pratiche in n&ordm; <%= Pager.PageCount %> pagine
						<% end if %>
					</td>
				</tr>
				<% if not rs.eof then %>
					<tr>
						<th class="L2">SEL.</th>
						<th class="L2">PRATICA</th>
						<th class="L2">CONTATTO</th>
					</tr>
					<%rs.AbsolutePage = Pager.PageNo
					while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
						<input type="hidden" name="PRA_<%= rs("pra_id") %>" value="<%= replace(cString(rs("pra_nome")), """", "'") %>">
						<input type="hidden" name="CNT_<%= rs("pra_id") %>" value="<%= replace(ContactFullName(rs), """", "'") %>">
						<tr>
							<td width="4%" class="content_center">
								<input type="radio" name="seleziona" class="checkbox" value="<%= rs("pra_id") %>" <%= Chk(Session("PraSel_SELECTED") = rs("pra_id")) %>
									   title="Click per selezionare la pratica"	
									   onclick="Selezione(this, PRA_<%= rs("pra_id") %>, CNT_<%= rs("pra_id") %>)">
							</td>
							<td width="40%" class="content<%= IIF(rs("pra_archiviata"), "", "pratiche") %>">
								<% PraticaLinkedName(rs) %>
							</td>
							<td class="content">
								<%  ContactLinkedName(rs) %>
							</td>
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
</body>
</html>
<% 
rs.close
conn.close
set rs = nothing
set rsr = nothing
set conn = nothing
%>