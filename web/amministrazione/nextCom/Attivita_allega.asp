<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<%
dim conn, rs, rsE, sql, Pager, checked, visualizza, emails
dim nome, field_id, field_nome, input_vis, input_hid, rubriche_visibili


'--------------------------------------------------------
sezione_testata = "Selezione degli allegati" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 


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
		Session("all_nome") = ""
		Session("all_pratica") = ""
		Session("all_allegati") = ""
		Session("all_contatto") = ""
	elseif request("cerca")<>"" then
		Session("all_nome") = request("all_nome")
		Session("all_pratica") = request("all_pratica")
		Session("all_allegati") = request("all_allegati")
		Session("all_contatto") = request("all_contatto")
	end if
elseif Session("COM_ATT_PRATICA") <> "" then
	Session("all_pratica") = conn.Execute("SELECT pra_nome FROM tb_pratiche WHERE pra_id="& Session("COM_ATT_PRATICA"))
end if

sql = " SELECT * FROM (tb_documenti d "& _
	  " LEFT JOIN tb_pratiche p ON d.doc_pratica_id=p.pra_id) "&_
	  " LEFT JOIN tb_indirizzario i ON p.pra_cliente_id=i.IDElencoIndirizzi "& _
	  " WHERE "& AL_query(conn, AL_DOCUMENTI)

'filtra su nome
if Session("all_nome") <> "" then
	    sql = sql & " AND " & SQL_FullTextSearch(Session("all_nome"), "doc_nome")
end if
'filtra su pratica
if Session("all_pratica") <> "" then
    sql = sql & " AND " & SQL_FullTextSearch(Session("all_pratica"), "pra_nome")
end if
'filtra su contatto
if Session("all_contatto") <> "" then
    sql = sql & " AND " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("all_contatto"))
end if
'filtra su allegati
if Session("all_allegati") <> "" then
	sql = sql & " AND doc_id IN (SELECT rel_documento_id FROM rel_documenti_files r "& _
						   		"INNER JOIN tb_files f ON r.rel_files_id=f.f_id "& _
								"WHERE " + SQL_FullTextSearch(session("all_pratica"), "f_original_name") + ")"
end if

sql = sql & " ORDER BY doc_nome"

input_vis = "visDoc"						'nome del campo visibile nell'opener
input_hid = "documenti"						'nome del campo nascosto nell'opener

CALL Pager.OpenSmartRecordset(conn, rs, sql, 20)
%>

<script language="JavaScript" type="text/javascript">
	function Clikka(selezionato, id, stringa) {
		if (selezionato) {
			opener.form1.<%= input_vis %>.value += stringa + ";"
			opener.form1.<%= input_hid %>.value += id + ";"
		} else {
			var re
			opener.form1.<%= input_vis %>.value = opener.form1.<%= input_vis %>.value.replace(stringa +';', '')
			re = eval('/' + id + ';/g');
			opener.form1.<%= input_hid %>.value = opener.form1.<%= input_hid %>.value.replace(re, '')
		}
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

<div id="content_ridotto">
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
						<th>CONTATTO</th>
						<th>PRATICA</th>
						<th>NOME</th>
						<th colspan="3">ALLEGATI</th>
					</tr>
					<tr>
						<td class="content_center" width="20%">
							<input type="text" name="all_contatto" value="<%= replace(session("all_contatto"), """", "&quot;") %>" style="width:100%;">
						</td>
						<td class="content_center" width="20%">
							<input type="text" name="all_pratica" value="<%= replace(session("all_pratica"), """", "&quot;") %>" style="width:100%;">
						</td>
						<td class="content_center" width="20%">
							<input type="text" name="all_nome" value="<%= replace(session("all_nome"), """", "&quot;") %>" style="width:100%;">
						</td>
						<td class="content_center" width="20%">
							<input type="text" name="all_allegati" value="<%= replace(session("all_allegati"), """", "&quot;") %>" style="width:100%;">
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
							<th>DOCUMENTO</th>
							<th>PRATICA</th>
							<th>CONTATTO</th>
						</tr>
						<%	rs.AbsolutePage = Pager.PageNo
							while not rs.eof and rs.AbsolutePage = Pager.PageNo
								if instr(";"& request("ELENCO"), ";"& rs("doc_id") & ";") > 0 then
									checked = "checked"
								else
									checked = ""
								end if %>
							<tr>
								<td class="content_center">
									<input <%= checked %> type="Checkbox" id="chk_<%= rs("doc_id") %>" name="chk_<%= rs("doc_id") %>" onclick="Clikka(this.checked, '<%= rs("doc_id") %>', '<%= rs("doc_nome") %>')" class="checkbox">
								</td>
								<td class="content" nowrap>
									<% DocLinkedName(rs) %>
								</td>
								<td class="content<%= IIF(rs("pra_archiviata"), "", " pratiche") %>">
									<% PraticaLinkedName(rs) %>
								</td>
								<td class="content">
									<% ContactLinkedName(rs) %>
								</td>
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
			</td>
		</tr>
	</table>
</form>
</div>
<% 
rs.close
conn.close 
set rs = nothing
set rsE = nothing
set conn = nothing%>
