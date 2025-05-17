<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_template_accesso, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoTemplateSalva.asp")
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - templates - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoTemplate.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, i, lingua
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("WEB_TEMPLATE_SQL"), "id_page", "SitoTemplateMod.asp")
end if

sql = "SELECT * FROM tb_pages WHERE id_page="& cIntero(request("ID"))
set rs = conn.Execute(sql)
%>

<script language="JavaScript" type="text/javascript">
	function strumenti(lingua){
		OpenAutoPositionedWindow("SitoPagineStrumenti.asp?TEMPLATE=<%= request("ID") %>&LINGUA=" + lingua, "strumenti", 500, 250)
	}
	
	function pagine(){
		OpenAutoPositionedScrollWindow("SitoTemplatePagine.asp?ID=<%= request("ID") %>", "pagine", 600, 600, true)
	}
</script>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati del template</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="template precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="template successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">DATI DEL TEMPLATE</th></tr>
		<tr>
			<td class="label" style="width:10%;">titolo:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_nomepage" value="<%= rs("nomepage") %>" maxlength="250" size="100">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label" style="width:40%;" colspan="2">apri il template per la modifica</td>
			<td class="content_center">
				<a class="button_l2_block" href="javascript:void(0);" style="width:100px;" onclick="OpenAutoPositionedScrollWindow('loadshock.asp?PAGINA=<%= rs("id_page") %>', 'editor', document.body.clientWidth, screen.height, true);">
					MODIFICA
				</a>
			</td>
			<td class="content" style="width:45%;">&nbsp;</td>
		</tr>
		<tr>
			<td class="label" colspan="2">visualizza il template</td>
			<td class="content_center">
				<a class="button_l2_block" href="dynalay.asp?PAGINA=<%= rs("id_page") %>&lingua=it" target="_blank" style="width:100px;">
					VEDI
				</a>
			</td>
			<td class="content">&nbsp;</td>
		</tr>
		<tr>
			<td class="label" colspan="2">apri in una nuova finestra gli strumenti del template</td>
			<td class="content_center">
				<a class="button_l2_block" href="javascript:void(0);" onClick="strumenti('<%=lingua%>')" style="width:100px;">
					APRI
				</a>
			</td>
			<td class="content">&nbsp;</td>
		</tr>
		<tr>
			<td class="label" colspan="2">visualizza le pagine associate al template</td>
			<td class="content_center">
				<a class="button_l2_block" href="javascript:void(0);" onclick="pagine()" title="pagine associate al template" style="width:100px;">
					VEDI
				</a>
			</td>
			<td class="content">&nbsp;</td>
		</tr>
		<% 
		if not Session("SITO_MOBILE") then
		%>
		<tr><th class="L2" colspan="4">PRORIET&Agrave;</th></tr>
		<tr>
			<td class="label" rowspan="2">tipo:</td>
			<td class="content" colspan="3">
				<input type="radio" class="checkbox" name="tfn_semplificata" value="0" <%= Chk(not rs("semplificata")) %>>
				per pagina normale
			</td>
		</tr>
		<tr>
			<td class="content" colspan="3">
				<table cellpadding="0" cellspacing="0">
					<tr>
						<td><input type="radio" class="checkbox" name="tfn_semplificata" value="1" <%= Chk(rs("semplificata")) %>></td>
						<td style="padding-right:4px;"><img src="../grafica/notReadKnow.gif" border="0" alt="Template per email con visualizzazione semplificata."></td>
						<td>per email semplificate</td>
					</tr>
				</table>
			</td>
		</tr>
		<% 
		end if
		'parte di form per la pubblicazione dei dati di creazione e modifica
		CALL Form_DatiModifica_EX(conn, rs, "page_", "dati del record", "L2")
		%>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva_torna" value="SALVA & TORNA AD ELENCO" style="width:22%">
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>

<%
set rs = nothing
conn.Close
set conn = nothing
%>