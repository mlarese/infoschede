<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
'controllo permessi
CALL CheckAutentication(session("PASS_ADMIN") <> "")

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ApplicazioniParamsSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(3)
dicitura.sottosezioni(1) = "APPLICAZIONI"
dicitura.links(1) = "Applicazioni.asp"
dicitura.sottosezioni(2) = "PARAMETRI"
dicitura.links(2) = "ApplicazioniParams.asp"
dicitura.sottosezioni(3) = "GRUPPI DI PARAMETRI"
dicitura.links(3) = "ApplicazioniParamsGruppi.asp"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "ApplicazioniParams.asp"
dicitura.sezione = "Gestione parametri - modifica"
dicitura.scrivi_con_sottosez()

dim i, conn, rs, rsc, sql, disabled
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("SQL_DESCRITTORI_APPLICATIVI"), "sid_ID", "ApplicazioniParamsMod.asp")
end if
sql = "SELECT * FROM tb_siti_descrittori WHERE sid_ID=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdtext
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati parametro degli applicativi</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="parametro precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="parametro successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">DATI DEL PARAMETRO</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 	if i = 0 then %>
			<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
		<% 	end if %>
			<td class="content" colspan="3">
				<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
				<input type="text" class="text" name="tft_sid_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("sid_nome_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
			<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
			</td>
		</tr>
		<%next%>
		<tr>
			<td class="label">codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_sid_codice" value="<%= rs("sid_codice") %>" maxlength="50" size="50">
				(*)
			</td>
			<td class="label">visibile agli utenti:</td>
			<td class="content" style="white-space: nowrap;">
				<input type="radio" class="checkbox" value="0" name="tfn_sid_admin" <%= Chk(NOT rs("sid_admin")) %>>
				si
				<input type="radio" class="checkbox" value="1" name="tfn_sid_admin" <%= Chk(rs("sid_admin")) %>>
				no
			</td>
		</tr>
		<tr>
			<td class="label">tipo:</td>
			<td class="content">
				<% CALL DesDropTipi("tfn_sid_tipo", "disabled", rs("sid_tipo")) %>
			</td>
			<td class="label" nowrap>caratteristica principale:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_sid_principale" <%= Chk(rs("sid_principale")) %>>
				si
				<input type="radio" class="checkbox" class="checkbox" value="0" name="tfn_sid_principale" <%= Chk(NOT rs("sid_principale")) %>>
				no
			</td>
		</tr>
        <%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">unit&agrave; di misura:</td>
				<% end if %>
				<td class="content" colspan="3">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_sid_unita_<%= Application("LINGUE")(i) %>" value="<%= rs("sid_unita_"& Application("LINGUE")(i)) %>" maxlength="255" size="20">
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">immagine:</td>
			<td class="content" colspan="3">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_sid_img", rs("sid_img"), "width:430px;", FALSE) %>
			</td>
		</tr>
		<tr>
			<td class="label">gruppo:</td>
			<td class="content" colspan="3">
				<% sql = "SELECT * FROM tb_siti_descrittori_raggruppamenti ORDER BY sdr_titolo_it"
                CALL dropDown(conn, sql, "sdr_id", "sdr_titolo_it", "tfn_sid_raggruppamento_id", rs("sid_raggruppamento_id"), false, "", LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr><th colspan="4">APPLICAZIONI ASSOCIATE</th></tr>
		<tr>
			<td colspan="4">
			<% 	sql = " SELECT *"& _
					  " FROM tb_siti s"& _
					  " LEFT JOIN rel_siti_descrittori r ON (s.id_sito = r.rsd_sito_id AND r.rsd_descrittore_id = "& cIntero(request("ID")) &")"& _
					  " ORDER BY sito_nome"
				CALL Write_Relations_Checker(conn, rsc, sql, 2, "id_sito", "sito_nome", "rsd_id", "car") %>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
   				(*) Campi obbligatori.
    			<input type="submit" class="button" name="salva" value="SALVA">
				<input type="submit" class="button" name="salva_elenco" value="SALVA & TORNA AD ELENCO">
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
set rsc = nothing
conn.close
set conn = nothing
%>