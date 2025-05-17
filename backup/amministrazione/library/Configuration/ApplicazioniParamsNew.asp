<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ApplicazioniParamsSalva.asp")
end if
%>
<!--#INCLUDE FILE="../ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="../../nextPassport/ToolsApplicazioni.asp" -->
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
dicitura.sezione = "Gestione parametri - nuovo"
dicitura.scrivi_con_sottosez()

dim conn, sql, i, rs
set conn = Server.CreateObject("ADODB.Connection")
conn.open GetConfigurationConnectionstring()
set rs = Server.CreateObject("ADODB.RecordSet")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_sid_personalizzato" value="0">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo parametro degli applicativi</caption>
		<tr><th colspan="4">DATI DEL PARAMETRO</th></tr>
		<%for i=lbound(LINGUE_CODICI) to ubound(LINGUE_CODICI)%>
			<tr>
				<% if i = 0 then %>
					<td class="label" rowspan="<%= ubound(LINGUE_CODICI)+1 %>">nome:</td>
				<% end if %>
				<td class="content" colspan="3">
					<img src="../../grafica/flag_<%= LINGUE_CODICI(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_sid_nome_<%= LINGUE_CODICI(i) %>" value="<%= request.form("tft_sid_nome_"& LINGUE_CODICI(i)) %>" maxlength="255" size="75">
				<% 	if LINGUE_CODICI(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_sid_codice" value="<%= request.form("tft_sid_codice") %>" maxlength="50" size="50">
				(*)
			</td>
			<td class="label">visibile agli utenti:</td>
			<td class="content" style="white-space: nowrap;">
				<input type="radio" class="checkbox" value="0" name="tfn_sid_admin" <%= Chk(request("tfn_sid_admin")<>"" AND cInteger(request("tfn_sid_admin"))=0) %>>
				si
				<input type="radio" class="checkbox" value="1" name="tfn_sid_admin" <%= Chk(request.servervariables("REQUEST_METHOD")<>"POST" OR cInteger(request("tfn_sid_admin"))>0) %>>
				no
			</td>
		</tr>
		<tr>
			<td class="label">tipo:</td>
			<td class="content">
				<% CALL DesDropTipi("tfn_sid_tipo", "", request.form("tfn_sid_tipo")) %>
			</td>
			<td class="label" nowrap>caratteristica principale:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_sid_principale" <%= Chk(cIntero(request("tfn_sid_principale"))>0) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_sid_principale" <%= Chk(request.servervariables("REQUEST_METHOD")<>"POST" OR cIntero(request("tfn_sid_principale"))=0) %>>
				no
			</td>
		</tr>
        <%for i=lbound(LINGUE_CODICI) to ubound(LINGUE_CODICI)%>
			<tr>
				<% if i = 0 then %>
					<td class="label" rowspan="<%= ubound(LINGUE_CODICI)+1 %>">unit&agrave; di misura:</td>
				<% end if %>
				<td class="content" colspan="3">
					<img src="../../grafica/flag_<%= LINGUE_CODICI(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_sid_unita_<%= LINGUE_CODICI(i) %>" value="<%= request.form("tft_sid_unita_"& LINGUE_CODICI(i)) %>" maxlength="255" size="20">
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">immagine:</td>
			<td class="content" colspan="3">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_sid_img", request.form("tft_sid_img"), "width:430px;", FALSE) %>
			</td>
		</tr>
		<tr>
			<td class="label">gruppo:</td>
			<td class="content" colspan="3">
				<% sql = "SELECT * FROM tb_siti_descrittori_raggruppamenti ORDER BY sdr_titolo_it"
                CALL dropDown(conn, sql, "sdr_id", "sdr_titolo_it", "tfn_sid_raggruppamento_id", request.form("tfn_sid_raggruppamento_id"), false, "", LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr><th colspan="4">APPLICAZIONI ASSOCIATE</th></tr>
		<tr>
			<td colspan="4">
			<% 	sql = " SELECT *, (NULL) AS rel"& _
					  " FROM tb_siti"& _
					  " ORDER BY sito_nome"
				CALL Write_Relations_Checker(conn, rs, sql, 2, "id_sito", "sito_nome", "rel", "car") %>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="" value="SALVA">
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
conn.close
set conn = nothing
%>