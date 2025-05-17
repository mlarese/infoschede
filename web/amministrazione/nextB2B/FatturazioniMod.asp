<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("FatturazioniSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione tipologie di fatturazioni - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Fatturazioni.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_FATTURAZIONI_SQL"), "fatt_id", "FatturazioniMod.asp")
end if

sql = "SELECT * FROM gtb_fatturazioni WHERE fatt_id="& cIntero(request("ID"))
set rs = conn.Execute(sql)
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati della tipologia di fatturazione</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="fatturazione precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="fatturazione successiva" <%= ACTIVE_STATUS %>>
							SUCCESSIVA &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="3">DATI DELLA TIPOLOGIA DI FATTURAZIONE</th></tr>
		<tr>
			<td class="label">codice:</td>
			<td class="content" style="width:45%;">
				<input type="text" class="text" name="tft_fatt_codice" value="<%= rs("fatt_codice") %>" maxlength="255" size="50">
				(*)
			</td>
			<td class="note">
				Codice della tipologia di fatturazione.
			</td>
		</tr>
		<tr>
			<td class="label">numero corrente:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tfn_fatt_numero_corrente" value="<%= rs("fatt_numero_corrente") %>" maxlength="5" size="5">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">data corrente:</td>
			<td class="content" colspan="2">
				<table>
				<tr>
					<td class="content">
						<% CALL WriteDataPicker_Input("form1", "tfd_fatt_data_corrente", rs("fatt_data_corrente"), " disabled ", "/", true, true, LINGUA_ITALIANO) %>				
					</td>
					<td>&nbsp;(*)</td>
				</tr>
				</table>	
			</td>
		</tr>
		<tr>
			<td class="label">serie:</td>
			<td class="content">
				<input type="text" class="text" name="tft_fatt_serie" value="<%= rs("fatt_serie") %>" maxlength="10" size="10">
				(*)
			</td>
			<td class="note">
				Suffisso da apporre al numero di fatturazione.
			</td>
		</tr>		
		<tr>
			<td class="footer" colspan="3">
				(*) Campi obbligatori.
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