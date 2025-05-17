<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1073741824 %>
<% response.Buffer = false %>

<!--#INCLUDE FILE="Intestazione.asp" -->
<!--#INCLUDE FILE="../nextB2B/Tools4Save_B2B.asp" -->
<% 

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Import ricambi"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Tabelle.asp"
dicitura.scrivi_con_sottosez()  


dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

dim objVariante 
set objVariante = new GestioneVariante
set objVariante.conn = conn

%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Creazione righe gtb_prezzi</caption>
		<tr><th colspan="3">ESECUZIONE</th></tr>
		<% 
		sql = "SELECT TOP 10000 rel_id FROM grel_art_valori WHERE NOT EXISTS(SELECT prz_id FROM gtb_prezzi WHERE prz_variante_id = rel_id) "
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		
		conn.beginTrans
		
		while not rs.eof
	
			objVariante.InsertDefaultRows(rs("rel_id"))
	
			rs.movenext
		wend
		rs.close

		'chiusura transazione di import
		conn.committrans 
				
		%>
		<tr>
			<td class="content_b" colspan="3">GENERAZIONE DATI COMPLETATO</td>
		</tr>
		<tr>
			<td class="footer" colspan="6" style="border-bottom:1px solid #999999;">
				<a class="button" href="default.asp">FINE</a>
			</td>
		</tr>
	</table>
</div>

<%
conn.close
set rs = nothing
set conn = nothing
%>