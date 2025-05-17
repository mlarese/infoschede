<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione liste codici - elenco"
dicitura.puls_new = "NUOVA LISTA DI CODICI"
dicitura.link_new = "ListeCodiciNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT *, (SELECT COUNT(*) FROM gtb_rivenditori WHERE riv_lstcod_id=lstCod_id) AS N_RIV, " + _
	  " (SELECT COUNT(*) FROM gtb_codici WHERE cod_lista_id = gtb_lista_codici.lstcod_id) AS N_CODICI " + _
	  " FROM gtb_lista_codici ORDER BY LstCod_Nome"
session("B2B_LSTCOD_SQL") = sql
Session("B2B_SQL_CODICI") =""
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco liste codici - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th class="center" style="width:5%;">ID</th>
			<th width="10%">CODICE</th>
			<th>NOME</th>
			<th class="center" style="width:15%;">DEI CLIENTI</th>
			<th style="width:16%;">N&ordm; personalizzati</th>
			<th colspan="3" class="center" style="width: 25%;">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("lstCod_id") %></td>
				<td class="content"><%= rs("LstCod_cod") %></td>
				<td class="content"><%= rs("LstCod_nome") %></td>
				<td class="content_center">
					<input disabled type="checkbox" name="clienti" value="1" class="checkbox" <%= chk(not rs("lstCod_sistema")) %>>
				</td>
				<td class="content_center">
					<%= cIntero(rs("N_CODICI")) %>
				</td>
				<td class="Content_center">
					<a class="button" href="ListeCodiciCodici.asp?ID=<%= rs("lstCod_id") %>">
						CODICI
					</a>
				</td>
				<td class="Content_center">
					<a class="button" href="ListeCodiciMod.asp?ID=<%= rs("lstCod_id") %>">
						MODIFICA
					</a>
				</td>
				<td class="Content_center">
					<% if rs("N_RIV") > 0 then %>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare la lista di codici: sono presenti rivenditori associati">
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('LISTA_CODICI','<%= rs("lstCod_id") %>');" >
							CANCELLA
						</a>
					<% end if %>
				</td>
			</tr>
			<%rs.movenext
		wend%>
	</table>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set conn = nothing%>