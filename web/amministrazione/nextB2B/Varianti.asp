<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione varianti - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA VARIANTE"
dicitura.link_new = "Tabelle.asp;VariantiNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

'controllo che non ci siano articoli correlati prima di cancellare
'escludo la variante base se gestita dal sys
session("B2B_VARIANTI_SQL") = "SELECT *, "& _
							  "(SELECT COUNT(*) FROM grel_art_vv INNER JOIN gtb_valori ON grel_art_vv.rvv_val_id=gtb_valori.val_id WHERE val_var_id=var_id) AS N_ART "& _
							  "FROM gtb_varianti ORDER BY var_nome_it"
rs.open session("B2B_VARIANTI_SQL"), conn, adOpenStatic, adLockReadOnly, adCmdText 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco varianti - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th class="center" style="width:5%;">ID</th>
			<th>NOME</th>
			<th class="center" width="15%">ORDINE</th>
			<th class="center" colspan="3" style="width: 20%;">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("var_id") %></td>
				<td class="content"><%= rs("var_nome_it") %></td>
				<td class="content_center"><%= rs("var_ordine") %></td>
				<td class="Content_center">
					<a class="button" href="VariantiMod.asp?ID=<%= rs("var_id") %>">
						MODIFICA
					</a>
				</td>
				<td class="Content_center">
					<% if rs("N_ART") > 0 then %>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare la variante: sono presenti dei valori associati" <%= ACTIVE_STATUS %>>
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('VARIANTI','<%= rs("var_id") %>');" >
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