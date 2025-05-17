
<%
dim conn, rs
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

'verifica se ci sono righe in ordine o in shopping-cart di questo tipo
session(name_session_sql) = " SELECT *, "& _
							   " (SELECT COUNT(*) FROM gtb_dett_cart WHERE dett_tipo_id = gtb_dettagli_ord_tipo.dot_id ) AS RIGHE_CART, " + _
							   " (SELECT COUNT(*) FROM gtb_articoli WHERE art_dettagli_ord_tipo_id = gtb_dettagli_ord_tipo.dot_id ) AS ARTICOLI, " + _
                               " (SELECT COUNT(*) FROM gtb_dettagli_ord WHERE det_tipo_id = gtb_dettagli_ord_tipo.dot_id ) AS RIGHE_ORDINI " + _
							   " FROM gtb_dettagli_ord_tipo ORDER BY dot_nome_it"
rs.open session(name_session_sql), conn, adOpenStatic, adLockReadOnly, adCmdText 
%>

<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco tipologie - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th class="center" style="width:5%;">ID</th>
			<th>NOME</th>
			<th class="center" style="width:15%;">N. ARTICOLI</th>
            <th class="center" style="width:15%;">N. RIGHE IN ORDINE</th>
            <th class="center" style="width:19%;">N. RIGHE SHOPPING CART</th>
			<th class="center" colspan="3" style="width: 20%;">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("dot_id") %></td>
				<td class="content"><%= rs("dot_nome_it") %></td>
				<td class="content_center"><%= rs("ARTICOLI") %></td>
				<td class="content_center"><%= rs("RIGHE_ORDINI") %></td>
				<td class="content_center"><%= rs("RIGHE_CART") %></td>
				<td class="Content_center">
					<a class="button" href="OrdiniRigheTipologieMod.asp?ID=<%= rs("dot_id") %>">
						MODIFICA
					</a>
				</td>
				<td class="Content_center">
					<% if cIntero(rs("ARTICOLI")) > 0 OR cIntero(rs("RIGHE_ORDINI")) > 0 then %>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare la tipologia: presenti righe d'ordine o righe di shopping cart" <%= ACTIVE_STATUS %>>
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('ORDINI_TIPOLOGIE_RIGHE','<%= rs("dot_id") %>');" >
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