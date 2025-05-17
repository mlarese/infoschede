<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione magazzini - elenco"
dicitura.puls_new = "NUOVO MAGAZZINO"
dicitura.link_new = "MagazziniNew.asp"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT * , " + _
	  " (SELECT COUNT(*) FROM gtb_movimenti WHERE mov_sorg_id= gtb_magazzini.mag_id OR mov_dest_id=gtb_magazzini.mag_id) AS N_MOV, " + _
	  " (SELECT COUNT(*) FROM gtb_carichi WHERE car_magazzino_id = gtb_magazzini.mag_id) AS N_CAR, " + _
	  " (SELECT COUNT(*) FROM gtb_ordini WHERE ord_magazzino_id = gtb_magazzini.mag_id) AS N_ORD " + _
	  " FROM gtb_magazzini ORDER BY mag_nome"
session("B2B_MAG_SQL") = sql
Session("B2B_SQL_GIACENZE") =""
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco magazzini - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th class="center" style="width:12%;">CODICE</th>
			<th>DENOMINAZIONE</th>
			<th class="center" style="width: 17%;">ORDINI PUBBLICO</th>
			<th class="center" style="width: 16%;">DISP. VENDITA</th>
			<th class="center" colspan="4" style="width: 25%;">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<td class="content_center"><%= rs("mag_codice") %></td>
				<td class="content"><%= rs("mag_nome") %></td>
				<td class="content_center"><input type="checkbox" class="checkbox" <%= chk(rs("mag_vendita_pubblico")) %> disabled></td>
				<td class="content_center"><input type="checkbox" class="checkbox" <%= chk(rs("mag_disponibilita")) %> disabled></td>
				<td class="Content_center">
					<a class="button" href="MagazziniCarichi.asp?IDMAG=<%= rs("mag_id") %>">
						CARICHI
					</a>
				</td>
				<td class="Content_center">
					<a class="button" href="MagazziniInventario.asp?ID=<%= rs("mag_id") %>">
						INVENTARIO
					</a>
				</td>
				<td class="Content_center">
					<a class="button" href="MagazziniMod.asp?ID=<%= rs("mag_id") %>">
						MODIFICA
					</a>
				</td>
				<td class="Content_center">
					<% if cInteger(rs("N_MOV"))>0 OR cInteger(rs("N_CAR"))>0 OR cInteger(rs("N_ORD"))>0 then%>
						<a class="button_disabled" href="javascript:void(0);" title="Impossibile cancellare il magazzino: la merce &egrave; stata movimentata.">
							CANCELLA
						</a>
					<% else %>
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('MAGAZZINI','<%= rs("mag_id") %>');" >
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
set conn = nothing

%>
