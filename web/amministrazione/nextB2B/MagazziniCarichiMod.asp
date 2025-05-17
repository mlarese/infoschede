<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("MagazziniCarichiSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
 

dim conn, rs, rsd, sql, aux, disabled,magazzinoID
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.Recordset")
set aux = Server.CreateObject("ADODB.Recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_CARICHI_SQL"), "car_ID", "MagazziniCarichiMod.asp")
end if

sql = "SELECT * FROM gtb_carichi INNER JOIN "& _
	  "gtb_magazzini ON gtb_carichi.car_magazzino_id = gtb_magazzini.mag_id "& _
	  "WHERE car_id="& cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
magazzinoID = rs("mag_id")
disabled = ""
if rs("car_movimentato") then
	disabled = "readonly"
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione carichi magazzino - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "MagazziniCarichi.asp?IDMAG=" & IIF(request.form("tfn_car_magazzino_id")<>"",request.form("tfn_car_magazzino_id"),magazzinoID)
dicitura.scrivi_con_sottosez()

%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica del carico a magazzino</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="stato precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="stato successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">DATI DEL CARICO</th></tr>
		<tr>
			<td class="label">codice fornitore:</td>
			<td class="content">
				<input type="text" class="text" name="tft_car_fornitore_cod" value="<%= rs("car_fornitore_cod") %>" maxlength="50" size="10" <% If rs("car_movimentato") then %>readonly<% end if  %>> (*)
			</td>
			<td class="label">fornitore:</td>
			<td class="content">
				<input type="text" class="text" name="tft_car_fornitore" value="<%= rs("car_fornitore") %>" maxlength="50" size="50" <% If rs("car_movimentato") then %>readonly<% end if  %>>
			</td>
		</tr>
		<tr>
			<td class="label">stato:</td>
			<td class="content" colspan="3">
				<input type="checkbox" class="checkbox"  name="chk_car_movimentato" <%= chk(rs("car_movimentato")) %>
				<% If rs("car_movimentato") then %>disabled<% end if  %>>Movimentato
				<% If rs("car_movimentato") then %>
					<input type="hidden" name="chk_car_movimentato" value="<%= rs("car_movimentato") %>">
				<% end if  %>
			</td>
		</tr>
		
		<tr>
		<% If rs("car_movimentato") then %>
			<td class="label">data di carico:</td>
			<td class="content">
				<table cellpadding="0" cellspacing="0">
				<tr>
					<td><%= DataEstesa(rs("car_data"),"IT") %></td>
					<td>&nbsp;(*)</td>
				</tr>
				</table>
			</td>
			<td class="label">magazzino:</td>
			<td class="content">
				<%=rs("mag_nome")  %> (*)
			</td>
		<% Else  %>
			<td class="label">data di carico:</td>
			<td class="content">
				<table cellpadding="0" cellspacing="0">
				<tr>
					<td><% CALL WriteDataPicker_Input("form1", "tfd_car_data", rs("car_data"), "", "/", true, true, LINGUA_ITALIANO) %></td>
					<td>&nbsp;(*)</td>
				</tr>
				</table>
			</td>
			<td class="label">magazzino:</td>
			<td class="content">
				<% CALL dropDown(conn, "SELECT * FROM gtb_magazzini ORDER BY mag_nome", "mag_id", "mag_nome", "tfn_car_magazzino_id", rs("mag_id"), true, "", LINGUA_ITALIANO) %> (*)
			</td>
		<% End If %>
		</tr>
		<tr><th colspan="4">ARTICOLI ARRIVATI</th></tr>
		<tr>
			<td colspan="4">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<% 	sql = "SELECT * FROM gv_carichi WHERE (rcv_car_id = "& cIntero(request("ID")) &")"
						rsd.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText%>
					<tr>
						<td class="label" colspan="5" style="width:80%;">
							<% if rsd.eof then %>
								nessun dettaglio definito per il carico.
							<% else %>
								n&ordm; <%= rsd.recordcount %> dettagli per l'ordine.
							<% end if %>
						</td>
						<% 	if disabled = "" then %>
						<td colspan="2" class="content_right" style="padding-right:0px;">
							<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ArticoliSelPrz.asp?CAR_ID=<%= request("ID") %>', 'CaricoDettagli', 600, 600, true)"
							   title="apre la finestra per l'inserimento di un nuovo dettaglio del carico" <%= ACTIVE_STATUS %>>
								NUOVO DETTAGLIO
							</a>
						</td>
						<% 	end if %>
					</tr>
					<% 	if not rsd.eof then
							dim ind_id %>
						<tr>
							<th class="L2">CODICE INT.</th>
							<th class="L2">CODICE PR.</th>
							<th class="L2" width="50%">NOME</th>
							<th class="L2">VARIANTE</th>
							<th class="L2">QUANTIT&Agrave;</th>
					<% 		if disabled = "" then %>
							<th class="l2_center" width="16%" colspan="2">OPERAZIONI</th>
					<% 		end if %>
						</tr>
					<%		while not rsd.eof 
								%>
							<tr>
								<td class="content"><%= rsd("rel_cod_int") %></td>
								<td class="content"><%= rsd("rel_cod_pro") %></td>
								<td class="content"><%= rsd("art_nome_it") %></td>
								<td class="content"><%= ListValoriVarianti(conn, aux, rsd("rel_id")) %></td>
								<td class="content"><%= rsd("rcv_qta") %></td>
					<% 			if disabled = "" then %>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('MagazziniCarichi_dettagliMod.asp?ID=<%= rsd("rcv_id") %>', 'CaricoDettagli', 510, 400, true)">
										MODIFICA
									</a>
								</td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('DETORD','<%= rsd("rcv_id") %>');">
										CANCELLA
									</a>
								</td>
					<% 			end if %>
							</tr>
							<%	rsd.MoveNext
							wend
						end if
						rsd.close%>
				</table>
			</td>
		</tr>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="3" name="tft_car_note" <%= disabled %>><%= rs("car_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="hidden" name="IDMAG" value="<%=magazzinoID %>">
				<% If rs("car_movimentato") then %>
					<a class="button" href="<%= dicitura.link_new %>" title="Torna indietro" onmouseover="return(status=this.title);" onmouseout="status=''; " >
							INDIETRO
						</a>
				<% Else  %>
				<input type="submit" style="width:23%;" class="button" name="mod" value="SALVA & TORNA ALL'ELENCO">
				<input type="submit" class="button" name="salva" value="SALVA">
				<% end if  %>
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>

<%
rs.close
set rs = nothing
set rsd = nothing
set aux = nothing
conn.Close
set conn = nothing
%>