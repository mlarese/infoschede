<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("CategorieIVASalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione categorie i.v.a."
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA CATEGORIA I.V.A."
dicitura.link_new = "Tabelle.asp;CategorieIVA.asp?NEW=1"
dicitura.scrivi_con_sottosez()  

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT *, (SELECT COUNT(*) FROM gtb_articoli WHERE art_iva_id=iva_id) AS N_ART, " + _
	  " (SELECT COUNT(*) FROM gtb_prezzi WHERE prz_iva_id=iva_id) AS N_PRZ " + _
	  " FROM gtb_iva ORDER BY iva_ordine"
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Elenco categorie I.V.A. - Trovati n&ordm; <%= rs.recordcount %> records</caption>
		<tr>
			<th>NOME</th>
			<th class="center" width="12%">PERCENTUALE</th>
			<th class="center" width="8%">ORDINE</th>
			<th class="center" width="20%" colspan="2">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<tr>
				<%if cInteger(request("ID")) = rs("iva_id") then%>
					<td class="content">
						<input type="text" class="text" name="tft_iva_nome" value="<%= rs("iva_nome") %>" size="67" maxlength="250">
						(*)
					</td>
					<td class="content">
						<input type="text" class="text" name="tfn_iva_valore" value="<%= rs("iva_valore") %>" size="4" maxlength="2">
						(*)
					</td>
					<td class="content">
						<input type="text" class="text" name="tfn_iva_ordine" value="<%= rs("iva_ordine") %>" size="2" maxlength="2">
						(*)
					</td>
					<td class="content_right" style="vertical-align:middle;">
						<input type="submit" class="button" name="salva" value="SALVA">
					</td>
					<td class="content_right" style="vertical-align:middle;">
						<input type="button" class="button" name="annulla" value="ANNULLA" onclick="document.location='CategorieIva.asp';">
					</td>
				<% else %>
					<td class="content"><%= rs("iva_nome") %></td>
					<td class="content_center"><%= rs("iva_valore") %> %</td>
					<td class="content_center"><%= rs("iva_ordine") %></td>
					<td class="content_center">
						<a class="button" href="CategorieIva.asp?ID=<%= rs("iva_id") %>">
							MODIFICA
						</a>
					</td>
					<td class="content_center">
						<% if (cInteger(rs("N_ART")) + cInteger(rs("N_PRZ"))) > 0 then %>
							<a class="button_disabled" title="categoria non cancellabile perch&egrave; sono presenti prodotti o righe dei listini ad essa associati">
								CANCELLA
							</a>
						<% else %>
							<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('CAT_IVA','<%= rs("iva_ID") %>');" >
								CANCELLA
							</a>
						<% end if %>
					</td>
				<% end if %>
			</tr>
			<%rs.movenext
		wend
		if request("NEW")<>"" then%>
			<tr>
				<td class="content">
					<input type="text" class="text" name="tft_iva_nome" value="<%= request("tft_iva_nome") %>" size="67" maxlength="250">
					(*)
				</td>
				<td class="content">
					<input type="text" class="text" name="tfn_iva_valore" value="<%= request("tfn_iva_valore") %>" size="4" maxlength="2">
					(*)
				</td>
				<td class="content">
					<input type="text" class="text" name="tfn_iva_ordine" value="<%= request("tfn_iva_ordine") %>" size="2" maxlength="2">
					(*)
				</td>
				<td class="content_right" style="vertical-align:middle;">
					<input type="submit" class="button" name="salva" value="SALVA">
				</td>
				<td class="content_right" style="vertical-align:middle;">
					<input type="button" class="button" name="annulla" value="ANNULLA" onclick="document.location='CategorieIva.asp';">
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="footer" colspan="5">
				(*) Campi obbligatori.
			</td>
		</tr>
	</table>
	</form>
</div>
</body>
</html>
<%rs.close
conn.close
set rs = nothing
set conn = nothing%>