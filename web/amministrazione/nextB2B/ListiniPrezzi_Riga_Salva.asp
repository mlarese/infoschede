<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE file="ListiniPrezzi_Tools.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<%

dim conn, rs, rsl, rsv, sql, listino, sql_where
dim prezzo_base, prezzo_attuale, prezzo_nuovo, iva_id, scontoQ_id, var_sconto, var_euro, visibile, promozione, personalizzato, variante_id, offerta_dal, offerta_al
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")
set rsl = Server.CreateObject("ADODB.RecordSet")

listino = cInteger(request.querystring("ID"))

sql = " SELECT *, (SELECT listino_id FROM gtb_listini WHERE listino_base_attuale=1) AS LB_ATTUALE " + _
	  " FROM gtb_listini WHERE listino_id=" & listino
rsl.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

if request("update_row")<>"" AND request("prz_variante_id")<>"" then
	'salva la riga singola
	variante_id = request("prz_variante_id")
	dim data_dal, data_al
	data_dal = request.form("offerta_dal_" & variante_id)
	if cString(data_dal) = "" then data_dal = "azzera"
	data_al = request.form("offerta_al_" & variante_id)
	if cString(data_al) = "" then data_al = "azzera"
	CALL Listino_SalvaRiga(conn, rsl, rs, listino, variante_id, _
			   			   cReal(request.form("prz_var_sconto_" & variante_id)), _
						   cReal(request.form("prz_var_euro_" & variante_id)), _
						   cInteger(request.form("prz_iva_id_" & variante_id)), _
						   cInteger(request.form("prz_scontoQ_id_" & variante_id)), _
						   request.form("vis_" & variante_id)<>"", _
						   request.form("promo_" & variante_id)<>"", _
						   cReal(request.form("prz_prezzo_" & variante_id)), _
						   request.form("prz_personalizzato_" & variante_id)<>"", _
						   data_dal, _
						   data_al)
end if
%>

<%'--------------------------------------------------------
sezione_testata = "registrazione modifica riga listino" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>
<div id="content_ridotto">
	<%if cInteger(variante_id)<>0 then
			sql_where = " AND prz_variante_id=" & cIntero(request("prz_variante_id"))
			CALL Listino_OpenRowRecordset(rsl, rs, listino, sql_where)
			CALL Listino_StatoRiga(rs, rsl, listino, prezzo_base, prezzo_nuovo, _
										    prezzo_attuale, var_sconto, var_euro, personalizzato, _
										    iva_id, scontoQ_id, visibile, promozione, offerta_dal, offerta_al)
			CALL Listino_Scheda(rsl, listino, false) %>
			<table cellspacing="1" cellpadding="0" class="tabella_madre">
				<tr>
					<td class="label">articolo</td>
					<td class="content" colspan="3">
						<% CALL ArticoloLink(rs("rel_art_id"), rs("art_nome_it"), rs("rel_cod_int"))
						if rs("art_varianti") then %>
							<%= ListValoriVarianti(conn, rsv, rs("rel_id")) %>
						<% end if %>
					</td>
				</tr>
				<tr>
					<td class="label">prezzo </td>
					<td class="content" colspan="3"><%= formatPrice(prezzo_attuale, 2, true) %> &euro;</td>
				</tr>
				<tr>
					<td class="label" rowspan="2">variazioni</td>
					<td class="label">in euro</td>
					<td class="content" colspan="2"><%= formatPrice(var_euro, 2, true) %> &euro;</td>
				</tr>
				<tr>
					<td class="label" style="width:20%;">in percentuale</td>
					<td class="content" colspan="2"><%= formatPrice(var_sconto, 2, true) %> %</td>
				</tr>
				<% if scontoQ_id>0 then 
					sql = "SELECT scc_nome FROM gtb_scontiQ_classi WHERE scc_id=" & scontoQ_id%>
					<tr>
						<td class="label">sconto per quantit&agrave; </td>
						<td class="content" colspan="3"><%= GetValueList(conn, rsv, sql) %>%</td>
					</tr>
				<% end if %>
				<tr>
					<td class="label" <%= IIF(rsl("listino_offerte"),"rowspan=""3""","")%>><%= IIF(rsl("listino_offerte"),"in offerta","visibile")%> </td>
					<td class="content" colspan="3">
						<input type="checkbox" class="checkbox" disabled <%= chk(visibile) %>>
						<% if rsl("listino_offerte") and visibile then %>	
							<span class="Icona Offerte" title="articolo attualmente in offerta speciale">&nbsp;</span>
						<% end if %>
					</td>
				</tr>
				<% if not rsl("listino_offerte") then %>
					<tr>
						<td class="label" style="width:20%;">in promozione </td>
						<td class="content" colspan="3">
							<input type="checkbox" class="checkbox" disabled <%= chk(promozione) %>>
							<% if cInteger(rs("OFFERTA"))>0 then %>
								<span class="Icona Offerte" title="articolo attualmente in offerta speciale">&nbsp;</span>
							<% end if %>
						</td>
					</tr>
				<% else %>
					<tr>
						<td class="label" style="width:20%;">dal</td>
						<td class="content" colspan="2"><%=rs("prz_offerta_dal")%></td>
					</tr>
					<tr>
						<td class="label" style="width:20%;">al</td>
						<td class="content" colspan="2"><%=rs("prz_offerta_al")%></td>
					</tr>
				<% end if %>
				<tr>
					<td class="footer" colspan="4">
						<input style="width:23%;" type="button" class="button" name="chiudi" value="CHIUDI" onclick="window.close();">
					</td>
				</tr>
			</table>
		<% rs.close
	else %>
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<tr>
				<td class="errore">ERRORE NELL'APERTURA DELLA FINESTRA: CONTATTARE L'AMMINISTRATORE.</td>
			</tr>
		</table>
	<% end if %>
</div>
</body>
</html>
<%rsl.close
conn.close
set rs = nothing
set rsl = nothing
set rsv = nothing
set conn = nothing %>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>