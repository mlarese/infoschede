<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("PromozioniSalva.asp")
end if

dim i, conn, rs, rsc, sql, disabled, lingua
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione promozioni - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Promozioni.asp"
dicitura.scrivi_con_sottosez() 

sql = "SELECT * FROM gtb_promozioni WHERE promo_ID=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdtext
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Modifica dati promozione</caption>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
			lingua = Application("LINGUE")(i)%>
				<tr>
				<% 	if i = 0 then %>
					<td colspan="2" style="width:17%;" class="label" rowspan="<%= ubound(Application("LINGUE")) + 1%>">nome:</td>
				<% 	end if %>
					<td class="content" colspan="2">
						<img src="../grafica/flag_<%= lingua %>.jpg">
						<input type="text" class="text" name="tft_promo_nome_<%= lingua %>" value="<%= rs("promo_nome_"& lingua) %>" maxlength="255" style="width:90%;">
						<% 	if lingua = LINGUA_ITALIANO then response.write "(*)" end if %>
					</td>
				</tr>
		<% next %>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
			lingua = Application("LINGUE")(i) %>
				<tr>
				<% 	if i = 0 then %>
					<td colspan="2" style="width:17%;" class="label" rowspan="<%= ubound(Application("LINGUE")) + 1%>">descrizione:</td>
				<% 	end if %>
					<td class="content" colspan="2">
						<img src="../grafica/flag_<%= lingua %>.jpg">
						<input type="text" class="text" name="tft_promo_descrizione_<%= lingua %>" value="<%= rs("promo_descrizione_"& lingua) %>" maxlength="1000" style="width:90%;">
						<% 	if lingua = LINGUA_ITALIANO then response.write "(*)" end if %>
					</td>
				</tr>
		<% next %>
		<tr>
			<td class="label">valore:</td>
			<td class="content" colspan="3">
				<input type="text" <%= disabled %> class="text" name="tfn_promo_valore" size="3" value="<%= rs("promo_valore") %>"> %
			</td>
		</tr>
		<tr>
			<td class="label">inizio validit&agrave;:</td>
			<td class="content" colspan="3">
				<% CALL WriteDataPicker_Input("form1", "tfd_promo_inizio_validita", DateIta(rs("promo_inizio_validita")), "", "/", true, true, LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">fine validit&agrave;:</td>
			<td class="content" colspan="3">
				<% CALL WriteDataPicker_Input("form1", "tfd_promo_fine_validita", DateIta(rs("promo_fine_validita")), "", "/", true, true, LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr><th colspan="4">ARTICOLI COLLEGATI</th></tr>
		<tr>
			<td colspan="4">
			<%
			sql = " SELECT *," & _
			      " ("& SQL_If(conn, "pa_promo_id = "& cIntero(request("ID")), "1", "NULL") &") AS rel" & _
				  " FROM gtb_articoli a" & _
				  " LEFT JOIN grel_promo_articoli r ON a.art_id = r.pa_art_id" & _
				  " WHERE art_tipologia_id = 1239 AND art_disabilitato = 0" & _
				  " ORDER BY art_nome_it"
			CALL Write_Relations_Checker(conn, rsc, sql, 2, "art_id", "art_nome_it", "rel", "art")
			%>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
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
rs.close
conn.close
set rs = nothing
set rsc = nothing
set conn = nothing
%>