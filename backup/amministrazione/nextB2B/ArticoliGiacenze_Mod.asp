<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<% 
dim conn, rs, rsp, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsp = Server.CreateObject("ADODB.Recordset")

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	if not isNumeric(request("gia_qta")) then
		Session("ERRORE") = "La giacenza immessa non &egrave; valida."
	elseif not isNumeric(request("gia_impegnato")) then
		Session("ERRORE") = "La quantit&agrave; impegnata immessa non &egrave; valida."
	elseif not isNumeric(request("gia_ordinato")) then
		Session("ERRORE") = "La quantit&agrave; ordinata immessa non &egrave; valida."
	end if
	if cInteger(request("old_gia_qta")) <> cInteger(request("gia_qta")) OR _
	   cInteger(request("old_gia_impegnato")) <> cInteger(request("gia_impegnato")) OR _
	   cInteger(request("old_gia_ordinato")) <> cInteger(request("gia_ordinato")) then
	   	'&egrave; variata almeno una quantita'
		conn.begintrans
		dim old_qta, new_qta
		
		'verifica variazione giacenza
		old_qta = cInteger(request("old_gia_qta"))
		new_qta = cInteger(request("gia_qta"))
		if old_qta <> new_qta then
			'esegue variazione giacenza
			CALL SetGiacenza(conn, request("gia_art_var_id"), IIF(new_qta < old_qta, "-", "+"), _
							 QTA_GIACENZA, request("gia_magazzino_id"), Abs(old_qta - new_qta))
		end if
		
		'verifica variazione quantita' impegnata.
		old_qta = cInteger(request("old_gia_impegnato"))
		new_qta = cInteger(request("gia_impegnato"))
		if old_qta <> new_qta then
			'esegue variazione quantita' impegnata
			CALL SetGiacenza(conn, request("gia_art_var_id"), IIF(new_qta < old_qta, "-", "+"), _
							 QTA_IMPEGNATA, request("gia_magazzino_id"), Abs(old_qta - new_qta))
		end if
		
		'verifica variazione quantita' ordinata a fornitore
		old_qta = cInteger(request("old_gia_ordinato"))
		new_qta = cInteger(request("gia_ordinato"))
		if old_qta <> new_qta then
			'esegue quantita' ordinata a fornitore
			CALL SetGiacenza(conn, request("gia_art_var_id"), IIF(new_qta < old_qta, "-", "+"), _
							 QTA_ORDINATA, request("gia_magazzino_id"), Abs(old_qta - new_qta))
		end if
		
		conn.committrans
	end if
	
	'salvo la data di arrivo della merce ordinata a fornitore
	if IsDate(request("gia_ordinato_data_arrivo")) OR cString(request("gia_ordinato_data_arrivo"))="" then
		conn.begintrans
		sql = " UPDATE grel_giacenze SET gia_ordinato_data_arrivo = "
		if cString(request("gia_ordinato_data_arrivo"))="" then
			sql = sql & "NULL"
		else
			sql = sql & SQL_DateTime(conn, ConvertForSave_Date(DateIta(request("gia_ordinato_data_arrivo"))))
		end if
		sql = sql & " WHERE gia_id=" & cIntero(request("ID"))

		conn.execute(sql)
		conn.committrans
	end if
	
	%>
	<script language="JavaScript" type="text/javascript">
		opener.location.reload(true);
		window.close();
	</script>
<%end if

sql = " SELECT * FROM grel_giacenze INNER JOIN gtb_magazzini ON grel_giacenze.gia_magazzino_id = gtb_magazzini.mag_id " + _
	  " INNER JOIN grel_art_valori ON grel_giacenze.gia_art_var_id = grel_art_valori.rel_id " + _
	  " INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + _
	  " INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " + _
	  " WHERE gia_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
%>

<%'--------------------------------------------------------
sezione_testata = "modifica giacenza dell'articolo"%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="gia_magazzino_id" value="<%= rs("gia_magazzino_id") %>">
		<input type="hidden" name="gia_art_var_id" value="<%= rs("gia_art_var_id") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>
				Modifica giacenza dell'articolo a magazzino
			</caption>
			<tr><th colspan="7">DATI ARTICOLO</th></tr>
			<% CALL ArticoloScheda (conn, rs, rsp) %>
			<tr>
				<td class="label" colspan="3">giacenza minima per ogni magazzino</td>
				<td class="content" colspan="4"><%= rs("art_giacenza_min") %></td>
			</tr>
			<tr>
				<td class="label" colspan="3">quantit&agrave; minima ordinabile</td>
				<td class="content" colspan="4"><%= rs("art_qta_min_ord") %></td>
			</tr>
			<tr>
				<td class="label" colspan="3">loto di riordino</td>
				<td class="content" colspan="4"><%= rs("art_lotto_riordino") %></td>
			</tr>
			<tr><th colspan="7">DATI MAGAZZINO E GIACENZE</th></tr>
			<tr>
				<td class="label" colspan="3">magazzino:</td>
				<td class="content" colspan="4"><%= rs("mag_nome") %></td>
			</tr>
			<% if rs("art_se_bundle") then %>
				<tr>
					<td class="label" colspan="3">giacenza a magazzino</td>
					<%if rs("gia_qta")<1 then	'esaurita
					%>
						<td class="content alert" title="esaurito" colspan="4"><%= rs("gia_qta") %></td>
					<%elseif rs("gia_qta") <= cInteger(rs("rel_giacenza_min")) then		'in esaurimento
					%>
						<td class="content warning" title="in esaurimento" colspan="4"><%= rs("gia_qta") %></td>
					<% else %>
						<td class="content ok" colspan="4"><%= rs("gia_qta") %></td>
					<% end if %>
				</tr>
				<tr>
					<td class="label" colspan="3">merce impegnata da ordini</td>
					<td class="content" colspan="4"><%= rs("gia_impegnato") %></td>
				</tr>
			<% else %>
				<tr>
					<td class="label" colspan="3">giacenza a magazzino</td>
					<%if rs("gia_qta")<1 then	'esaurita
					%>
						<td class="content alert" title="esaurito" colspan="4">
					<%elseif rs("gia_qta") <= cInteger(rs("rel_giacenza_min")) then		'in esaurimento
					%>
						<td class="content warning" title="in esaurimento" colspan="4">
					<% else %>
						<td class="content ok" colspan="4">
					<% end if %>
						<input type="hidden" name="old_gia_qta" value="<%= rs("gia_qta")%>">
						<input type="text" class="number" name="gia_qta" value="<%= rs("gia_qta")%>" size="7">
						(*)
					</td>
				</tr>
				<tr>
					<td class="label" colspan="3">merce impegnata da ordini</td>
					<td class="content" colspan="4">
						<input type="hidden" name="old_gia_impegnato" value="<%= rs("gia_impegnato")%>">
						<input type="text" class="number" name="gia_impegnato" value="<%= rs("gia_impegnato")%>" size="7">
						(*)
					</td>
				</tr>
				<tr>
					<td class="label" colspan="3">merce ordinata a fornitore</td>
					<td class="content" colspan="4">
						<input type="hidden" name="old_gia_ordinato" value="<%= rs("gia_ordinato")%>">
						<input type="text" class="number" name="gia_ordinato" value="<%= rs("gia_ordinato")%>" size="7">
						(*)
					</td>
				</tr>
				<tr>
					<td class="label" colspan="3">data arrivo merce ordinata a fornitore</td>
					<td class="content" colspan="4">
						<% CALL WriteDataPicker_Input("form1", "gia_ordinato_data_arrivo", DateIta(rs("gia_ordinato_data_arrivo")), "", "/", true, true, LINGUA_ITALIANO) %>
					</td>
				</tr>
			<% end if %>
			<tr>
				<td class="footer" colspan="7">
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
					<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
				</td>
			</tr>
		</table>
	</form>
</div>
</body>
</html>
<% rs.close
conn.close
set rs = nothing
set rsp = nothing
set conn = nothing %>