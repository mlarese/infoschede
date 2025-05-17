<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<% 
dim conn, rs, sql, tipo
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	
	conn.beginTrans
	sql = "SELECT * FROM gtb_articoli WHERE art_id=" & cIntero(request("ID"))
	rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
	'aggiorna prezzo base dell'articolo
	rs("art_iva_id") = request("tfn_art_iva_id")
	rs.update
	rs.close
	
	'aggiorna categoria iva nei listini
	select case cInteger(request("aggiornamento_listini"))
		case 0
			sql = "UPDATE gtb_prezzi SET prz_iva_id=" & cIntero(request("tfn_art_iva_id")) & " WHERE prz_iva_id=" & cIntero(request("iva_old")) & _
				  " AND prz_variante_id IN (SELECT rel_id FROM grel_Art_valori WHERE rel_Art_id=" & cIntero(request("ID")) & ")"
			CALL conn.execute(sql, , adExecuteNoRecords)
		case 1
			sql = "UPDATE gtb_prezzi SET prz_iva_id=" & cIntero(request("tfn_art_iva_id")) & " WHERE " & _
				  " prz_variante_id IN (SELECT rel_id FROM grel_Art_valori WHERE rel_Art_id=" & cIntero(request("ID")) & ")"
			CALL conn.execute(sql, , adExecuteNoRecords)
	end select
	
	conn.commitTrans%>
	<script language="JavaScript" type="text/javascript">
		opener.location.reload(true);
		window.close();
	</script>
<%end if


sql = "SELECT * FROM gtb_articoli INNER JOIN gtb_iva ON gtb_articoli.art_iva_id = Gtb_iva.iva_id WHERE art_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

if rs("art_se_bundle") then
	tipo = "del bundle"
elseif rs("art_se_confezione") then
	tipo = "della confezione"
elseif rs("art_varianti") then
	tipo ="dell'articolo con varianti"
else
	tipo ="dell'articolo singolo"
end if
%>

<%'--------------------------------------------------------
sezione_testata = "modifica categoria i.v.a. " & tipo %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="iva_old" value="<%= rs("art_iva_id") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>
				Modifica della categoria I.V.A. <%= tipo %>
			</caption>
			<tr><th colspan="3">DATI BASE ARTICOLO</th></tr>
			<tr>
				<td class="label" rowspan="2">articolo</td>
				<td class="content" colspan="2"><%= rs("art_nome_it") %></td>
			</tr>
			<tr>
				<% if rs("art_se_bundle") then %>
					<td colspan="2" class="content bundle">bundle</td>
				<% elseif rs("art_se_confezione") then %>
					<td colspan="2" class="content confezione">confezione</td>
				<% elseif rs("art_varianti") then %>
					<td colspan="2" class="content varianti">articolo con varianti</td>
				<% else %>
					<td colspan="2" class="content">articolo singolo</td>
				<% end if %>
			</tr>
			<tr>
				<td class="label">categoria i.v.a.:</td>
				<td class="content" colspan="2">
					<% sql = "SELECT * FROM gtb_iva ORDER BY iva_ordine"
					CALL dropDown(conn, sql, "iva_id", "iva_nome", "tfn_art_iva_id", rs("art_iva_id"), true, "", LINGUA_ITALIANO)%>
				</td>
			</tr>
			<tr><th colspan="3">AGGIORNAMENTO LISTINI</th></tr>
			<tr>
				<td class="label" rowspan="3">tipo:</td>
				<td class="content_center">
					<input type="radio" class="checkbox" name="aggiornamento_listini" value="0" <%= chk(cInteger(request("aggiornamento_listini"))=0) %>>
				</td>
				<td class="content">
					sostituisci solo <%= rs("iva_nome") %><br>
					<span class="note">Sostituisce solo la categoria <%= rs("iva_nome") %> nei listini.</span>
				</td>
			</tr>
			<tr>
				<td class="content_center">
					<input type="radio" class="checkbox" name="aggiornamento_listini" value="1" <%= chk(cInteger(request("aggiornamento_listini"))=1) %>>
				</td>
				<td class="content">
					sostituisci tutte<br>
					<span class="note">Imposta la categoria in tutti i listini.</span>
				</td>
			</tr>
			<tr>
				<td class="content_center">
					<input type="radio" class="checkbox" name="aggiornamento_listini" value="2" <%= chk(cInteger(request("aggiornamento_listini"))=2) %>>
				</td>
				<td class="content">
					lascia categorie i.v.a. invariate in tutti i listini.
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="3">
					(*) Campi obbligatori.
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
					<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
				</td>
			</tr>
	</form>
		</table>
</div>
</body>
</html>
<% rs.close
conn.close
set rs = nothing
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>