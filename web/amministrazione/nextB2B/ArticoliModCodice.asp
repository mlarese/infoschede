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


sql = " SELECT * FROM gtb_articoli INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " + _
	  " WHERE art_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	if request("tft_art_cod_int")="" then
		Session("ERRORE") = "Codice articolo non valido."
	end if
	
	if request("tft_art_cod_int") = rs("art_cod_int") then
		Session("ERRORE") = "Codice non modificato."
	end if
	
	if cIntero(rs("art_external_id"))>0 then
		sql = "SELECT COUNT(*) FROM gItb_articoli WHERE Iart_x_cod_int LIKE'" + request("tft_art_cod_int") + "' AND iart_id<>" & rs("art_external_id")
		if cIntero(GetValueList(conn, rsp, sql))>0 then
			Session("ERRORE") = "Codice articolo già presente nella struttura di import."
		end if
	end if
	
	sql = "SELECT COUNT(*) FROM gtb_articoli WHERE art_cod_int LIKE'" + ParseSQL(request("tft_art_cod_int"), adChar) + "' AND art_id<>" & cIntero(request("ID"))
	response.write sql
	if cIntero(GetValueList(conn, rsp, sql))>0 then
		Session("ERRORE") = "Il nuovo codice è già stato assegnato ad un articolo."
	end if
	
	if Session("ERRORE") = "" then
		'applica modifiche via query: attenzione sostituisce anche eventuale collegamento con articolo di import collegato.
		'(Giacomo - 06/09/2013 - Ho modificato le query di update in modo che filtrino per id nel caso di due articoli diverso con lo stesso codice)
		
		'sql = "DECLARE @oldCodice nvarchar(30)" + vbCrlf + _
		'	   "SET @oldCodice = '" & ParseSql(rs("art_cod_int"), adChar) & "'" & vbCrlf & _
		sql = "DECLARE @newCodice nvarchar(30)" & vbCrlf & _  
			  "SET @newCodice = '" & ParseSql(request("tft_art_cod_int"), adChar) & "'" & vbCrlf & _
			  "UPDATE gtb_articoli SET art_cod_int = @newCodice WHERE art_id = " & rs("art_id") & vbCrLf
			  '"update gtb_articoli set art_cod_int = @newCodice where art_cod_int LIKE @oldCodice" + vbCrLf
		if cIntero(rs("art_external_id"))>0 then
			  sql = sql + "UPDATE grel_art_valori SET rel_cod_int = @newCodice WHERE rel_art_id = " & rs("art_id") & vbCrlf & _
			  			  "UPDATE gItb_articoli SET Iart_x_cod_int = @newCodice WHERE Iart_id = " & rs("art_external_id")
			  'sql = sql + "update grel_art_valori set rel_cod_int = @newCodice where rel_cod_int LIKE @oldCodice" + vbCrlf + _
			  '			  "update gItb_articoli set Iart_x_cod_int = @newCodice where Iart_x_cod_int LIKE @oldCodice"
		end if
		
		conn.begintrans
		
		CALL WriteLogAdmin(conn,"gtb_articoli", rs("art_id"), "CambioCodice_Articolo", "Cambio codice articolo da '" + rs("art_cod_int") + "' A '" + request("tft_art_cod_int") + "'")
		rs.close
		
		CALL conn.execute(sql)
		conn.committrans
		
		%>
		<script language="JavaScript" type="text/javascript">
			opener.location.reload(true);
			window.close();
		</script>
		<% response.end
	end if
end if
%>

<%'--------------------------------------------------------
sezione_testata = "modifica giacenza dell'articolo"%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>
				Modifica codice dell'articolo
			</caption>
			<tr><th colspan="7">DATI ARTICOLO</th></tr>
			<% CALL ArticoloScheda (conn, rs, rsp) %>
			<tr><th colspan="7">CODICE ARTICOLO</th></tr>
			<tr>
				<td class="label" colspan="3">attuale:</td>
				<td class="content" colspan="4"><%= rs("art_cod_int") %></td>
			</tr>
			<tr>
				<td class="label" colspan="3">nuovo:</td>
				<td class="content" colspan="4">
					<input type="text" class="text" name="tft_art_cod_int" value="<%= rs("art_cod_int") %>" maxlength="50" size="15">
				</td>
			</tr>
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
			