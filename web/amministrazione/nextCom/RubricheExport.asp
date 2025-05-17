<!--#INCLUDE FILE="../library/Tools.ASP" -->
<%
dim conn, rs, fine, sql
set conn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.recordset")
conn.open Application("DATA_ConnectionString")

conn.BeginTrans

fine = false
'tolgo SELECT *
sql = LTrim(session(request.querystring("sql")))
sql = " "& right(sql, len(sql) - InStr(1, sql, "FROM", vbTextCompare) + 1)
'tolgo ORDER BY
if InStr(1, sql, "ORDER BY", vbTextCompare) > 0 then
	sql = left(sql, InStr(1, sql, "ORDER BY", vbTextCompare) - 1)
end if

if request.form("nuovo") <> "" then							'nuova rubrica
	rs.open "SELECT * FROM tb_rubriche", conn, adOpenKeySet, adLockOptimistic
	rs.addNew
		rs("nome_rubrica") = request.form("nome")
		rs("locked_rubrica") = true
		rs("rubrica_esterna") = true
		rs("syncroTable") = request.querystring("tabella")
	rs.update
	
	conn.execute(" INSERT INTO rel_rub_ind(id_indirizzo, id_rubrica) "& _
				 " SELECT DISTINCT IDelencoIndirizzi, "& rs("id_rubrica") & sql)
				 
	conn.execute(" INSERT INTO tb_rel_gruppiRubriche(id_dellaRubrica, id_gruppo_assegnato) "& _
				 " SELECT "& rs("id_rubrica") &", id_gruppo FROM tb_rel_dipGruppi WHERE id_impiegato="& session("ID_ADMIN"))
	rs.close
	fine = true

elseif request.form("ok") <> "" then						'cancello o salvo
	if request.querystring("CANC") <> "" then				'cancello rubrica
		conn.execute("DELETE FROM tb_rubriche WHERE syncroTable='"& ParseSQL(request.querystring("tabella"), adChar) &"' AND id_rubrica="& cIntero(request.querystring("CANC")))
	elseif request.querystring("SALVA") <> "" then			'salvo rubrica
		if cIntero(request.querystring("SALVA")) > 0 then
			conn.execute("DELETE FROM rel_rub_ind WHERE id_rubrica="& cIntero(request.querystring("SALVA")))
			conn.execute(" INSERT INTO rel_rub_ind(id_indirizzo, id_rubrica) "& _
						 " SELECT DISTINCT IDelencoIndirizzi, "& cIntero(request.querystring("SALVA")) & sql)
		end if
	end if
	fine = true

end if

conn.CommitTrans

if fine then
%>
<SCRIPT LANGUAGE="javascript" type="text/javascript">
	opener.document.location = opener.document.location
	window.close();
</SCRIPT>
<%
end if

set rs = nothing
conn.close
set conn = nothing
%>

<html>
	<head>
		<title>Criteri cliente</title>
		<link rel="stylesheet" type="text/css" href="../library/stili.css">
		<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
		<meta name="robots" content="noindex,nofollow" />
		<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
		<SCRIPT LANGUAGE="javascript" src="../library/utils.js" type="text/javascript"></SCRIPT>
	</head>
<body topmargin="9" onload="window.focus()">
<form action="" method="post" name="form1">
<input type="hidden" name="tipo_cri" value="Clienti">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td>
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
				<% 	if request.querystring("canc") <> "" OR request.querystring("salva") <> "" then %>
					<caption class="border">Risultati salvati come rubriche</caption>
					<tr>
						<td class="content_center" colspan="2">
							<img src="../grafica/alert_anim.gif">
						</td>
					</tr>
					<tr>
						<td>
							<table cellpadding="0" cellspacing="0" width="100%" style="padding-bottom:10px;">
								<tr>
									<td class="content_center" colspan="2">
										Aggiornare il contenuto della rubrica l'elenco dei clienti cercati?
									</td>
								</tr>
								<tr>
									<td class="content_center">
										<input type="submit" name="ok" value="CONFERMA" class="button" style="width:80px;">
									</td>
									<td class="content_center">
										<input type="button" name="annulla" value="ANNULLA" class="button" style="width:80px;" onclick="window.close();">
									</td>
								</tr>
								<tr>
									<td class="note" colspan="2">
										ATTENZIONE: le associazioni dei clienti attualmente presenti verranno cancellate.
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td class="footer">
							<input type="button" onclick="window.close();" class="button" name="chiudi" value="CHIUDI">
						</td>
					</tr>
				<% 	else %>
					<caption>Registraizone risultati in una nuova rubrica</caption>
					<tr><th colspan="2">NOME RUBRICA</th></tr>
					<tr>
						<td class="label">nome:</td>
						<td class="content">
							<input type="text" name="nome" value="" style="width:100%;">
						</td>
					</tr>
					<tr>
						<td colspan="2" class="footer">
							<input type="submit" name="nuovo" value="SALVA" class="button">
							<input type="button" onclick="window.close();" class="button" name="chiudi" value="ANNULLA">
						</td>
					</tr>
				<% 	end if %>
				</table>
			</td>
		</tr>
	</table>
</form>
</body>
</html>