<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = true %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/ClassIndex.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/ClassContent.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->
<%
'controllo permessi
CALL CheckAutentication(Session("PASS_ADMIN") <> "" OR Session("PASS_AMMINISTRATORI") <> "")

dim conn, rs, sql,ID
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
ID = CIntero(request("ID"))

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("SQL_APPLICAZIONI"), "id_sito", "ApplicazioniParamsModifica.asp")
end if

'salva
if request.form("salva") <> "" then
	conn.BeginTrans
	dim lingua
	'svuoto i valori dei descrittori al posto di cancellarli (i descrittori booleani non vengono modificati se false)
	sql = " UPDATE rel_siti_descrittori SET "& _
		  SQL_MultiLanguage("rsd_valore_<LINGUA> = ''", ",") &", "& _
		  SQL_MultiLanguage("rsd_memo_<LINGUA> = ''", ",") & _
		  " WHERE rsd_sito_id = "& ID
	conn.Execute(sql)
	
	'non cancello i descrittori vuoti (perchè la relazione valore e' usata anche come associazione descrittore sito)
	CALL DesSave(conn, ID, "rel_siti_descrittori", "rsd_valore_", "rsd_memo_", "rsd_sito_id", "rsd_descrittore_id", " AND 1=0")
	
	'cancello solo i doppioni
	sql = " DELETE FROM rel_siti_descrittori WHERE rsd_sito_id = "& ID & _
	 	  " AND rsd_id = (SELECT MIN(rsd_id) FROM rel_siti_descrittori r"& _
		  " 			   WHERE rsd_sito_id = "& ID & _
		  "				   AND r.rsd_descrittore_id = rel_siti_descrittori.rsd_descrittore_id"& _
		  " 			   GROUP BY rsd_sito_id HAVING COUNT(*) > 1)"
	conn.Execute(sql)
	
	'aggiorno la data di modifica dei parametri
	sql = "UPDATE tb_webs SET webs_modData_parametri = " + SQL_Now(conn)
	conn.Execute(sql)
	
	conn.CommitTrans
	response.redirect "Applicazioni.asp"
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione applicazioni - modifica"
if session("PASS_ADMIN") = "" then
	dicitura.puls_new = "INDIETRO;ACCESSI"
	dicitura.link_new = "Applicazioni.asp;ApplicazioniAccessi.asp?ID=" & ID
else
	dicitura.puls_new = "INDIETRO;DATI APPLICAZIONE;ACCESSI;TABELLE DATI"
	dicitura.link_new = "Applicazioni.asp;ApplicazioniMod.asp?ID=" & request("ID") & ";ApplicazioniAccessi.asp?ID=" & ID & ";ApplicazioniTabelle.asp?ID=" & ID
end if
dicitura.scrivi_con_sottosez()
%>
<style type="text/css">
	td.nomedescrittore{
		width:25% !important;
	}
</style>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica parametri dell'applicazione "<%= GetValueList(conn, rs, "SELECT sito_nome FROM tb_siti WHERE id_sito = "& ID) %>"</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= ID %>&goto=PREVIOUS" title="applicazione precedente">
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= ID %>&goto=NEXT" title="applicazione successiva">
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">PARAMETRI</th></tr>
		<%  sql = " SELECT * FROM (tb_siti_descrittori_raggruppamenti g"& _
				  " RIGHT JOIN tb_siti_descrittori d ON g.sdr_id = d.sid_raggruppamento_id)"& _
				  " INNER JOIN rel_siti_descrittori r ON d.sid_id = r.rsd_descrittore_id"& _
				  " WHERE rsd_sito_id = "& ID & _
				  IIF(session("PASS_ADMIN") = "", " AND NOT "& SQL_IsTrue(conn, "sid_admin"), "")
			CALL DesFullFormConn(conn, conn, sql & " ORDER BY sdr_ordine, sid_codice", _
						 "tb_siti_descrittori", "sid_id", "sid_codice", "sid_tipo", "sid_unita_it", "sid_nome_it", "", "", "rsd_valore_", "rsd_memo_", _
						 "sdr_titolo_it", false, 4) %>
		<tr>
			<td class="footer" colspan="4">
				<% 	if CIntero(GetValueList(conn, rs, Replace(sql, "*", "COUNT(*)"))) = 0 then %>
				<a href="Applicazioni.asp" class="simul_puls_1">INDIETRO</a>
				<% 	else %>
				<input type="submit" class="button" name="salva" value="SALVA">
				<% 	end if %>
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<%
set rs = nothing
conn.close
set conn = nothing
%>