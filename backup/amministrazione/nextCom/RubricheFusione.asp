<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<% dim conn, sql, rsr, rs

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")
%>

<%'--------------------------------------------------------
sezione_testata = "Rubriche - modifica - fusione rubriche" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'-----------------------------------------------------

sql = "SELECT * FROM tb_rubriche WHERE id_rubrica=" & cIntero(request("rubrica_sorgente"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption class="border">
			Fusione rubriche
		</caption>
		<tr><th colspan="3">RUBRICA SORGENTE</th></tr>
		<tr>
			<td class="label_no_width" style="width:25%;">Nome</td>
			<td class="content">
				<%= rs("nome_rubrica") %>
			</td>
			<td class="content notes" style="width:30%;" rowspan="2">
				Rubrica i cui contatti vengono associati alla rubrica di destinazione.
			</td>
		</tr>
		<tr>
			<td class="label_no_width">N&ordm; contatti associati</td>
			<% sql = "SELECT COUNT(*) FROM rel_rub_ind WHERE id_rubrica=" & rs("id_rubrica") %>
			<td class="content"><%= cInteger(GetValueList(conn, rsr, sql)) %></td>
		</tr>
		<tr><th colspan="3">RUBRICA DESTINAZIONE</th></tr>
		<% if request("fondi")="" OR cIntero(request("rubrica_destinazione"))=0 then %>
			<tr>
				<td class="label">Scegli la rubrica</td>
				<td class="content" colspan="2">
					<% sql = "SELECT * FROM tb_rubriche WHERE id_rubrica <> " & rs("id_rubrica") & " AND " & SQL_IfIsNull(conn, "Syncrotable", "''") & "='' ORDER BY nome_rubrica"
					CALL dropDown(conn, sql, "id_rubrica", "nome_rubrica", "rubrica_destinazione", request("rubrica_destinazione"), true, _
														  "", LINGUA_ITALIANO)%>
				</td>
			</tr>
		<% else 
			sql = "SELECT * FROM tb_rubriche WHERE id_rubrica=" & cIntero(request("rubrica_destinazione"))
			rsr.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText %>
			<tr>
				<td class="label">Nome</td>
				<td class="content" colspan="2">
					<%= rsr("nome_rubrica") %>
				</td>
			</tr>
			<tr>
				<td class="label_no_width" rowspan="2">N&ordm; contatti associati</td>
				<td class="label_no_width">prima della fusione</td>
				<% sql = "SELECT COUNT(*) FROM rel_rub_ind WHERE id_rubrica=" & rsr("id_rubrica") %>
				<td class="content"><%= cInteger(GetValueList(conn, NULL, sql)) %></td>
			</tr>
			<% 'esegue fusione
			sql = "INSERT INTO rel_rub_ind (id_rubrica, id_indirizzo) " & _
				  "	SELECT " & ParseSQL(request("rubrica_destinazione"), adChar) & ", id_indirizzo " & _
				  "	FROM rel_rub_ind WHERE id_rubrica=" & cIntero(request("rubrica_sorgente")) & " AND " & _
				  						 " id_indirizzo NOT IN (SELECT id_indirizzo FROM rel_rub_ind WHERE id_rubrica=" & cIntero(request("rubrica_destinazione")) & ") "
			call conn.execute(sql)
			%>
			<tr>
				<td class="label_no_width">dopo la fusione</td>
				<% sql = "SELECT COUNT(*) FROM rel_rub_ind WHERE id_rubrica=" & rsr("id_rubrica") %>
				<td class="content"><%= cInteger(GetValueList(conn, NULL, sql)) %></td>
			</tr>
			<% rsr.close
		end if %>
		<tr>
			<td class="footer" colspan="3">
				<input type="submit" class="button" name="fondi" value="FONDI RUBRICHE">
				<input type="button" class="button" name="chiudi" value="ANNULLA" onclick="window.close();">
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
set rsr = nothing
set conn = nothing %>
<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>
