<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
dim nextWeb_Conn, rs, sql, lingua, value
set nextWeb_Conn = Server.CreateObject("ADODB.Connection")
nextWeb_Conn.open Application("l_conn_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")

if request("salva")<>"" then
	sql = "UPDATE tb_pages SET lingua='" & ParseSql(request.form("lingua"), adChar) & "' WHERE id_page=" & cIntero(request("pagina"))
	CALL nextWeb_Conn.execute(sql) %>
	<script language="JavaScript">
		opener.SetPreview( <%= request("pagina") %>);
		window.close()
	</script>
<% end if

'--------------------------------------------------------
sezione_testata = "Gestione lingua di spedizione della next-email" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

sql = "SELECT * FROM tb_pages INNER JOIN tb_webs ON tb_pages.id_webs = tb_webs.id_webs WHERE id_page=" & cIntero(request("pagina"))
rs.open sql, nextWeb_Conn, adOpenStatic, adLockOptimistic
%>

<div id="content_ridotto">
<form action="" method="post" name="form1" id="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<%= sezione_testata %>
		</caption>
		<tr>
			<th colspan="2">SELEZIONA LA LINGUA</th>
		</tr>
		<tr>
			<td class="label" rowspan="2">lingua:</td>
			<td class="content">
				<select name="lingua">
					<% for each lingua in application("LINGUE")
						if lingua = LINGUA_ITALIANO then
							value = true
						elseif rs("lingua_" & lingua) then
							value = true
						else
							value = false
						end if
						if value then %>
							<option value="<%= lingua %>" <%= IIF(rs("lingua") = lingua, "selected", "") %>><%= GetNomeLingua(lingua)%></option>
						<% end if
					next %>
				</select>
			</td>
		</tr>
		<tr>
			<td class="content notes">
				La selezione della lingua permette di generare la pagina da spedire nella lingua desiderata.
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="2">
				<input type="button" class="button" name="chiudi" value="ANNULLA" onclick="window.close();">
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
</form>
</div>
</body>
</html>

<% rs.close
nextWeb_Conn.close
set rs = nothing
set nextWeb_Conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>