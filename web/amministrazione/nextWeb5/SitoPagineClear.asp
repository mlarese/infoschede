<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<% 
dim conn, rs, sql
set conn = server.createObject("ADODB.Connection")
set rs = server.createObject("ADODB.RecordSet")
conn.Open Application("DATA_ConnectionString"),"",""

conn.beginTrans

sql = " SELECT * FROM tb_layers " + _
 	  " WHERE ( id_tipo=1 OR id_tipo=5 ) " + _
	  		" AND id_pag=" & cIntero(request("PAGINA"))
rs.open sql, conn, adOpenStatic, adLockOptimistic
while not rs.eof
	if cString(rs("HTML"))<>"" then
		rs("HTML") =  ClearString(rs("HTML"), true)
		rs("TESTO") =  ClearString(rs("TESTO"), true)
		rs.update
	end if
	rs.moveNext
wend
rs.close

conn.CommitTrans

conn.close
set rs = nothing
set conn = nothing
%>

<%'--------------------------------------------------------
sezione_testata = "Gestione siti - indice delle pagine - ripulisci la pagina" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption class="border">
				Pulizia pagina da caratteri spuri
			</caption>
			<tr>
				<td class="content_b">
					Pulizia eseguita correttamente.
				</td>
			</tr>
			<tr>
				<td class="note">
					Questa finestra si chiuder&agrave; automaticamente tra 5 secondi.
				</td>
			</tr>
			<tr>
				<td class="footer">
					<input type="button" class="button" name="chiudi" value="CHIUDI" onclick="window.close();">
				</td>
			</tr>
		</form>
	</table>
</div>
</body>
</html>
<script language="JavaScript">
	window.setTimeout("close();", 5000);
</script>
<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>
