<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<% 
dim conn, rs, sql, field

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = Session(request.querystring("QRY"))

sql = "SELECT " + _
	  " ( COUNT(ord_id) ) AS [Numero ordini]," + _
	  "	( CONVERT(money, SUM(ord_totale)) ) AS [Totale ordinato], " + _
	  " ( CONVERT(money, SUM(ord_totale_spese)) ) AS [Totale addebiti spese], " + _
	  " ( SUM(ord_colli) ) AS [Totale colli], " + _
	  " ( SUM(ord_totale_volume) ) AS [Volume totale], " + _
	  " ( SUM(ord_totale_peso_netto) ) AS [Peso netto], " + _
	  " ( SUM(ord_totale_peso_lordo) ) AS [Peso lordo] " + _
	  " FROM gtb_ordini" & _
	  " WHERE ord_id IN (" & sql &") "
	  
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>

<html>
	<head>
		<title>Export totali ordine</title>
		<link rel="stylesheet" type="text/css" href="../library/stili.css">
		<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
		<meta name="robots" content="noindex,nofollow" />
		<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
		<SCRIPT LANGUAGE="javascript" src="../library/utils.js" type="text/javascript"></SCRIPT>
	</head>
<body topmargin="9" onload="window.focus()">
<!-- <%=sql%> -->
<form action="" method="post" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0">
				<tr>
					<td class="caption">PRENOTAZIONI</td>
					<td align="right" style="padding-right:5px;">
						<input type="button" onclick="window.close();" class="button" name="chiudi" value="CHIUDI">
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="2">Totali</th></tr>
		<% if not rs.eof then 
			for each field in rs.fields %>
				<%= field.type %>
				<tr>
					<td class="content"><%=field.name%></td>
					<td class="content_right" style="padding-right:20px;">	
						<% if field.type = adCurrency then %>
							<%=FormatPrice(field.value, 2, true)%>&euro;
						<% else %>
							<%=field.value%>
						<% end if %>
					</td>
				</tr>
			<% next
		end if %>
		<tr>
			<td colspan="2" class="footer">
				<input type="button" onclick="window.close();" class="button" name="chiudi" value="CHIUDI">
			</td>
		</tr>
	</table>
</form>
</body>
</html>

<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>

<% 
rs.close
conn.close 
set rs = nothing
set conn = nothing
%>