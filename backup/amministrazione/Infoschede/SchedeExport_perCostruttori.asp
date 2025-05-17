<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = true %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/TOOLS4ADMIN.ASP" -->
<!--#INCLUDE FILE="Tools_Infoschede_Const.asp" -->

<%
dim conn, rs, rss, rsd, sql, sessionSQL, modello, tot_ricambi
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rss = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.RecordSet")

	
if cString(Session("INFOSCHEDE_SCHEDE_SQL")) = "" then
	%>
	<script language="JavaScript" type="text/javascript">
		window.close();
	</script>
	<%
	response.end
end if
sessionSQL = Session("INFOSCHEDE_SCHEDE_SQL")
sessionSQL = Right(sessionSQL, Len(sessionSQL) - inStr(sessionSQL, " FROM "))
sessionSQL = "SELECT sc_id " & sessionSQL
sessionSQL = Left(sessionSQL, inStrRev(sessionSQL, " ORDER BY "))


' raggruppato per costruttore
sql = Replace(sessionSQL, "SELECT sc_id FROM", "SELECT sc_modello_id FROM")
sql = " SELECT NomeOrganizzazioneElencoIndirizzi, CognomeElencoIndirizzi, NomeElencoIndirizzi, isSocieta, " & _
	  " riv_id " & _
	  " FROM gv_rivenditori WHERE riv_profilo_id = "&COSTRUTTORI&" AND " & _
	  " riv_id IN (SELECT mar_anagrafica_id FROM gv_articoli WHERE rel_id IN ("&sql&"))" & _
	  " ORDER BY ModoRegistra"

	
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
%>

<html>
	<head>
		<title>Export schede ricercate</title>
		<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
		<META NAME="copyright" CONTENT="Copyright &copy;2003 - next-aim.com">
		<meta name="robots" content="noindex,nofollow" />
		<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
		<link rel="stylesheet" type="text/css" href="../library/stili.css">
		<SCRIPT LANGUAGE="javascript"  src="../library/utils.js" type="text/javascript"></SCRIPT>
	</head>
	
	<body onload="window.focus();" leftmargin="4" topmargin="3">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption class="border">
				<table width="100%" border="0" cellspacing="0">
					<tr>
						<td class="caption">Export schede, per costruttore</td>
						<td align="right" style="padding-right:5px;"><a class="button" href="javascript:window.close();">CHIUDI</a></td>
					</tr>
				</table>
			</caption>
			<% while not rs.eof %>
				<tr>
					<td>
						<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin:8px;">
							<caption class="border"><%=ContactFullName(rs)%></caption>
							<tr>
								<!-- sezione ricambi -->
								<td style="width:70%; vertical-align:top;">
									<% 
									sql = " SELECT dts_ricambio_codice, dts_ricambio_nome, dts_ricambio_prezzo, SUM(dts_ricambio_qta) AS dts_ricambio_qta, " & _
										  " dts_ricambio_sconto, SUM(dts_prezzo_totale) AS dts_prezzo_totale " & _
										  " FROM grel_art_valori RIGHT JOIN sgtb_dettagli_schede " & _
										  " ON grel_art_valori.rel_id = sgtb_dettagli_schede.dts_ricambio_id " & _
										  " WHERE dts_scheda_id IN (" & sessionSQL & ") " & _
										  " AND dts_scheda_id IN (SELECT sc_id FROM sgtb_schede WHERE sc_modello_id IN (SELECT rel_id FROM gv_articoli WHERE mar_anagrafica_id = "&rs("riv_id")&")) " & _
										  " GROUP BY dts_ricambio_codice, dts_ricambio_nome, dts_ricambio_prezzo, dts_ricambio_qta, " & _
										  " dts_ricambio_sconto, dts_prezzo_totale " & _
										  " ORDER BY dts_ricambio_codice "
										  
										  '" AND grel_art_valori.rel_art_id IN (SELECT art_id FROM gv_articoli WHERE mar_anagrafica_id = "&rs("riv_id")&") " & _
									
									rsd.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
									tot_ricambi = 0
									%>
									<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-left:0px; border-bottom:0px;">
										<tr><td colspan="6" class="header">Ricambi</td></tr>
										<% if rsd.eof then %>
											<tr><td colspan="6" class="content">Nessun ricambio utilizzato.</td></tr>
										<% else %>
											<tr>
												<th class="L2" width="14%">codice</th>
												<th class="L2">ricambio</th>
												<th class="l2_center" width="11%">prezzo</th>
												<th class="l2_center" width="8%">quantit&agrave;</th>
												<th class="l2_center" width="10%">sconto</th>
												<th class="l2_center" width="13%">totale</th>
											</tr>
										
											<% while not rsd.eof %>
												<tr>
													<td class="content"><%= rsd("dts_ricambio_codice")%></td>
													<td class="content"><%= rsd("dts_ricambio_nome")%></td>
													<td class="content_center"><%= FormatPrice(cReal(rsd("dts_ricambio_prezzo")), 2, false)%> &euro;</td>
													<td class="content_center"><%= rsd("dts_ricambio_qta")%></td>
													<td class="content_center"><%= rsd("dts_ricambio_sconto")%> %</td>
													<td class="content_center" nowrap><%= FormatPrice(cReal(rsd("dts_prezzo_totale")), 2, false)%> &euro;</td>
													<% tot_ricambi = tot_ricambi + cReal(rsd("dts_prezzo_totale"))
													%>
												</tr>
												<% rsd.moveNext %>
											<% wend %>
											<tr>
												<td colspan="5" class="header" >&nbsp;</td>
												<td class="header" align="center"><%= FormatPrice(cReal(tot_ricambi), 2, false)%> &euro;</td>
											</tr>
										<% end if %>
										<% rsd.close %>
									</table>
								</td>
								
								<!-- sezione dettagli schede -->
								<%
								sql = " SELECT SUM(sc_ora_manodopera_intervento) AS ore_manodopera, SUM(sc_prezzo_manodopera*sc_ora_manodopera_intervento) AS sc_prezzo_manodopera, " & _
									  " SUM(sc_costo_presa) AS sc_costo_presa, SUM(sc_costo_riconsegna) AS sc_costo_riconsegna " & _
									  " FROM sgtb_schede WHERE sc_id IN ("&sessionSQL&") " & _
									  " AND sc_modello_id IN (SELECT rel_id FROM gv_articoli WHERE mar_anagrafica_id = "&rs("riv_id")&") "
								
								rss.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
								%>
								<td style="vertical-align:top;">
									<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
										<tr><td colspan="2" class="header">Riepilogo</td></tr>
										<tr>
											<td class="content" width="60%">ore manodopera</td>
											<td class="content"><%= cReal(rss("ore_manodopera"))%></td>
										</tr>
										<tr>
											<td class="content">tot. manodopera</td>
											<td class="content"><%= FormatPrice(cReal(rss("sc_prezzo_manodopera")),2,false) %> &euro;</td>
										</tr>
										<tr>
											<td class="content">costo presa</td>
											<td class="content"><%= FormatPrice(cReal(rss("sc_costo_presa")),2,false) %> &euro;</td>
										</tr>
										<tr>
											<td class="content">costo riconsegna</td>
											<td class="content"><%= FormatPrice(cReal(rss("sc_costo_riconsegna")),2,false) %> &euro;</td>
										</tr>
									</table>
								</td>
								<% rss.close %>
							</tr>
							<%
							sql = " SELECT sc_numero, sc_modello_id, sc_data_ricevimento, sc_modello_altro " & _
								  " FROM sgtb_schede WHERE sc_id IN ("&sessionSQL&") " & _
								  " AND sc_modello_id IN (SELECT rel_id FROM gv_articoli WHERE mar_anagrafica_id = "&rs("riv_id")&") " & _
								  " ORDER BY sc_numero "
							rss.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
							%>
							<tr>
								<td colspan="2">
									<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-left:0px; border-bottom:0px; border-top:1px solid #999999;">
										<tr><td colspan="6" class="header">Schede</td></tr>
										<tr>
											<th class="L2" width="9%">numero</th>
											<th class="L2" width="15%">del</th>
											<th class="L2">modello</th>
										</tr>
										<% while not rss.eof %>
											<tr>
												<td class="content"><%= rss("sc_numero")%></td>
												<td class="content"><%= rss("sc_data_ricevimento")%></td>
												<%
												if rss("sc_modello_altro")<>"" then
													modello = rss("sc_modello_altro")
												else
													sql = "SELECT art_nome_it FROM gv_articoli WHERE rel_id = " & rss("sc_modello_id")
													modello = GetValueList(conn, NULL, sql)
												end if
												%>
												<td class="content"><%= modello%></td>
											</tr>
											<% rss.moveNext %>
										<% wend %>
									</table>
								</td>
							</tr>
							<% rss.close %>
							
						</table>
					</td>
				</tr>
				<% rs.moveNext %>
			<% wend %>
			</tr>
			<tr>
				<td class="footer" align="right" style="padding-right:5px;" colspan="2">
					<a class="button" href="javascript:window.close();">CHIUDI</a>
				</td>
			</tr>
		</table>
	</body>
</html>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>

<%
rs.close
conn.close 
set rs = nothing
set rss = nothing
set rsd = nothing
set conn = nothing
%>

