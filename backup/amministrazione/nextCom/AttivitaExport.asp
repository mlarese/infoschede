<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = true %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/TOOLS4ADMIN.ASP" -->
<!--#INCLUDE FILE="../library/ExportTools.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<%
	dim conn, rs, rsv, sql, sessionSQL
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open Application("DATA_ConnectionString")
	set rs = Server.CreateObject("ADODB.RecordSet")
	set rsv = Server.CreateObject("ADODB.RecordSet")
	dim rubriche_visibili
	'recupera rubriche visibili all'utente
rubriche_visibili = GetList_Rubriche(conn, rs)
	
if request("esporta")="" then%>
	<html>
		<head>
			<title>Export dati ricercati</title>
			<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
			<META NAME="copyright" CONTENT="Copyright &copy;2003 - next-aim.com">
			<meta name="robots" content="noindex,nofollow" />
			<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
			<link rel="stylesheet" type="text/css" href="../library/stili.css">
			<SCRIPT LANGUAGE="javascript"  src="../library/utils.js" type="text/javascript"></SCRIPT>
		</head>
		
		<body onload="window.focus();" leftmargin="4" topmargin="3">
			<form action="" method="post" target="export" id="export" name="Export">
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						<table width="100%" border="0" cellspacing="0">
							<tr>
								<td class="caption">Export dei dati</td>
								<td align="right" style="padding-right:5px;"><a class="button" href="javascript:window.close();">CHIUDI</a></td>
							</tr>
						</table>
					</caption>
					<tr><th colspan="2">SELEZIONA UNA RUBRICA</td></tr>
					<tr>
						<td class="content" colspan="2" width="100%">
							<script language="JavaScript" type="text/javascript">
												function ShowName(obj){
													var value = obj.options(obj.selectedIndex).text;
													if (value.length>33)
														alert(obj.options(obj.selectedIndex).text);
												}
											</script>
											<% sql = " SELECT " & _ 
													 IIF(DB_Type(conn) = DB_SQL, "(' ' + CAST(id_rubrica AS nvarchar(8)) + ' ') ", "(' ' & id_rubrica & ' ')") & " AS ID, " &_
													 " nome_rubrica FROM tb_rubriche " &_
													 " WHERE id_rubrica IN (" & rubriche_visibili & ") " &_
													 " ORDER BY rubrica_esterna, nome_rubrica"
											CALL dropDown(conn, sql, "ID", "nome_rubrica", "search_rubriche", Session("search_rubriche"), true, _
														  "multiple size=""20"" style=""width:100%;"" onDblClick=""ShowName(this);""", LINGUA_ITALIANO)%>
											<div class="note">
												Ctrl + Click per selezioni multiple.<br>
												Doppio click per visualizzare il nome.
							</div>
						</td>
					</tr>
					<tr><th colspan="2">FORMATO DEI DATI</td></tr>
					<tr>
						<td class="content" width="15%"><INPUT  value="XLS_XP" class="checkbox" type="radio" name="export"></td>
						<td class="content">foglio Excel XP o successivo</td>
					</tr>
					<tr>
						<td class="content"><INPUT  value="XLS_2000" class="checkbox" type="radio" name="export"></td>
						<td class="content">foglio Excel 2000 o precedente</td>
					</tr>
					<tr>
						<td class="content"><INPUT  value="HTML" class="checkbox" type="radio" name="export"></td>
						<td class="content">pagina HTML</td>
					</tr>
					<tr>
						<td class="content"><INPUT  value="TXT" class="checkbox" type="radio" name="export"></td>
						<td class="content">file di testo separato da &quot;;&quot;</td>
					</tr>
					<tr>
						<td class="content"><INPUT  value="XML" class="checkbox" type="radio" name="export"></td>
						<td class="content">file XML</td>
					</tr>
					<tr>
						<td class="footer" colspan="2">
							<input onclick="window.close();" type="button" name="esporta" value="ANNULLA" class="button" style="width: 49%;">
							<input type="submit" name="esporta" value="ESPORTA DATI" class="button" style="width: 49%;">
						</td>
					</tr>
				</table>
			</form>
		</body>
	</html>
<%else

	
	'recupera tipi di valori
	rsv.open "tb_TipNumeri", conn, adOpenStatic, adLockReadOnly, adCmdTable
	
	sessionSQL = request.querystring("sql")
	if request.querystring("sql") = "" then
		sessionSQL = "SQL_ELENCO"
	end if
	
	sql = " SELECT     IDElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi, NomeElencoIndirizzi, " & _ 	
			"CognomeElencoIndirizzi, IndirizzoElencoIndirizzi, att_id, att_oggetto, " & _
            "att_dataCrea, att_dataChiusa, att_dataS, att_priorita, att_conclusa, att_pubblica, att_eredita, att_sistema, att_domanda_id, att_mittente_id,  " & _ 
            " att_pratica_id, att_inSospeso, att_utente_chiusura, att_note, att_testo, ID_admin, admin_nome, admin_cognome, admin_email, admin_login,  " & _ 
            " admin_password, admin_scadenza, admin_note, OrdScadenza, pra_codice, pra_nome " & _ 
			"FROM         Vista_Attivit WHERE     (IDElencoIndirizzi IN  " & _ 
            " (SELECT     rel_rub_ind.id_indirizzo " & _
            " FROM          rel_rub_ind INNER JOIN " & _
            " tb_rubriche ON rel_rub_ind.id_rubrica = tb_rubriche.id_Rubrica " & _
            " WHERE      (tb_rubriche.id_Rubrica IN (@numerub))))"
	
		dim elenco_rubriche
		elenco_rubriche = request("search_rubriche")
		if elenco_rubriche="" then
			
			response.end
		end if
		sql = replace(sql,"@numerub",elenco_rubriche)
		rs.Open sql, conn, AdOpenforwardOnly, adLockReadOnly, adCmdText
	
	Server.ScriptTimeout = 3600
	  
	Select case request("Export")
		case "XLS_XP"
			CALL ExportRecordset_EXCEL_XML(rs)
		case "XLS_2000"
			CALL Export_Excel2000(rs)
		case "HTML"
			CALL ExportRecordset_HTML(rs)
		case "TXT"
			CALL ExportRecordset_TXT(rs)
		case "XML"
			CALL Export_XML(rs, false)
		case else
			response.redirect "AttivitaExport.asp"
	end select
	
	rs.close
	rsv.close

end if
	conn.close 
	set rs = nothing
	set rsv = nothing
	set conn = nothing
%>
