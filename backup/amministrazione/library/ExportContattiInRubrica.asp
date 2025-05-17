<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = true %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/TOOLS4ADMIN.ASP" -->
<!--#include file="../nextCom/Imports/Tools_Import.asp"-->
<%
dim conn
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
	
if request("esporta")="" then%>
	<html>
		<head>
			<title>Crea Rubrica</title>
			<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
			<meta name="robots" content="noindex,nofollow" />
			<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
			<link rel="stylesheet" type="text/css" href="../library/stili.css">
			<SCRIPT LANGUAGE="javascript"  src="../library/utils.js" type="text/javascript"></SCRIPT>
		</head>
		
		<body onload="window.focus();" leftmargin="4" topmargin="3">
			<form action="" method="post" id="form2" name="form2">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="width:650px;">
					<caption>
						<table width="100%" border="0" cellspacing="0">
							<tr>
								<td class="caption">Esporta contatti in una rubrica</td>
								<td align="right" style="padding-right:5px;"><a class="button" href="javascript:window.close();">CHIUDI</a></td>
							</tr>
						</table>
					</caption>
					<% if Session("Messaggio")<>"" OR request.querystring("RUBRICA")<>"" OR request.querystring("N_CONTATTI")<>"" then %>
						<th colspan="2">Import contatti completato</th>
						<tr>
							<td class="label" style="width:28%;">nome rubrica</td>
							<td class="content"><%= request.querystring("RUBRICA")%></td>
						</tr>
						<tr>
							<td class="label">n. contatti associati</td>
							<td class="content"><%= request.querystring("N_CONTATTI")%></td>
						</tr>
						<tr>
							<td class="content_center" colspan="2"><br><b><%= Session("Messaggio") %><b><br>&nbsp;</td>
						</tr>
					<% else %>
						<th>Scegli una rubrica</th>
						<% if Session("Avviso")<>"" then %>
							<tr>
								<td class="content" colspan="2"><%= Session("Avviso") %></td>
							</tr>
						<% end if %>
						<tr>
							<td colspan="2">
								<table cellspacing="1" cellpadding="0" style="border:0px; width:100%;">
									<tr>
										<td class="label" style="width:22%;">n. contatti da esportare:</td>
										<td class="label_no_width" colspan="2">
											<% dim n_contatti
											sessionSQL = cString(Session("sql_export_in_rubrica"))
											campo_id = cString(Session("campo_id_export_in_rubrica"))
											
											sql = "SELECT COUNT(DISTINCT "&campo_id&") " & right(sessionSQL, len(sessionSQL) + 1 - instr(1, sessionSQL, "FROM", vbTextCompare))
											if instrRev(sql, "ORDER BY", vbTrue,vbTextCompare) > 0 then
												sql = left(sql, instrRev(sql, "ORDER BY", vbTrue,vbTextCompare) - 1)
											end if
											if instrRev(sql, "GROUP BY", vbTrue,vbTextCompare) > 0 then
												sql = left(sql, instrRev(sql, "GROUP BY", vbTrue,vbTextCompare) - 1)
											end if
											response.write "<!-- session sql:" & sessionSQL & "-->" & vbCrlf
											response.write "<!-- sql contatti: " & sql & "-->"
											n_contatti = GetValueList(conn, NULL, sql)
											n_contatti = cIntero(n_contatti)
											%>
											<%=n_contatti%>
											<input type="hidden" name="n_contatti" value="<%=n_contatti%>">
										</td>
									</tr>										
									<tr>
										<td class="label" rowspan="2">rubrica di destinazione:</td>
										<td class="label_no_width" style="width:12%;">
											<input type="radio" class="checkbox" name="sel_tipo_rubrica" id="sel_tipo_rubrica_0" <%=chk(cInteger(request("sel_tipo_rubrica"))=0)%> value="0" onclick="SetStato_TIPO()">
											esistente:
										</td>
										<td class="content">
											<% sql = " SELECT id_rubrica, nome_rubrica FROM tb_rubriche ORDER BY nome_rubrica"
											CALL dropDown(conn, sql, "id_rubrica", "nome_rubrica", "rubrica_import", request("rubrica_import"), false, "", LINGUA_ITALIANO)%>(*)<br>
											<span class="note">Selezionare la rubrica nella quale verranno inseriti i contatti.</span>
										</td>
									</tr>  
									<tr>
										<td class="label_no_width">
											<input type="radio" class="checkbox" name="sel_tipo_rubrica" id="sel_tipo_rubrica_1" <%=chk(cInteger(request("sel_tipo_rubrica"))=1)%> value="1" onclick="SetStato_TIPO()">
											nuova:
										</td>
										<td class="content">
											<input type="text" class="text" name="nuova_rubrica" value="<%= request("nuova_rubrica") %>" maxlength="250" size="65">
										</td>
									</tr>
									<tr>
										<td colspan="3">
											<table cellspacing="1" cellpadding="0" id="option" style="border:0px;  width:100%;">
												<tr>
													<td class="label" rowspan="2" style="width:22%;">opzioni di import:</td>
													<td class="label_no_width">
														<input type="radio" class="checkbox" name="sel_option" id="sel_option_0" <%=chk(cInteger(request("sel_option"))=0)%> value="0" onclick="">
														aggiungi i nuovi contatti a quelli gi&agrave; presenti nella rubrica.
													</td>
												</tr>  
												<tr>
													<td class="label_no_width">
														<input type="radio" class="checkbox" name="sel_option" id="sel_option_1" <%=chk(cInteger(request("sel_option"))=1)%> value="1" onclick="">
														svuota la rubrica prima di inserire quelli nuovi.
													</td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
								<script language="JavaScript1.1" type="text/javascript">
									function SetStato_TIPO(){
										var option = document.getElementById('option');
										EnableIfChecked(form2.sel_tipo_rubrica_0, form2.rubrica_import);
										DisableControl(form2.nuova_rubrica, form2.sel_tipo_rubrica_0.checked);
										if (form2.sel_tipo_rubrica_0.checked)
										{
											option.style.visibility = "visible";
											option.style.display = "block";
										}
										else
										{
											option.style.visibility = "hidden";
											option.style.display = "none";
										}
									}
									
									SetStato_TIPO();
								</script>
							</td>
						</tr>
					<% end if %>
					<tr>
						<td class="footer" style="text-align:right;" colspan="2">
							<% if Session("Messaggio")="" then %>
								<input onclick="window.close();" type="button" name="esporta" value="ANNULLA" class="button" style="width: 20%;">
								<input type="submit" <%=IIF(n_contatti>0,"","disabled title=""nessun contatto selezionato""")%> name="esporta" value="ESPORTA" class="button" style="width: 20%;">
							<% else %>
								<input onclick="window.close();" type="button" name="esporta" value="CHIUDI" class="button" style="width: 20%;">
							<% end if 
							Session("Messaggio") = "" 
							Session("Avviso") = ""
							%>
						</td>
					</tr>
				</table>
			</form>
		</body>
	</html>

<% else

	dim rs, rsv, sql, sessionSQL, id_r, is_nuova, id_gruppo_lavoro, campo_id, nome_rubica
	
	if cIntero(request("sel_tipo_rubrica")) = 0 then 
		is_nuova = false
	elseif cIntero(request("sel_tipo_rubrica")) = 1 then
		is_nuova = true
	end if

	if (not is_nuova AND cIntero(request("rubrica_import"))>0) OR _
			(is_nuova AND request("nuova_rubrica")<>"") then
		'apertura transazione di creazione
		conn.begintrans
		set rs = Server.CreateObject("ADODB.RecordSet")
		
		if is_nuova then
			sql = "SELECT * FROM tb_rubriche WHERE nome_Rubrica LIKE '"+ParseSQL(request("nuova_rubrica"), adChar)+"'"
		else
			sql = "SELECT * FROM tb_rubriche WHERE id_Rubrica = " & cIntero(request("rubrica_import"))
		end if
		
		rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
		if rs.eof then
			rs.AddNew()
			rs("nome_Rubrica") = ParseSQL(request("nuova_rubrica"), adChar)
			rs("note_Rubrica") = "Rubrica salva ricerca"
			rs("locked_rubrica") = false
			'if cIntero(Session("id_sito")) = cIntero(NEXTCOM) then
				rs("rubrica_esterna") = false
			'else
			'	rs("rubrica_esterna") = true
			'end if
			rs.update
			id_r = rs("id_Rubrica")
			
			'associo gruppo di lavoro
			if cIntero(Application("NextCom_DefaultWorkGroup"))=0 then
				id_gruppo_lavoro = cIntero(GetValueList(conn,NULL,"SELECT TOP 1 nome_gruppo FROM tb_gruppi"))
			else
				id_gruppo_lavoro = cIntero(Application("NextCom_DefaultWorkGroup"))
			end if
		else
			id_r = rs("id_Rubrica") 'La rubrica esiste già
		end if
		nome_rubica = rs("nome_Rubrica")
		rs.close
		
	
		sessionSQL = cString(Session("sql_export_in_rubrica"))
		campo_id = cString(Session("campo_id_export_in_rubrica"))

		
		if not is_nuova AND cIntero(request("sel_option"))=1 then
			'cancello i contatti associati alla rubrica prima di inserire quelli nuovi
			sql = "DELETE FROM rel_rub_ind WHERE id_rubrica = " & id_r 
			conn.execute(sql)
		end if

		'inserisco associazione rubrica-contatto
		sql = "SELECT DISTINCT " & id_r & ", "&campo_id&" " & right(sessionSQL, len(sessionSQL) + 1 - instr(1, sessionSQL, "FROM", vbTextCompare))
		if instrRev(sql, "ORDER BY", vbTrue,vbTextCompare) > 0 then
			sql = left(sql, instrRev(sql, "ORDER BY", vbTrue,vbTextCompare) - 1)
		end if
		if instrRev(sql, "GROUP BY", vbTrue,vbTextCompare) > 0 then
			sql = left(sql, instrRev(sql, "GROUP BY", vbTrue,vbTextCompare) - 1)
		end if
		sql = "INSERT INTO rel_rub_ind (id_rubrica,id_indirizzo)  " & sql &_
			" AND "&campo_id&" NOT IN (SELECT id_indirizzo FROM rel_rub_ind WHERE id_rubrica IN ( " & id_r & " )) "
		conn.execute(sql)
		
		'ripulisco rel_rub_ind togliendo le righe nelle quali id_indirizzo è NULL 
		sql = " DELETE FROM rel_rub_ind WHERE id_rubrica = "& id_r &" AND ISNULL(id_indirizzo,0)=0 " 
		conn.execute(sql)

		
		if is_nuova then
			Session("Messaggio") ="Rubrica creata correttamente"
		else
			Session("Messaggio") ="Rubrica aggiornata correttamente"
		end if
		'chiusura transazione di import
		conn.committrans
	else
		Session("Avviso") = "ATTENZIONE! Inserire i dati mancanti."
	end if
	conn.close 
	set rs = nothing
	set conn = nothing
	response.redirect "ExportContattiInRubrica.asp?RUBRICA="&nome_rubica&"&N_CONTATTI="&request("n_contatti")
	%>
	<script language="JavaScript" type="text/javascript">
	//	window.location.href = "ExportContattiInRubrica.asp?RUBRICA="+'<%=nome_rubica%>'+"&N_CONTATTI="+'<%=request("n_contatti")%>';
	</script>
	<%	
end if

%>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>

