<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<%
dim conn, rsr, rst, sql, i, sql_campagne
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rsr = Server.CreateObject("ADODB.RecordSet")
set rst = Server.CreateObject("ADODB.RecordSet")

'cancello l'associazione aiivita' - campagna
if cIntero(request("REMOVE_CAMPAGNA"))>0 and cIntero(request("MOD_ID")) > 0 then
	sql = "UPDATE tb_Indirizzario_attivita SET ina_campagna_conclusa_id = 0 WHERE ina_id = " & cIntero(request("MOD_ID"))
	conn.execute(sql)
	sql = "UPDATE rel_cnt_campagne SET rcc_data_conclusione = NULL " & _
		  "WHERE rcc_campagna_id = " & cIntero(request("REMOVE_CAMPAGNA")) & " AND rcc_cnt_id = " & cIntero(request("ID"))
	conn.execute(sql)
	response.redirect "ContattiAttivita_iFrame.asp?ID=" & request("ID") & "&MOD_ID=" & request("MOD_ID")
end if


if request("SALVA")<>"" OR Request.ServerVariables("REQUEST_METHOD")="POST" then

	'modifica esistenti
	dim inaIds
	inaIds = split(replace(request("ina_id"), " ", ""), ",")
	for each i in inaIds 
		'modifica la riga
		CALL SalvaCampiEsterniUltra(conn, rsr, _
									"SELECT * FROM tb_Indirizzario_attivita", _
									"ina_id", i, "ina_anagrafica_id", request("id"), _
									"mod_"&i&"_C_ina_preso_appuntamento;mod_"&i&"_C_ina_non_interessati;mod_"&i&"_C_ina_non_raggiungibili;mod_"&i&"_C_ina_da_richiamare;mod_"&i&"_C_ina_richiamare_fatto;mod_"&i&"_C_ina_appuntamento_fatto;", _
									null, request.Form, "mod_" & i & "_")
									
		if cIntero(request("mod_"&i&"_n_ina_campagna_conclusa_id"))>0 then
			sql = "UPDATE rel_cnt_campagne SET rcc_data_conclusione = " & SQL_date(conn, Now()) & _
				  "WHERE rcc_campagna_id = " & request("mod_"&i&"_n_ina_campagna_conclusa_id") & " AND rcc_cnt_id = " & cIntero(request("ID"))
			conn.execute(sql)
		end if		
	next
	
	'salvo i campi nuovi
	i = 1
	if cIntero(request("new_"&i&"_n_ina_tipo_id"))>0 then
		CALL SalvaCampiEsterniUltra(conn, rsr, _
											"SELECT * FROM tb_Indirizzario_attivita", _
											"ina_id", 0, "ina_anagrafica_id", request("id"), _
											"new_"&i&"_C_ina_preso_appuntamento;new_"&i&"_C_ina_non_interessati;new_"&i&"_C_ina_non_raggiungibili;new_"&i&"_C_ina_da_richiamare;", _
											null, request.Form, "new_" & i & "_")
											
		if cIntero(request("new_"&i&"_n_ina_campagna_conclusa_id"))>0 then
			sql = "UPDATE rel_cnt_campagne SET rcc_data_conclusione = " & SQL_date(conn, Now()) & _
				  "WHERE rcc_campagna_id = " & request("new_"&i&"_n_ina_campagna_conclusa_id") & " AND rcc_cnt_id = " & cIntero(request("ID"))
			conn.execute(sql)
		end if										
	end if
	
	response.redirect "ContattiAttivita_iFrame.asp?ID=" & request("ID")
end if
%>
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<% '*******************************************************************************************************************************
ParentFrameName = "IFrameContattiAttivita" 
%>
<!--#INCLUDE FILE="../library/Intestazione_iframe.asp" -->
<% '*******************************************************************************************************************************


'gestione contatti interni
sql = "SELECT * FROM tb_Indirizzario_attivita " & _
	  " WHERE ina_anagrafica_id=" & request("ID") & " ORDER BY ina_insData DESC, ina_id DESC "
rsr.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText

sql_campagne = " SELECT * FROM tb_indirizzario_campagne WHERE inc_id IN " & _
			   " (SELECT rcc_campagna_id FROM rel_cnt_campagne WHERE rcc_cnt_id = "&request("ID")&" AND rcc_data_conclusione IS NULL)"
								 
%>
<form action="" method="post" id="form1" name="form1"  style="margin-bottom:0px;">
	<table cellspacing="1" cellpadding="0" class="tabella_madre contattiattivita" style="border-right:0px; border-left:0px; border-bottom:0px;" style="width:100% !important;">
	<tr>
		<th colspan="4">ATTIVITA' DI CONTATTO</th>
	</tr>
	<% if request("ADD_NEW")="" and request("MOD_ID")="" then
		%>
		<tr>
			<td colspan="4" class="content_right">
				aggiungi:&nbsp;
				<%
				sql = "SELECT * FROM tb_indirizzario_attivita_tipi ORDER BY iat_ordine "
				rst.open sql, conn
				while not rst.eof %>
					<a class="button_L2" href="javascript:void(0);" onclick="window.location = 'ContattiAttivita_iFrame.asp?ID=<%=request("id")%>&ADD_NEW=true&new_1_n_ina_tipo_id=<%=rst("iat_id")%>';">
						<%=rst("iat_nome")%>
					</a>
					&nbsp;
					<%
					rst.moveNext				
				wend 
				rst.close
				%>
			</td>
		</tr>
	<% else %>
		<tr>
			<td colspan="4" class="content_right">
				&nbsp;
			</td>
		</tr>
	<% end if %>
	<tr>
		<td colspan="4">
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
				<% if request("ADD_NEW")<>"" then
					%>
					<%
					'NUOVO RECORD
					%>
					<% i = 1 %>
					<input type="hidden" name="new_<%=i%>_n_ina_anagrafica_id" value="<%=request("id")%>">
					<input type="hidden" name="new_<%=i%>_n_ina_insAdmin_id" value="<%=Session("ID_ADMIN")%>">
					<tr>
						<td class="content" style="padding:0px;" rowspan="3">
							<table cellspacing="0" cellpadding="0">
								<% 
								sql = "SELECT * FROM tb_indirizzario_attivita_tipi ORDER BY iat_ordine "
								rst.open sql, conn
								while not rst.eof
									%>
									<tr>
										<td style="width:17px; text-align:center;">
											<img src="<%= GetSiteUrl(null, 0, 0) & rst("iat_icona") %>" />
										</td>
										<td>
											<input type="radio" class="noborder" style="margin-right:0px;" name="new_<%=i%>_n_ina_tipo_id" <%=chk(cIntero(request("new_"&i&"_n_ina_tipo_id"))=rst("iat_id"))%> value="<%=rst("iat_id")%>" />
											<%=rst("iat_nome")%>
										</td>
									</tr>	
									<%
									rst.moveNext
								wend
								rst.close
								%>
							</table>
						</td>
						<td class="content" rowspan="3">
							<% CALL WriteDataPicker_Input_Manuale("form1", "new_"&i&"_d_ina_insData", dateITA(Now()), "", "/", false, true, LINGUA_ITALIANO, "", false, "width:70px;") %>
						</td>
						<td class="content" colspan="2">
							<textarea style="width:100%;" rows="4" cols="" name="new_<%=i%>_t_ina_note"><%=request("new_"&i&"_t_ina_note")%></textarea>
						</td>
						<td class="content_center" rowspan="3">
							<a class="button_L2" href="javascript:void(0);" onclick="document.form1.submit();">
								SALVA
							</a>
							&nbsp;
							<a class="button_L2" href="javascript:void(0);" onclick="window.location = 'ContattiAttivita_iFrame.asp?ID=<%=request("id")%>';">
								ANNULLA
							</a>
						</td>
					</tr>
					<%
					rst.open sql_campagne, conn
					%>
					<tr>
						<td colspan="2" <%=IIF(rst.eof, "rowspan=""2""", "")%>>
							<table cellspacing="1" cellpadding="1" style="width:100%;">
								<tr>
									<td class="content" style="width:33%;">
										<input type="radio" class="noborder"  name="new_<%=i%>_C_ina_da_richiamare" <%=chk(request("new_"&i&"_C_ina_da_richiamare")=1)%> value="1" onclick="SetUniqueRadioButton('<%="new_"&i&"_C_"%>', this); SetDataField('<%="new_"&i&"_d_"%>ina_data_appuntamento');" />Da richiamare il
									</td>
									<td class="content">
										<% CALL WriteDataPicker_Input_Manuale("form1", "new_"&i&"_d_ina_data_ricontatto", request("new_"&i&"_d_ina_data_ricontatto"), "", "/", true, true, LINGUA_ITALIANO, "", false, "width:70px;") %>
									</td>
									<td class="content" style="width:27%;">
										<input type="radio" class="noborder"  name="new_<%=i%>_C_ina_non_interessati" <%=chk(request("new_"&i&"_C_ina_non_interessati")=1)%> value="1" onclick="SetUniqueRadioButton('<%="new_"&i&"_C_"%>', this);" />Non interessati
									</td>
								</tr>
								<tr>
									<td class="content">
										<input type="radio" class="noborder"  name="new_<%=i%>_C_ina_preso_appuntamento" <%=chk(request("new_"&i&"_C_ina_preso_appuntamento")=1)%> value="1" onclick="SetUniqueRadioButton('<%="new_"&i&"_C_"%>', this); SetDataField('<%="new_"&i&"_d_"%>ina_data_ricontatto');" />Preso appuntamento il
									</td>
									<td class="content">
										<% CALL WriteDataPicker_Input_Manuale("form1", "new_"&i&"_d_ina_data_appuntamento", request("new_"&i&"_d_ina_data_appuntamento"), "", "/", true, true, LINGUA_ITALIANO, "", false, "width:70px;") %>
									</td>
									<td class="content">
										<input type="radio" class="noborder"  name="new_<%=i%>_C_ina_non_raggiungibili" <%=chk(request("new_"&i&"_C_ina_non_raggiungibili")=1)%> value="1" onclick="SetUniqueRadioButton('<%="new_"&i&"_C_"%>', this);" />Non raggiungibili
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<% if not rst.eof then %>
						<tr>
							<td colspan="3">
								<table cellspacing="0" cellpadding="1" style="width:100%;">
									<tr>
										<td class="content" style="width:50%;">
											<input type="checkbox" class="noborder" name="pulsante_<%=i%>_abilita_campagna" value="<%=request("pulsante_"&i&"_abilita_campagna")%>" onclick="EnableIfChecked(this, document.getElementById('new_<%=i%>_n_ina_campagna_conclusa_id'));" />
											Questa attivita' fa riferimento alla campagna:
										</td>
										<td class="content">
											<% CALL dropDown(conn, sql_campagne, "inc_id", "inc_nome", "new_"&i&"_n_ina_campagna_conclusa_id", cInteger(request("new_"&i&"_n_ina_campagna_conclusa_id")), false, "", LINGUA_ITALIANO)%>
										</td>
										<% if cInteger(request("new_"&i&"_n_ina_campagna_conclusa_id")) = 0 then %>
											<script language="JavaScript" type="text/javascript">
												document.getElementById('new_<%=i%>_n_ina_campagna_conclusa_id').disabled = true;
												document.getElementById('new_<%=i%>_n_ina_campagna_conclusa_id').className = 'disabled';
											</script>
										<% end if %>
									</tr>
								</table>
							</td>
						</tr>
					<% else %>
						<tr></tr>
					<% end if %>
					<% rst.close %>
				<% end if %>
				<% if not rsr.eof then %>
					<tr>
						<th class="L2" style="width:13%;">tipo</th>
						<th class="L2" style="width:7%;">del</th>
						<th class="L2">note</th>
						<th class="L2" style="width:20%">conclusione</th>
						<th class="l2_center" width="16%" style="text-align:center;">operazioni</th>
					</tr>
				<% end if %>
				<% while not rsr.eof %>
					<%
					'MODIFICA RECORD
					%>
					<% i = rsr("ina_id") %>
					<% if cIntero(request("MOD_ID")) = rsr("ina_id") then %>
						<input type="hidden" name="ina_id" value="<%=rsr("ina_id")%>">
						<input type="hidden" name="mod_<%=i%>_d_ina_modData" value="<%=Now()%>">
						<input type="hidden" name="mod_<%=i%>_n_ina_modAdmin_id" value="<%=Session("ID_ADMIN")%>">
						<tr>
							<td class="content" style="padding:0px;" rowspan="3">
								<table cellspacing="0" cellpadding="0">
									<% 
									sql = "SELECT * FROM tb_indirizzario_attivita_tipi ORDER BY iat_ordine "
									rst.open sql, conn
									while not rst.eof
										%>
										<tr>
											<td style="width:17px; text-align:center;">
												<img src="<%= GetSiteUrl(null, 0, 0) & rst("iat_icona") %>" />
											</td>
											<td>
												<input type="radio" class="noborder"  name="mod_<%=i%>_n_ina_tipo_id" <%=chk(rsr("ina_tipo_id")=rst("iat_id"))%> value="<%=rst("iat_id")%>" />
												<%=rst("iat_nome")%>
											</td>
										</tr>	
										<%
										rst.moveNext
									wend
									rst.close
									%>
								</table>
							</td>
							<td class="content" rowspan="3">
								<% CALL WriteDataPicker_Input_Manuale("form1", "mod_"&i&"_d_ina_insData", dateITA(rsr("ina_insData")), "", "/", false, true, LINGUA_ITALIANO, "", false, "width:70px;") %>
							</td>
							<td class="content" colspan="2">
								<textarea style="width:100%;" rows="4" cols="" name="mod_<%=i%>_t_ina_note"><%=rsr("ina_note")%></textarea>
							</td>
							<td class="content_center" rowspan="3">
								<a class="button_L2" href="javascript:void(0);" onclick="document.form1.submit();" style="display:block;">
									SALVA
								</a>
								<br />
								<script language="JavaScript" type="text/javascript">
									function SetAttivita<%=i%>Conclusa(){
										form1.mod_<%=i%>_C_ina_richiamare_fatto.checked = form1.mod_<%=i%>_C_ina_da_richiamare.checked
										form1.mod_<%=i%>_C_ina_appuntamento_fatto.checked = form1.mod_<%=i%>_C_ina_preso_appuntamento.checked;
										document.form1.submit();
									}
								</script>
								<a class="button_L2" href="javascript:void(0);" onclick="SetAttivita<%=i%>Conclusa();" style="display:block;">
									SALVA E SEGNA L'ATTIVITA' COME CONSCLUSA
								</a>
								<br /><br /><br /><br /><br /><br /><br />
								<a class="button_L2" href="javascript:void(0);" onclick="window.location = 'ContattiAttivita_iFrame.asp?ID=<%=request("id")%>';" style="display:block;">
									ANNULLA
								</a>
							</td>
						</tr>
						<% 
						rst.open sql_campagne, conn
						%>
						<tr>
							<td colspan="2" <%=IIF(rst.eof, "rowspan=""2""", "")%>>
								<%
								dim disabled_style
								if cBoolean(rsr("ina_appuntamento_fatto"), false) OR cBoolean(rsr("ina_richiamare_fatto"), false) then
									disabled_style = "_disabled"
								else
									disabled_style = ""
								end if
								%>
								<script language="JavaScript" type="text/javascript">
									function RichiamareSelected(){
										document.getElementById("container_richiamare_<%=i%>").className = "content";
										form1.mod_<%=i%>_d_ina_data_ricontatto.className = "PickerDateInput";
										form1.mod_<%=i%>_d_ina_data_ricontatto.disabled = false;
										document.getElementById("container_richiamare_fatto_<%=i%>").className = "content";
										form1.mod_<%=i%>_C_ina_richiamare_fatto.className = "noborder";
										form1.mod_<%=i%>_C_ina_richiamare_fatto.disabled = false;
										//---
										document.getElementById("container_appuntamento_<%=i%>").className = "content_disabled";
										form1.mod_<%=i%>_d_ina_data_appuntamento.className += " disabled";
										form1.mod_<%=i%>_d_ina_data_appuntamento.disabled = true;
										document.getElementById("container_appuntamento_fatto_<%=i%>").className = "content_disabled";
										form1.mod_<%=i%>_C_ina_appuntamento_fatto.className += " disabled";
										form1.mod_<%=i%>_C_ina_appuntamento_fatto.disabled = true;
										//---
										document.getElementById("container_interessati_<%=i%>").className = "content_disabled";
										//--
										document.getElementById("container_raggiungibili_<%=i%>").className = "content_disabled";
									}
									
									function AppuntamentoSelected(){
										document.getElementById("container_richiamare_<%=i%>").className = "content_disabled";
										form1.mod_<%=i%>_d_ina_data_ricontatto.className += " disabled";
										form1.mod_<%=i%>_d_ina_data_ricontatto.disabled = true;
										document.getElementById("container_richiamare_fatto_<%=i%>").className = "content_disabled";
										form1.mod_<%=i%>_C_ina_richiamare_fatto.className += " disabled";
										form1.mod_<%=i%>_C_ina_richiamare_fatto.disabled = true;
										//---
										document.getElementById("container_appuntamento_<%=i%>").className = "content";
										form1.mod_<%=i%>_d_ina_data_appuntamento.className = "PickerDateInput";
										form1.mod_<%=i%>_d_ina_data_appuntamento.disabled = false;
										document.getElementById("container_appuntamento_fatto_<%=i%>").className = "content";
										form1.mod_<%=i%>_C_ina_appuntamento_fatto.className = "noborder";
										form1.mod_<%=i%>_C_ina_appuntamento_fatto.disabled = false;
										//---
										document.getElementById("container_interessati_<%=i%>").className = "content_disabled";
										//--
										document.getElementById("container_raggiungibili_<%=i%>").className = "content_disabled";
									}
									
									function InteressatiSelected(){
										document.getElementById("container_richiamare_<%=i%>").className = "content_disabled";
										form1.mod_<%=i%>_d_ina_data_ricontatto.className += " disabled";
										form1.mod_<%=i%>_d_ina_data_ricontatto.disabled = true;
										document.getElementById("container_richiamare_fatto_<%=i%>").className = "content_disabled";
										form1.mod_<%=i%>_C_ina_richiamare_fatto.className += " disabled";
										form1.mod_<%=i%>_C_ina_richiamare_fatto.disabled = true;
										//---
										document.getElementById("container_appuntamento_<%=i%>").className = "content_disabled";
										form1.mod_<%=i%>_d_ina_data_appuntamento.className += " disabled";
										form1.mod_<%=i%>_d_ina_data_appuntamento.disabled = true;
										document.getElementById("container_appuntamento_fatto_<%=i%>").className = "content_disabled";
										form1.mod_<%=i%>_C_ina_appuntamento_fatto.className += " disabled";
										form1.mod_<%=i%>_C_ina_appuntamento_fatto.disabled = true;
										//---
										document.getElementById("container_interessati_<%=i%>").className = "content";
										//--
										document.getElementById("container_raggiungibili_<%=i%>").className = "content_disabled";
									}
									
									function RaggiungibiliSelected(){
										document.getElementById("container_richiamare_<%=i%>").className = "content_disabled";
										form1.mod_<%=i%>_d_ina_data_ricontatto.className += " disabled";
										form1.mod_<%=i%>_d_ina_data_ricontatto.disabled = true;
										document.getElementById("container_richiamare_fatto_<%=i%>").className = "content_disabled";
										form1.mod_<%=i%>_C_ina_richiamare_fatto.className += " disabled";
										form1.mod_<%=i%>_C_ina_richiamare_fatto.disabled = true;
										//---
										document.getElementById("container_appuntamento_<%=i%>").className = "content_disabled";
										form1.mod_<%=i%>_d_ina_data_appuntamento.className += " disabled";
										form1.mod_<%=i%>_d_ina_data_appuntamento.disabled = true;
										document.getElementById("container_appuntamento_fatto_<%=i%>").className = "content_disabled";
										form1.mod_<%=i%>_C_ina_appuntamento_fatto.className += " disabled";
										form1.mod_<%=i%>_C_ina_appuntamento_fatto.disabled = true;
										//---
										document.getElementById("container_interessati_<%=i%>").className = "content_disabled";
										//--
										document.getElementById("container_raggiungibili_<%=i%>").className = "content";
									}
									
								</script>
								<table cellspacing="1" cellpadding="1" style="width:100%;">
									<tr>
										<td class="content" style="width:36%;" id="container_richiamare_<%=i%>">
											<input type="radio" class="noborder<%=IIF(cBoolean(rsr("ina_da_richiamare"), false)," selected","")%>"  name="mod_<%=i%>_C_ina_da_richiamare" <%=chk(cBoolean(rsr("ina_da_richiamare"), false))%> value="1" onclick="SetUniqueRadioButton('<%="mod_"&i&"_C_"%>', this); SetDataField('mod_<%=i%>_d_ina_data_appuntamento'); RichiamareSelected();" />Da richiamare il
										</td>
										<td class="content">
											<% CALL WriteDataPicker_Input_Manuale("form1", "mod_"&i&"_d_ina_data_ricontatto", rsr("ina_data_ricontatto"), "", "/", true, true, LINGUA_ITALIANO, "", false, "width:70px;") %>
										</td>
										<td class="content" id="container_richiamare_fatto_<%=i%>">
											<input type="checkbox" class="noborder" name="mod_<%=i%>_C_ina_richiamare_fatto" <%=chk(cBoolean(rsr("ina_richiamare_fatto"), false))%>>
											fatto
										</td>
									</tr>
									<tr>
										<td class="content" id="container_appuntamento_<%=i%>">
											<input type="radio" class="noborder<%=IIF(cBoolean(rsr("ina_preso_appuntamento"), false)," selected","")%>"  name="mod_<%=i%>_C_ina_preso_appuntamento" <%=chk(cBoolean(rsr("ina_preso_appuntamento"), false))%> value="1" onclick="SetUniqueRadioButton('<%="mod_"&i&"_C_"%>', this); SetDataField('mod_<%=i%>_d_ina_data_ricontatto'); AppuntamentoSelected();" />Preso appuntamento il
										</td>
										<td class="content">
											<% CALL WriteDataPicker_Input_Manuale("form1", "mod_"&i&"_d_ina_data_appuntamento", rsr("ina_data_appuntamento"), "", "/", true, true, LINGUA_ITALIANO, "", false, "width:70px;") %>
										</td>
										<td class="content" id="container_appuntamento_fatto_<%=i%>">
											<input type="checkbox" class="noborder" name="mod_<%=i%>_C_ina_appuntamento_fatto" <%=chk(cBoolean(rsr("ina_appuntamento_fatto"), false))%>>
											fatto
										</td>
									</tr>
									<tr>
										<td class="content" colspan="3" id="container_interessati_<%=i%>">
											<input type="radio" class="noborder<%=IIF(cBoolean(rsr("ina_non_interessati"), false)," selected","")%>"  name="mod_<%=i%>_C_ina_non_interessati" <%=chk(cBoolean(rsr("ina_non_interessati"), false))%> value="1" onclick="SetUniqueRadioButton('<%="mod_"&i&"_C_"%>', this); InteressatiSelected();" />Non interessati
										</td>
									</tr>
									<tr>
										<td class="content" colspan="3" id="container_raggiungibili_<%=i%>">
											<input type="radio" class="noborder<%=IIF(cBoolean(rsr("ina_non_raggiungibili"), false)," selected","")%>"  name="mod_<%=i%>_C_ina_non_raggiungibili" <%=chk(cBoolean(rsr("ina_non_raggiungibili"), false))%> value="1" onclick="SetUniqueRadioButton('<%="mod_"&i&"_C_"%>', this); RaggiungibiliSelected();" />Non raggiungibili
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<script language="JavaScript" type="text/javascript">
							<% if cBoolean(rsr("ina_da_richiamare"), false) then %>
								RichiamareSelected();
							<% elseif cBoolean(rsr("ina_preso_appuntamento"), false) then %>
								AppuntamentoSelected();
							<% elseif cBoolean(rsr("ina_non_interessati"), false) then %>
								InteressatiSelected();
							<% elseif cBoolean(rsr("ina_non_raggiungibili"), false) then %>
								RaggiungibiliSelected();
							<% end if %>
						</script>
						<% if not rst.eof then %>
							<tr>
								<td colspan="2">
									<table cellspacing="0" cellpadding="1" style="width:100%;">
										<tr>
											<% if cIntero(rsr("ina_campagna_conclusa_id")) > 0 then %>
												<td class="content" style="width:48%;">
													Questa attivita' fa riferimento alla campagna:
												</td>
												<td class="content_b">
													<%= GetValueList(conn, NULL, "SELECT inc_nome FROM tb_indirizzario_campagne WHERE inc_id = " & cIntero(rsr("ina_campagna_conclusa_id"))) %>
												</td>
												<td class="content"  style="width:12%;">
													<a class="button_L2" href="javascript:void(0);" onclick="window.location = 'ContattiAttivita_iFrame.asp?ID=<%=request("id")%>&MOD_ID=<%=request("MOD_ID")%>&REMOVE_CAMPAGNA=<%=rsr("ina_campagna_conclusa_id")%>';">
														ANNULLA
													</a>
												</td>
											<% else %>
												<td class="content" style="width:55%;">
													<input type="checkbox" class="noborder" name="pulsante_<%=i%>_abilita_campagna" value="<%=request("pulsante_"&i&"_abilita_campagna")%>" onclick="EnableIfChecked(this, document.getElementById('mod_<%=i%>_n_ina_campagna_conclusa_id'));" />
													Questa attivita' fa riferimento alla campagna:
												</td>
												<td class="content">
													<% CALL dropDown(conn, sql_campagne, "inc_id", "inc_nome", "mod_"&i&"_n_ina_campagna_conclusa_id", cInteger(request("mod_"&i&"_n_ina_campagna_conclusa_id")), false, "", LINGUA_ITALIANO)%>
												</td>
												<% if cInteger(request("mod_"&i&"_n_ina_campagna_conclusa_id")) = 0 then %>
													<script language="JavaScript" type="text/javascript">
														document.getElementById('mod_<%=i%>_n_ina_campagna_conclusa_id').disabled = true;
														document.getElementById('mod_<%=i%>_n_ina_campagna_conclusa_id').className = 'disabled';
													</script>
												<% end if %>
											<% end if %>
										</tr>
									</table>
								</td>
							</tr>
						<% else %>
							<tr></tr>
						<% end if %>
						<% rst.close %>
					<% else %>
						<%
						'VISUALIZZA RECORD
						%>
						<tr>
							<td class="content" style="padding-bottom:10px;">
								<table cellspacing="0" cellpadding="0">
									<% sql = "SELECT * FROM tb_indirizzario_attivita_tipi WHERE iat_id = " & rsr("ina_tipo_id")
									rst.open sql, conn
									%>
									<tr>
										<td style="width:22px; text-align:left;">
											<img src="<%= GetSiteUrl(null, 0, 0) & rst("iat_icona") %>" />
										</td>
										<td>
											<%=rst("iat_nome") %>
										</td>
									</tr>
									<% rst.close%>
								</table>
							</td>
							<td class="content">
								<%=dateITA(rsr("ina_insData"))%>
							</td>
							<td class="content">
								<%=TextEncode(rsr("ina_note"))%>
							</td>
							<td class="content" style="vertical-align:top;">
								<table cellspacing="0" cellpadding="0" style="width:100%;">
									<tr>
										<td style="width:22px;">
											<% 
											dim scaduto 
											scaduto = false
											%>
											<% if cBoolean(rsr("ina_richiamare_fatto"), false) OR cBoolean(rsr("ina_appuntamento_fatto"), false) then %>
												<img src="<%= GetSiteUrl(null, 0, 0) & "/amministrazione/grafica/attivita-fatta.jpg" %>" />
											<% elseif cBoolean(rsr("ina_da_richiamare"), false) OR cBoolean(rsr("ina_preso_appuntamento"), false) then %>
												<% if (cBoolean(rsr("ina_da_richiamare"), false) AND (isDate(rsr("ina_data_ricontatto")) AND DateISO(rsr("ina_data_ricontatto")) < DateISO(Now()))) _
													OR (cBoolean(rsr("ina_preso_appuntamento"), false) AND (isDate(rsr("ina_data_appuntamento")) AND DateISO(rsr("ina_data_appuntamento")) < DateISO(Now()))) then 
													scaduto = true %>
													<img src="<%= GetSiteUrl(null, 0, 0) & "/amministrazione/grafica/attivita-dimenticata.gif" %>" />
												<% else %>
													<img src="<%= GetSiteUrl(null, 0, 0) & "/amministrazione/grafica/attivita-da-fare.jpg" %>" />
												<% end if %>
											<% end if %>
										</td>
										<% 
										dim output
										output = "-"
										if cBoolean(rsr("ina_da_richiamare"), false) then
											output = "Da richiamare"
											if rsr("ina_data_ricontatto")<>"" then
												output = output & " il " & rsr("ina_data_ricontatto")
											end if
											if cBoolean(rsr("ina_richiamare_fatto"), false) then
												output = "<strike>" & output & "</strike>"
											end if
										elseif cBoolean(rsr("ina_preso_appuntamento"), false) then
											output = "Preso appuntamento"
											if rsr("ina_data_appuntamento")<>"" then
												output = output & " il " & rsr("ina_data_appuntamento")
											end if
											if cBoolean(rsr("ina_appuntamento_fatto"), false) then
												output = "<strike>" & output & "</strike>"
											end if
										elseif cBoolean(rsr("ina_non_interessati"), false) then
											output = "Non interessati"
										elseif cBoolean(rsr("ina_non_raggiungibili"), false) then
											output = "Non raggiungibili"
										end if
										
										if scaduto then
											output = "<b>" & output & "</b>"
										end if
										%>
										<td>
											<%
											response.write output
											%>
										</td>
									</tr>
									<% if cIntero(rsr("ina_campagna_conclusa_id")) > 0 then %>
										<tr>
											<td style="background-color:#b9ddb9;" colspan="2">
												campagna: <b><%=GetValueList(conn, NULL, "SELECT inc_nome FROM tb_indirizzario_campagne WHERE inc_id = " & cIntero(rsr("ina_campagna_conclusa_id"))) %></b>
											</td>
										</tr>
									<% end if %>
								</table>
							</td>
							<td class="content_center">
								<a class="button_L2" href="javascript:void(0);" onclick="window.location = 'ContattiAttivita_iFrame.asp?ID=<%=request("id")%>&MOD_ID=<%=rsr("ina_id")%>';">
									MODIFICA
								</a>
								&nbsp;
								<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('CONTATTI_ATTIVITA','<%= rsr("ina_id") %>&CNT_ID=<%=request("ID")%>');">
									CANCELLA
								</a>
							</td>
						</tr>
					<% end if %>
					<%rsr.movenext
				wend
				%>
				<script language="JavaScript" type="text/javascript">
					function SetUniqueRadioButton(nameregex, current){
						var varTemp;
						varTemp = document.getElementBy
						toBeDisabled = false;
						re = new RegExp(nameregex);
						for(i = 0; i < document.forms[0].elements.length; i++){
							elm = document.forms[0].elements[i]
							if (elm.type == 'radio'){
								if (re.test(elm.name) && elm.name != current.name){ //disabilito tutti i radio button, tranne il selezionato
									elm.checked = false;
									elm.className = elm.className.replace(' selected', '');
								}
								
							}
						}
						if (current.className.indexOf(' selected') > 0){
							current.checked = false;
							current.className = current.className.replace(' selected', '');
						}
						else{
							current.checked = true; // abilito il radiobutton cliccato
							current.className = current.className + ' selected';
						}
					}

					function SetDataField(dataFieldToDelete)
					{
						var elm;
						elm = document.getElementById(dataFieldToDelete);
						elm.value = '';
					}
				</script>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="4" class="content_right">
			&nbsp;
		</td>
	</tr>
	</table>
</form>
</div>
</body>
</html>
<% 
rsr.close
conn.close 
set rsr = nothing
set rst = nothing
set conn = nothing

%>
