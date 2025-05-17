<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if (request("salva")<>"" OR request("salva_calendario")<>"" OR request("salva_elenco")<>"") AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ImpegniSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura, data
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione impegni/appuntamenti - modifica"
dicitura.puls_new = "INDIETRO"
if cString(request("RETURN_DATE"))<>"" then
	dicitura.link_new = "ImpegniCalendarioView.asp?FIRSTDATE=" & cString(request("RETURN_DATE"))
else
	dicitura.link_new = "Impegni.asp"
end if
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")

dim intervallo
'intervallo = cIntero(Session("AGENDA_INTERVALLO_CALENDARIO"))
'if intervallo = "" then
	intervallo = 5
'end if

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("SQL_IMPEGNI"), "imp_id", "ImpegniMod.asp")
end if

sql = "SELECT * FROM mtb_impegni WHERE imp_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<input type="hidden" name="first_day_week" value="<%=cString(request("RETURN_DATE"))%>">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati impegno/appuntamento</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="impegno precedente">
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="inpegno successivo">
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
			<% 	end if %>
				<td class="content" colspan="3">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_imp_titolo_<%= Application("LINGUE")(i) %>" value="<%= rs("imp_titolo_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
					<% 	if i = 0 then %>(*)<% end if %>
				</td>
			</tr>
		<% next %>
		
		<%
		sql = "SELECT * FROM mtb_tipi_impegni ORDER BY tim_nome_it"
		if GetValueList(conn, NULL, sql) <>"" then %>
		<tr>
			<td class="label" style="width:20%;">tipologia:</td>
			<td class="content" colspan="3">
				<% CALL dropDown(conn, sql, "tim_id", "tim_nome_it", "tfn_imp_tipo_id", rs("imp_tipo_id"), true, "", Session("LINGUA")) %>
			</td>
		</tr>
		<% end if %>
		
		<tr>
			<td class="label">visibilit&agrave;:</td>
			<td class="content" colspan="3">
				<table border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td><input class="noBorder" type="radio" name="imp_protetto" id="chk_imp_protetto_false" value="" <%=chk(cBoolean(not rs("imp_protetto"), false))%> onClick=""></td>
						<td width="10%">
							pubblico
						</td>
						<td><input class="noBorder" type="radio" name="imp_protetto" id="chk_imp_protetto_true" value="1" <%=chk(cBoolean(rs("imp_protetto"), false))%> onClick=""></td>
						<td width="10%">
							<img src="../grafica/padlock.gif" border="0" alt="Pagina appartenente all'area protetta">
							privato
						</td>
						<td class="label" style="text-align:right; width:80%; vertical-align:middle;">Se protetto, l'impegno sar&agrave; visibile solo agli utenti o ai profili associati ad esso.</td>
					</tr>
				</table>
			</td>
		</tr>
		
		<tr>
			<td class="label">orario inizio:</td>
			<td class="content" colspan="3">
				<%CALL WriteDropDownOrario("orario_inizio",intervallo,TimeIta(rs("imp_data_ora_inizio")),"","")%>
			</td>
		</tr>	
		<tr>
			<td class="label">orario fine:</td>				
			<td class="content" colspan="3">
				<%CALL WriteDropDownOrario("orario_fine",intervallo,TimeIta(rs("imp_data_ora_fine")),"","")%>
			</td>
		</tr>
		<tr>
			<td class="label">data inizio:</td>		
			<td class="content" colspan="3">
				<% CALL WriteDataPicker_Input("form1", "data_inizio",DateIta(rs("imp_data_ora_inizio")), "", "/", false, true, LINGUA_ITALIANO) %>
			</td>
		</tr>
		<tr>
			<td class="label">data fine:</td>		
			<td class="content">
				<% CALL WriteDataPicker_Input("form1", "data_fine", IIF(Trim(DateIta(rs("imp_data_ora_fine")))<>Trim(DateIta(DATA_SENZA_FINE)),DateIta(rs("imp_data_ora_fine")),""), "", "/", true, true, LINGUA_ITALIANO) %>
			</td>
			<td class="label" style="text-align:right; width:80%;" colspan="2">
				Se la data di fine non corrisponde alla data di inizio, l'impegno verr&agrave; ripetuto in tutti i giorni compresi nell'intervallo data inizio - data fine, estremi inclusi.
			</td>
		</tr>
		
		<tr><th colspan="4">SPEDIZIONE AVVISI VIA E-MAIL</th></tr>
		<tr>
			<td class="label" rowspan="2">avvisi:</td>
			<td class="content" colspan="2">
				<input type="checkbox" class="checkbox" name="avviso_dopo_salvataggio" value="1">
				spedisci avviso dopo il salvataggio
			</td>
			<td colspan="1" class="content_right" style="padding-right:0px;width:22%;">
				<a class="button_L2" style="width:100%;text-align:center;" href="javascript:void(0);" 
						onclick="OpenPositionedWindow('ImpegniSendAvvisiNow.asp?ID=<%=request("ID")%>', '_blank', '300', '300','400px', '200px');" 
								title="spedisco ora le e-mail di avviso agli utenti interessati dall'impegno">
					SPEDISCI AVVISI
				</a>
			</td>
		</tr>
		<tr>
			<td class="content" colspan="3">
				<input type="checkbox" class="checkbox" name="chk_imp_invia_avviso" value="1" <%=chk(cBoolean(rs("imp_invia_avviso"), false))%> >
				spedisci avviso a
				<% CALL WriteDropDownMinutiAnticipo("tfn_imp_anticipo_avviso",rs("imp_anticipo_avviso")) %>
				dall'impegno.
			</td>
		</tr>

		<tr><th colspan="4">UTENTI IMPEGNATI</th></tr>		
		<% sql = "SELECT pro_id FROM mtb_profili" %>
		<% if cString(GetValueList(conn, NULL, sql)) <> "" then %>
			<tr>
				<td class="label" style="width:20%;">profili:</td>
				<td class="content" colspan="3">
					<% dim rsp
					   set rsp = Server.CreateObject("ADODB.RecordSet")
					   sql = " SELECT * FROM mtb_profili LEFT JOIN mrel_impegni_profili " + _
							 " ON (mtb_profili.pro_id = mrel_impegni_profili.rip_profilo_id AND mrel_impegni_profili.rip_impegno_id = " & cIntero(request("ID")) & ")" + _
							 " ORDER BY pro_nome_it"
					   CALL Write_Relations_Checker(conn, rsp, sql, 3, "pro_id", "pro_nome_it", "rip_id", "profili_associati")%>
				</td>
			</tr>
		<% end if %>
		<% if cBoolean(Session("CONDIVISIONE_PUBBLICA"), false) then %>
			<tr>
				<td class="label">utenti area riservata:</td>
				<td class="content" colspan="3">
					<% dim utenti
					   sql = "SELECT ut_NextCom_id FROM tb_utenti WHERE ut_ID IN (SELECT riu_utente_id FROM mrel_impegni_utenti WHERE riu_impegno_id = " & cIntero(request("ID")) & ")"
					   utenti = GetValueList(conn, NULL, sql)
					%>
					<% CALL WriteContactPicker_Input(conn, NULL, " ut_ID IN (SELECT rel_ut_id FROM rel_utenti_sito WHERE rel_permesso = 1 OR rel_permesso = 2) ", "", "form1", "utenti_associati", utenti, "LOGINMANDATORY", true, false, false, "")  %>
				</td>
			</tr>
		<% end if %>
		<tr><th colspan="4">ALTRI DATI</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label_no_width" rowspan="<%= ubound(Application("LINGUE"))+1 %>">descrizione:</td>
				<% end if %>
				<td class="content" colspan="3">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="16" alt="" border="0" style="vertical-align: top;">
					<textarea name="tft_imp_descrizione_<%= Application("LINGUE")(i) %>" 
								style="width:94.5%;" rows="4"><%= rs("imp_descrizione_"& Application("LINGUE")(i)) %></textarea>
				</td>
			</tr>
		<%next %>
		
		<tr>
			<%
			CALL Form_DatiModifica(conn, rs, "imp_")
			%>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
				<% if request("RETURN_DATE")<> "" then %>
					<input type="submit" class="button" name="salva_calendario" value="SALVA & TORNA AL CALENDARIO">
				<% else %>
					<input type="submit" class="button" name="salva_elenco" value="SALVA & TORNA ALL'ELENCO">
				<% end if %>
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
