<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if (request("salva")<>"" OR request("salva_continua")<>"") AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("FattureSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione fatture - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Fatture.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, rsa, is_post, i, sql, new_numero, anno
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")

is_post = cBoolean(Request.ServerVariables("REQUEST_METHOD")="POST", false)
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova fattura</caption>
		<tr><th colspan="2">DATI DELLA FATTURA</th></tr>
		<tr>
			<td class="label">data fattura</td>
			<td class="content" colspan="5">
				<%
				CALL WriteDataPicker_Input_Manuale2("form1", "tfd_fa_data_fattura", IIF(cString(request("tfd_fa_data_fattura"))="",Date(),request("tfd_fa_data_fattura")), "", "/", false, true, LINGUA_ITALIANO, "", true, "", "window.form1.submit();") 
				%>
			</td>
			<script type="text/javascript" language="javascript">
				//function AfterDataSelected(){
				//	window.location.href = window.location.href.split('?')[0] + "?DATE=" + form1.today_date.value;
				//}
			</script>
		</tr>
		<tr>
			<td class="label">data scadenza</td>
			<td class="content" colspan="5">
				<% CALL WriteDataPicker_Input_Manuale("form1", "tfd_fa_data_scadenza", IIF(cString(request("tfd_fa_data_scadenza"))="","",request("tfd_fa_data_scadenza")), "", "/", false, true, LINGUA_ITALIANO, "", true, "") 
				%>
			</td>
		</tr>
		<tr>
			<td class="label">numero bozza</td>
			<td class="content" colspan="3">
				<%
				if cString(request("tfd_fa_data_fattura")) <> "" then
					anno = Year(request("tfd_fa_data_fattura"))
				else
					anno = Year(Now())
				end if
				sql = "SELECT MAX(fa_numero) FROM gtb_fatture WHERE ISNULL(fa_is_bozza, 0) = 1 AND fa_anno = " & anno
				new_numero = cIntero(GetValueList(conn, rs, sql)) + 1
				response.write new_numero & " del " & anno
				%>
				<input type="hidden" name="tfn_fa_numero" value="<%=new_numero%>">
				<input type="hidden" name="tfn_fa_anno" value="<%=anno%>">
			</td>
		</tr>
		<%
		' Session("FATTURE_ID_EMITTENTE_DEFAULT")
		%>
		<tr>
			<td class="label" <%=IIF(cIntero(request.form("tfn_fa_emittente_id")) > 0,"rowspan=""2""","")%>>emittente</td>
			<td class="content" colspan="3">
				<%
                sql = "SELECT * FROM tb_utentu WHERE ut_id IN (SELECT ag_id FROM gtb_agenti) "
                CALL WriteContactPicker_Input(conn, rs, sql, "", "form1", "tfn_fa_emittente_id", request("tfn_fa_emittente_id"), "LOGIN LOGINID EMAIL", false, false, false, "")
                %>
			</td>
		</tr>
		<% if cIntero(request.form("tfn_fa_emittente_id")) > 0 then %>
			<%
			sql = "SELECT * FROM gv_rivenditori WHERE riv_id = " & cIntero(request.form("tfn_fa_emittente_id"))
			rsa.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
			%>
			<tr>
				<td colspan="5">
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr><th class="L2" colspan="2">DATI ANAGRAFICA</th></tr>
						<tr>
							<td class="label">indirizzo:</td>
							<td class="content"><%=ContactAddress(rsa)%></td>
						</tr>
						<tr>
							<td class="label">codice fiscale:</td>
							<td class="content"><%=rsa("CF")%></td>
						</tr>
						<tr>
							<td class="label">p. iva:</td>
							<td class="content"><%=rsa("partita_iva")%></td>
						</tr>
						<% dim Obj_Cnt 
						set Obj_Cnt = new IndirizzarioLock 
						Obj_Cnt.LoadFromDB(rsa("IDElencoIndirizzi"))
						%>
						<tr>
							<td class="label">telefono:</td>
							<td class="content"><%=Obj_Cnt("telefono")%></td>
						</tr>
						<tr>
							<td class="label">fax:</td>
							<td class="content"><%=Obj_Cnt("fax")%></td>
						</tr>
						<tr>
							<td class="label">cellulare:</td>
							<td class="content"><%=Obj_Cnt("cellulare")%></td>
						</tr>
						<tr>
							<td class="label">e-mail:</td>
							<td class="content"><a href="mailto:<%= Obj_Cnt("email") %>"><%= Obj_Cnt("email") %></a></td>
						</tr>
					</table>
				</td>
			</tr>
			<% rsa.close %>
		<% end if %>
		<tr>
			<td class="label">intestatario</td>
			<td class="content" colspan="3">
				<%
                sql = "SELECT * FROM tb_utentu WHERE ut_id IN (SELECT riv_id FROM gtb_rivenditori) "
                CALL WriteContactPicker_Input(conn, rs, sql, "", "form1", "tfn_fa_intestatario_id", request("tfn_fa_intestatario_id"), "LOGIN LOGINID EMAIL", false, false, false, "")
                %>
			</td>
		</tr>
		<tr>
			<td class="label">esenzione iva</td>
			<td class="content" colspan="3">
				<input type="checkbox" class="checkbox" name="chk_fa_int_esente_iva">
					intestatario esente iva
			</td>
		</tr>
		<%
		'= Chk(IIF(request("ID")<>"" OR request.serverVariables("REQUEST_METHOD")="POST", rs("fa_int_esente_iva"), false))
		%>
		
		<tr><th colspan="2">NOTE</th></tr>
		<tr>
			<td class="content" colspan="2">
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td><textarea style="width:100%;" rows="5" name="tft_note"><%= request("tft_note") %></textarea></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="2">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>