<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_IndexContent.asp" -->
<%
'check dei permessi dell'utente
if NOT index.ChkPrm(prm_indice_accesso, 0) then %>
<script type="text/javascript">
	window.close()
</script>
<%
end if

dim conn, rs, sql, ID, i, main_URL, alt_URL, non_rw_URL, confirm
ID = CIntero(request("ID"))

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
main_URL = 0
non_rw_URL = 0
alt_URL = 0


'SALVATAGGIO
if request.form("salva") <> "" AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	Dim lista, riu_id, lunghezza, stringa, idx_id, idx_id_non_rw
	'CALL ListRequest()
	'CALL ListSession()
	'response.end
	if request.form("idx_id_scelto") = "" then
		session("ERRORE") = "Scegliere ramo dell'indice."
	else
		if Len(request("riu_ids")) < 2 AND Len(request("idx_id")) < 2 AND Len(request("idx_id_non_rw")) < 2 then
			session("ERRORE") = "Selezionare almeno un URL."
		else
			session("ERRORE") = ""
			conn.beginTrans
			
			'copio gli url principali come indirizzi alternativi sulla voce dell'indice scelta
			if Len(request("idx_id")) > 1 then
				idx_id = Trim(Replace(request("idx_id"),",",""))
				sql = "SELECT * FROM tb_contents_index WHERE idx_id = " & CIntero(idx_id)
				rs.open sql, conn, AdOpenStatic, adLockReadOnly, adCmdText
				for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
					sql = " INSERT INTO rel_index_url_redirect(riu_idx_id,riu_url,riu_lingua,riu_insData,riu_insAdmin_id) " & _
						  " VALUES ("&cIntero(request("idx_id_scelto"))&",'"&rs("idx_link_url_rw_" & Application("LINGUE")(i))&"','"&CString(Application("LINGUE")(i)) & _
																										"',"&SQL_Now(conn)&","&Session("ID_ADMIN")&")"
					conn.Execute(sql) 
					main_URL = main_URL + 1
				next
				rs.close
			end if
			
			'copio gli url non rewrited come indirizzi alternativi sulla voce dell'indice scelta
			if Len(request("idx_id_non_rw")) > 1 then
				idx_id_non_rw = Trim(Replace(request("idx_id_non_rw"),",",""))
				sql = "SELECT * FROM tb_contents_index WHERE idx_id = " & CIntero(idx_id_non_rw)
				rs.open sql, conn, AdOpenStatic, adLockReadOnly, adCmdText
				for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
					sql = " INSERT INTO rel_index_url_redirect(riu_idx_id,riu_url,riu_lingua,riu_insData,riu_insAdmin_id) " & _
						  " VALUES ("&cIntero(request("idx_id_scelto"))&",'" & Replace(rs("idx_link_url_" & Application("LINGUE")(i)),"?","default.aspx?") & "','" & _
															CString(Application("LINGUE")(i)) & "',"&SQL_Now(conn)&","&Session("ID_ADMIN")&")"
					conn.Execute(sql) 
					non_rw_URL = non_rw_URL + 1
				next
				rs.close
			end if
			
			'copio gli url alternativi
			stringa = ParseSQL(request("riu_ids"), adChar)
			lunghezza = Len(stringa)
			lista = Split(Right(stringa,lunghezza-1), ",")	
			for each riu_id in lista
				if Trim(riu_id)<>"" then
					sql = "UPDATE rel_index_url_redirect SET riu_idx_id="&cIntero(request("idx_id_scelto"))&", riu_modData="&SQL_Now(conn)& _
																	", riu_modAdmin_id="&Session("ID_ADMIN")&" WHERE riu_id = " & riu_id
					conn.Execute(sql)
					alt_URL = alt_URL + 1
				end if
			next
			
			conn.commitTrans
		end if
	end if
end if

if session("ERRORE") = "" then
	confirm = true
end if

'--------------------------------------------------------
sezione_testata = "copia indirizzi alternativi" %>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 


sql = " SELECT * FROM tb_contents_index WHERE idx_id = "& CIntero(request("IDX")) 
rs.Open sql, conn, AdOpenStatic, adLockReadOnly, adCmdText


'PAGINA DI CONFERMA DELL'OPERAZIONE DOPO IL SALVATAGGIO
if request.form("salva") <> "" AND Request.ServerVariables("REQUEST_METHOD")="POST" AND confirm then
	%>
	<div id="content_ridotto">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption class="border">Elenco indirizzi della voce "<%= GetValueList(conn, NULL, "SELECT co_titolo_it FROM v_indice WHERE idx_id = "& cIntero(request("IDX"))) %>"</caption>
			<tr>
				<td class="label" colspan="3">
					L'operazione di salvataggio degli URL è stata completata. Sono stati copiati <%=main_URL%> URL principali, <%=non_rw_URL%> URL non rewrited e spostati <%=alt_URL%> URL alternativi.
				</td>
			</tr>
			<tr>
				<td class="footer">
					<input style="width:70px;" type="button" class="button" name="chiudi" value="CHIUDI" onclick="window.close();">
				</td>
			</tr>
		</table>
	</div>
	</body>
	</html>
	<script language="JavaScript" type="text/javascript">
		FitWindowSize(this);
	</script>
	<%
	response.end
end if


%>
<script language="JavaScript" type="text/javascript">

	function seleziona(name,id){
		if (document.getElementById("chk_" + id).checked){
			document.getElementById(name).value = document.getElementById(name).value + id + ",";
		}
		else{
			document.getElementById(name).value = document.getElementById(name).value.replace("," + id + ",", ",");
		}		
	}
	
	function Tutti() {
	for(var i=0; i < form1.elements.length; i++)
		if (!form1.elements[i].checked && form1.elements[i].name.substring(0, 4) == "chk_"){
			form1.elements[i].click();
		}
	}
	
</script>
<div id="content_ridotto">
<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption class="border">Elenco indirizzi della voce "<%= GetValueList(conn, NULL, "SELECT co_titolo_it FROM v_indice WHERE idx_id = "& cIntero(request("IDX"))) %>"</caption>
		<tr>
			<td class="content" colspan="3">
				<b>Seleziona gli URL che desideri copiare.</b>
			</td>
		</tr>
		<tr>
			<td class="label" colspan="3">
				Gli URL principali e i non rewrited verranno <u>copiati</u> come indirizzi alternativi mentre gli URL alternativi verranno <u>spostati</u>.
				<br>&nbsp;
			</td>
		</tr>
		<tr>
			<th style="width:5%;">&nbsp;</th>
			<th>URL principali</th>
			<th style="text-align:center; width:10%;">LINGUA</th>
		</tr>
		<%	if not rs.eof then %>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
				<% if i = 0 then %>
					<td class="content" style="vertical-align:middle;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">
						<input type="checkbox" class="checkbox" id="chk_<%=rs("idx_id")%>" name="chk_<%=rs("idx_id")%>" value="" onclick="seleziona('idx_id', <%=rs("idx_id")%>)">
					</td>
				<% end if %>	
					<td class="content"><%= rs("idx_link_url_rw_" & Application("LINGUE")(i)) %></td>
					<td class="content_center">
						<img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					</td>
				</tr>
			<% next %>
			
			<tr>
				<th style="width:5%;">&nbsp;</th>
				<th>URL non rewrited</th>
				<th style="text-align:center; width:10%;">LINGUA</th>
			</tr>
		
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
				<% if i = 0 then %>
					<td class="content" style="vertical-align:middle;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">
						<input type="checkbox" class="checkbox" id="chk_<%=rs("idx_id")%>" name="chk_<%=rs("idx_id")%>" value="" onclick="seleziona('idx_id_non_rw', <%=rs("idx_id")%>)">
					</td>
				<% end if %>	
					<td class="content"><%= Replace(rs("idx_link_url_" & Application("LINGUE")(i)),"?","default.aspx?") %></td>
					<td class="content_center">
						<img src="<%= GetAmministrazionePath() %>grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					</td>
				</tr>
			<% next %>
		<% end if %>
		
		<%
		rs.close
		sql = " SELECT * FROM rel_index_url_redirect WHERE riu_idx_id = "& CIntero(request("IDX")) 
		rs.Open sql, conn, AdOpenStatic, adLockReadOnly, adCmdText
		%>
		<% if not rs.eof then %>
			<tr>
				<th style="width:5%;">&nbsp;</th>
				<th>URL alternativi</th>
				<th style="text-align:center; width:10%;">LINGUA</th>
			</tr>
		<% end if %>
		<%	while not rs.eof %>
			<tr>
				<td class="content">
					<input type="checkbox" class="checkbox" id="chk_<%=rs("riu_id")%>" name="chk_<%=rs("riu_id")%>" value="" onclick="seleziona('riu_ids', <%=rs("riu_id")%>)" >
				</td>
				<td class="content"><%= rs("riu_url") %></td>
				<td class="content_center">
					<img src="<%= GetAmministrazionePath() %>grafica/flag_<%= rs("riu_lingua") %>.jpg" width="26" height="15" alt="" border="0">
				</td>
			</tr>
			<% rs.moveNext %>
		<% wend %>
		<input type="hidden" id="idx_id" name="idx_id" value=",">
		<input type="hidden" id="idx_id_non_rw" name="idx_id_non_rw" value=",">
		<input type="hidden" id="riu_ids" name="riu_ids" value=",">
		<tr>
			<th style="width:5%;">&nbsp;</th>
			<th>VOCE DI DESTINAZIONE</th>
			<th style="text-align:center; width:10%;">&nbsp;</th>
		</tr>
		<tr>
			<td colspan="3" class="content">
				<%
				CALL index.WritePicker("", "", "form1", "idx_id_scelto", "", Application("AZ_ID"), true, false, 75, false, true)
				%>
			</td>
		</tr>
		<tr>
			<td colspan="3" class="label">Scegli la voce dell'indice nella quale copiare gli URL selezionati.<br>&nbsp;</td>
		</tr>
		<tr>
			<td colspan="3" class="footer">(*) Campi obbligatori.
				<input style="width:70px;" type="submit" class="button" name="salva" value="SALVA">
				<input style="width:70px;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
			</td>
		</tr>
	</table>
</form>
</div>
</body>
</html>
<%
rs.close
conn.close
set rs = nothing
set conn = nothing %>
<script language="JavaScript" type="text/javascript">
	Tutti();
	FitWindowSize(this);
</script>