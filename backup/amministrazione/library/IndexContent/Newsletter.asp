<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<% '--------------------------------------------------------
sezione_testata = "gestione delle newsletter" 

%>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<!--#INCLUDE FILE="Tools_IndexContent.asp" -->
<% 
'check dei permessi dell'utente
if NOT index.content.ChkPrm(index.content.GetID(request("co_F_table_id"), request("co_F_key_id"))) then %>
<script type="text/javascript">
	window.close()
</script>
<%
end if
'----------------------------------------------------- 


'salvataggio
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	dim lista, id, newsletter
	if cIntero(request("co_id"))>0 then
		
		index.conn.BeginTrans()
		
		newsletter = false
		sql = "DELETE FROM tb_newsletters_contents WHERE ISNULL(nlc_data_invio,0)=0 AND nlc_co_id = " & cIntero(request("co_id"))
		index.conn.execute(sql)
		
		
		if cString(request("tipi_newsletter"))<>"" then
		lista = split(request("tipi_newsletter"), ",")
			for each id in lista
				if cIntero(id) > 0 then
					newsletter = true
					sql = " INSERT INTO tb_newsletters_contents(nlc_co_id, nlc_tipo_id, nlc_insAdmin_id, nlc_insData, nlc_modAdmin_id, nlc_modData, nlc_ordine)" & _
						  " SELECT "&cIntero(request("co_id"))&","&cIntero(id)&","&Session("ID_ADMIN")&","&SQL_date(index.conn, Now())&","&Session("ID_ADMIN")&","&SQL_date(index.conn, Now())&", ISNULL(MAX(nlc_ordine),0)+1 " & _
						  " FROM tb_newsletters_contents WHERE nlc_tipo_id = "&cIntero(id)&" AND ISNULL(nlc_data_invio,0)=0"
					index.conn.execute(sql)
				end if
			next
		end if
		
		index.conn.CommitTrans()
	end if
	
	if request("salva_anteprima")<>"" then %>
		
	<% else %>
		<script language="JavaScript" type="text/javascript">
			var linkCollegamento = opener.document.getElementById("newsletter_<%= request("co_F_key_id") %>");
			if (linkCollegamento) {
				<% if newsletter then  %>
					linkCollegamento.className = linkCollegamento.className + ' newsletter';
				<% else %>
					linkCollegamento.className =  linkCollegamento.className.replace(' newsletter','');
				<% end if %>
			}
			
			<% if request("salva_chiudi")<>"" then %>
				window.close();
			</script>
			<% response.end
		else %>
			</script>
		<% end if 
	end if
end if


index.conn.BeginTrans()
CALL Index_UpdateItem(index.conn, index.GetTableName(request("co_F_table_id")), request("co_F_key_id"), true)
index.conn.CommitTrans()

'index.content.co_F_table_id = request("co_F_table_id")
'index.content.co_F_key_id = request("co_F_key_id")

'CALL index.content.Associazioni()

dim rsc, rsi, sql, value, nlc_id, main_sql, data_ultimo_invio, ordine
		
set rsc = server.createobject("adodb.recordset")
set rsi = server.CreateObject("ADODB.recordset")
%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:20px;">
		<caption>Pubblicazione nella prossima newsletter</caption>
		<%  'recupera tutti i tipi di newsletter
		sql = " SELECT * FROM tb_newsletters WHERE " & SQL_IsTrue(index.conn, "nl_gestione_dinamica_contenuti") & " ORDER BY nl_nome_it "
		rsc.open sql, index.conn, adOpenStatic, adLockOptimistic, adCmdText 
		
		if not rsc.eof then 
			main_sql = " SELECT TOP 1 nlc_data_invio, nlc_id, co_id, co_titolo_it, tab_colore, tab_titolo, nlc_ordine " & _
						" FROM (tb_contents INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id) " & _
						" 		LEFT JOIN tb_newsletters_contents ON tb_contents.co_id = tb_newsletters_contents.nlc_co_id " & _
						" WHERE co_F_key_id = "& cIntero(request("co_F_key_id")) & " AND co_F_table_id = " & cIntero(request("co_F_table_id"))
				
			rsi.open main_sql, index.conn, adOpenStatic, adLockOptimistic, adCmdText 
			%>
			<input type="hidden" name="co_id" value="<%=rsi("co_id")%>">
			<tr>
				<th colspan="4">Dati del contenuto che si vuole inserire in newsletter</th>
			</tr>
			<tr>
				<td class="label_no_width" style="width:15px;">titolo:</td>
				<td class="content" colspan="3"><%=rsi("co_titolo_it")%></td>
			</tr>
			<tr>
				<td class="label_no_width">tipo:</td>
				<td class="content" colspan="3"><%= index.content.WriteTipoRS(rsi)%></td>
			</tr>
			<tr>
				<th colspan="4">Tipologia newsletter</th>
			</tr>
		<% 	rsi.close
		end if
		while not rsc.eof
			sql = main_sql & " AND nlc_tipo_id = "&rsc("nl_id")&" AND ISNULL(nlc_data_invio,0)=0 "
			rsi.open sql, index.conn, adOpenStatic, adLockOptimistic, adCmdText
			if rsi.eof then
				nlc_id = 0
			else
				nlc_id = cIntero(rsi("nlc_id"))
			end if
			
			
			sql = main_sql & " AND nlc_tipo_id = "&rsc("nl_id")&" AND NOT ISNULL(nlc_data_invio,0)=0 ORDER BY nlc_data_invio DESC "
			data_ultimo_invio = cString(GetValueList(index.conn, NULL, sql))
			
			%>
			<tr>
				<td class="content" style="height:20px; vertical-align:middle;">
					<input type="checkbox" class="checkbox" name="tipi_newsletter" value="<%=rsc("nl_id")%>" <%=chk(not nlc_id=0)%>>
					<% if Trim(cString(rsc("nl_lingua")))<>"" then %>
						&nbsp;
						<img src="../../grafica/flag_mini_<%= rsc("nl_lingua") %>.jpg">
						&nbsp;
					<% end if %>
					<%=rsc("nl_nome_it")%>
				</td>
				<td class="note" style="width:30%;">
					<% if data_ultimo_invio <> "" then %>
						ultimo invio: <%=data_ultimo_invio%>
					<% end if %>
				</td>
				<td class="content_right" style="width:20%">
					<a class="button_L2" id="anteprima_<%=rsc("nl_id")%>" target="_blank" href="<%=GetPageSiteUrl(index.conn, rsc("nl_pagina_id"), rsc("nl_lingua"))%>&TIPO_NEWSLETTER=<%=rsc("nl_id")%>&HTML_FOR_EMAIL=1">
						VEDI ANTEPRIMA
					</a>
				</td>
			</tr>
			<% rsi.close
			rsc.movenext	
		wend 
		%>
		<tr>
			<td class="footer" colspan="4">
				<input type="submit" class="button" name="salva" value="SALVA">
				<input type="submit" class="button" name="salva_chiudi" value="SALVA & CHIUDI">
			</td>
		</tr>
	</table>
	</form>
</div>
%>
<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>