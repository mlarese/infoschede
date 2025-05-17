<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
dim nextWeb_Conn, rs, sql, lingua, value, min_ord, max_ord, ord, sub_id
set nextWeb_Conn = Server.CreateObject("ADODB.Connection")
nextWeb_Conn.open Application("l_conn_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")


if Request.ServerVariables("REQUEST_METHOD")="POST" then
	if cIntero(request("id_new_ord"))>0 then
		nextWeb_Conn.BeginTrans()
	
		sql = " UPDATE tb_newsletters_contents " & _
			  " SET nlc_ordine = "&cIntero(request("new_ord"))&" WHERE nlc_id = " & cIntero(request("id_new_ord"))
		nextWeb_Conn.execute(sql)
		
		sql = " UPDATE tb_newsletters_contents " & _
			  " SET nlc_ordine = "&cIntero(request("old_ord"))&" WHERE nlc_id = " & cIntero(request("id_sub_ord"))
		nextWeb_Conn.execute(sql)
		
		nextWeb_Conn.CommitTrans()
	end if

	
	if request("idx_padre_id")<>"" then
		'inserimento
		dim co_id
		sql = "SELECT idx_content_id FROM tb_contents_index WHERE idx_id = " & cIntero(request("idx_padre_id"))
		co_id = cIntero(GetValueList(nextWeb_Conn, NULL, sql))
		
		nextWeb_Conn.BeginTrans()
		
		sql = "DELETE FROM tb_newsletters_contents WHERE ISNULL(nlc_data_invio,0)=0 AND nlc_co_id = " & co_id & " AND nlc_tipo_id = " & cIntero(request("TIPO_NEWSLETTER"))
		nextWeb_Conn.execute(sql)
		
		sql = " INSERT INTO tb_newsletters_contents(nlc_co_id, nlc_tipo_id, nlc_insAdmin_id, nlc_insData, nlc_modAdmin_id, nlc_modData, nlc_ordine)" & _
			  " SELECT "&co_id&","&cIntero(request("TIPO_NEWSLETTER"))&","&Session("ID_ADMIN")&","&SQL_date(nextWeb_Conn, Now())&","&Session("ID_ADMIN")&","&SQL_date(nextWeb_Conn, Now())&", ISNULL(MAX(nlc_ordine),0)+1 " & _
			  " FROM tb_newsletters_contents WHERE nlc_tipo_id = "&cIntero(request("TIPO_NEWSLETTER"))&" AND ISNULL(nlc_data_invio,0)=0"
		nextWeb_Conn.execute(sql)
		
		nextWeb_Conn.CommitTrans()
	else
		'rimozione
		dim nlc_id, var
		nlc_id = cIntero(request("id_elimina"))
		if nlc_id > 0 then
			sql = "DELETE FROM tb_newsletters_contents WHERE nlc_id = " & nlc_id
			nextWeb_Conn.execute(sql)
		end if
	end if
end if

'--------------------------------------------------------
sezione_testata = "Gestione contenuti presenti nella prossima newsletter"
body_attributes = " onunload=""ricaricaFrameset()"" " %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<%'----------------------------------------------------- 

sql = " SELECT nlc_data_invio, nlc_id, co_id, co_titolo_it, tab_colore, tab_titolo, nlc_ordine " & _
	  " FROM (tb_contents INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id) " & _
	  " INNER JOIN tb_newsletters_contents ON tb_contents.co_id = tb_newsletters_contents.nlc_co_id " & _
	  " WHERE ISNULL(nlc_data_invio,0)=0 AND nlc_tipo_id = " & cIntero(request("TIPO_NEWSLETTER")) & _
	  " ORDER BY nlc_ordine, co_titolo_it "

rs.open sql, nextWeb_Conn, adOpenStatic, adLockOptimistic
%>

<script language="JavaScript">

	var ricarica = true;

	function cancella(id,titolo){
		ricarica = false;
		var answer = confirm ("Vuoi rimuovere il contenuto '"+titolo+"' dalla newsletter?");
		if (answer){
			form1.id_elimina.value = id;
			form1.id_elimina.onchange();
		}
	}
	
	function riordina(id, new_ord, sub_id, old_ord){
		ricarica = false;
		form1.id_new_ord.value = id;
		form1.new_ord.value = new_ord;
		
		form1.id_sub_ord.value = sub_id;
		form1.old_ord.value = old_ord;
		
		form1.id_new_ord.onchange();
	}
	
	function ricaricaFrameset(){
		if (ricarica){
			opener.SetPreview(<%=request("pagina")%>);
		}
		ricarica = true;
	}
</script>

<div id="content_ridotto">
<form action="" method="post" name="form1" id="form1">
	<input type="hidden" name="idx_padre_id" id="idx_padre_id" value="">
	<input type="hidden" name="view_idx_padre_id" id="view_idx_padre_id" value="">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table border="0" cellspacing="0" cellpadding="1" align="right">
				<tr>
					<td style="font-size: 1px; padding-right:1px;" nowrap>
						<input type="button" class="button" name="nuovo" id="nuovo" value="AGGIUNGI CONTENUTO" 
							onclick="OpenAutoPositionedScrollWindow('<%= GetAmministrazionePath() %>library/IndexContent/IndexSeleziona.asp?formname=form1&inputname=idx_padre_id&<%= IIF(true, "SoloFoglie=1&", "") %><%= IIF(cIntero(0)>0, "WebIdFilter=" & "0" & "&", "") %>selected=' + form1.idx_padre_id.value, 'selezione_voce', 760, 450, true)">
					</td>
				</tr>
			</table>
			<%= sezione_testata %>
		</caption>
		<tr>
			<th colspan="4">CONTENUTI PRESENTI NELLA NEWSLETTER</th>
		</tr>
		
		<input type="hidden" id="id_elimina" name="id_elimina" value="0" onchange="form1.submit();">
		<input type="hidden" id="id_new_ord" name="id_new_ord" value="0" onchange="form1.submit();">
		<input type="hidden" id="new_ord" name="new_ord" value="0" onchange="">
		<input type="hidden" id="id_sub_ord" name="id_sub_ord" value="0" onchange="">
		<input type="hidden" id="old_ord" name="old_ord" value="0" onchange="">
		<%
		if not rs.eof then 
			min_ord = rs("nlc_ordine")
			rs.moveLast
			max_ord = rs("nlc_ordine")
			rs.moveFirst
		end if
		%>

		<% while not rs.eof %>
			<tr>
				<td class="content" style="height:18px;">
					<%=index.content.WriteNomeETipo(rs)%>
				</td>
				<td class="content_right" style="width:5%;">
					<% if cIntero(rs("nlc_ordine"))>min_ord then %>
						<% rs.movePrevious
						ord = cIntero(rs("nlc_ordine"))
						sub_id = rs("nlc_id")
						rs.moveNext
						%>
						<a id="cancella" class="button_L2 freccia_up" style="width:16px; height:16px;" href="javascript:void(0);" onclick="riordina('<%=rs("nlc_id")%>','<%=ord%>','<%=sub_id%>','<%=rs("nlc_ordine")%>')" title="Sposta in alto nella lista" <%= ACTIVE_STATUS %>>
							&nbsp;
						</a>
					<% end if %>
					&nbsp;			
				</td>
				<td class="content_right" style="width:5%;">
					<% if cIntero(rs("nlc_ordine"))<max_ord then %>
						<% rs.moveNext
						ord = cIntero(rs("nlc_ordine"))
						sub_id = rs("nlc_id")
						rs.movePrevious
						%>
						<a id="cancella" class="button_L2 freccia_down" style="width:16px; height:16px;" href="javascript:void(0);" onclick="riordina('<%=rs("nlc_id")%>','<%=ord%>','<%=sub_id%>','<%=rs("nlc_ordine")%>')" title="Sposta in basso nella lista" <%= ACTIVE_STATUS %>>
							&nbsp;
						</a>
					<% else %>
						&nbsp;
					<% end if %>
				</td>
				<td class="content_right" style="width:27%;">												
					<a id="cancella" class="button_L2" href="javascript:void(0);" onclick="cancella('<%=rs("nlc_id")%>','<%=rs("co_titolo_it")%>')" title="Apre la visualizzazione ad albero." <%= ACTIVE_STATUS %>>
						RIMUOVI DALLA NEWSLETTER
					</a>
				</td>
			</tr>
			<% rs.moveNext%>
		<% wend 
		%>
		
		<tr>
			<td class="footer" colspan="4">
				<input type="button" class="button" name="chiudi" value="CHIUDI" onclick="window.close();">
			</td>
		</tr>
	</table>
</form>
</div>
</body>
</html>

<% rs.close
nextWeb_Conn.close
set rs = nothing
set nextWeb_Conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>