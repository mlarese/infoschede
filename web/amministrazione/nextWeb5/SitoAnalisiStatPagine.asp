<%@ Language=VBScript CODEPAGE=65001%>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1000 %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="SitoAnalisiStat_TOOLS.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_strumenti_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(3)
dicitura.sottosezioni(1) = "STATISTICHE INDICE"
dicitura.links(1) = "SitoAnalisiStatIndice.asp"
dicitura.sottosezioni(2) = "STORICO PAGINE"
dicitura.links(2) = "SitoAnalisiStoricoPagine.asp"
dicitura.sottosezioni(3) = "STORICO INDICE"
dicitura.links(3) = "SitoAnalisiStoricoIndice.asp"

dicitura.sezione = "Statistiche di accesso alle pagine"

dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoAnalisi.asp"
dicitura.scrivi_con_sottosez()

dim conn, sql, rs, rsp, lingua, i

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet") 

dim totUtenti, totCrawler, totAltro, totCont
%>
<div id="content">
	<% CALL WRITE_StatisticheGenerali(conn, rs, Session("AZ_ID")) %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Statistiche di accesso alle pagine - elenco:</caption>
		<tr>
			<th class="center" width="3%" rowspan="2">ID</th>
			<th class="center" width="5%" rowspan="2">HOME</th>
			<th rowspan="2">TITOLO</th>
			<th class="center" colspan="4" style="border-bottom:0px;">NUMERO VISITE DELLA PAGINA</th>
		</tr>
		<tr>
			<th class="right" style="width:8%;">UTENTI</th>
			<th class="right" style="width:16%;">MOTORI DI RICERCA</th>
			<th class="right" style="width:8%;">ALTRI</th>
			<th class="right" style="width:8%;">TOTALE</th>
		</tr>
		<% 
		totUtenti = 0
		totCrawler = 0
		totAltro = 0
		
		sql = " SELECT tb_paginesito.*,"& _
			  " ("& SQL_if(conn, "tb_paginesito.id_paginesito=tb_webs.id_home_page", "1", "0") &") AS HOME " &_
			  " FROM (tb_PagineSito INNER JOIN tb_webs ON tb_pagineSito.id_web=tb_webs.id_webs) " &_
			  " WHERE tb_paginesito.id_web=" & Session("AZ_ID") & _
			  " ORDER BY tb_paginesito.nome_ps_IT, tb_paginesito.nome_ps_interno"
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adAsyncFetch
		
		sql = "SELECT contatore, contUtenti, contCrawler, contAltro FROM tb_pages WHERE id_page="
		while not rs.eof %>
			<tr>
				<td class="content_center" rowspan="<%= Session("LINGUE_ATTIVE") %>"><%= rs("id_pagineSito") %></td>
				<td class="content_center" rowspan="<%= Session("LINGUE_ATTIVE") %>">
					<input class="checkbox" disabled type="checkbox" name="home" value="1" <%= chk(rs("home")) %>>
				</td>
				<% for each lingua in application("LINGUE") 
					if Session("LINGUA_" & lingua) then
						rsp.open sql & cIntero(rs("id_pagDyn_"& lingua)), conn, adOpenstatic, adLockOptimistic, adCmdText
						if not rsp.eof then
							if uCase(lingua) <> "IT" then %>
								<tr>
							<% end if %>
							<td class="content">
								<table border="0" cellspacing="0" cellpadding="0" align="left">
									<tr>
										<td class="content_center"><img src="../grafica/flag_mini_<%= lingua %>.jpg" alt="" border="0"></td>
										<td class="content"><%= PaginaSitoNome(rs, lingua) %></td>
									</tr>
								</table>
							</td>
							<td class="content_right<%= IIF(cIntero(rsp("contUtenti")) = 0, " notes", "") %>"><%= cIntero(rsp("contUtenti")) %></td>
							<td class="content_right<%= IIF(cIntero(rsp("contCrawler")) = 0, " notes", "") %>"><%= cIntero(rsp("contCrawler")) %></td>
							<td class="content_right<%= IIF(cIntero(rsp("contAltro")) = 0, " notes", "") %>"><%= cIntero(rsp("contAltro")) %></td>
							<td class="content_right<%= IIF(cIntero(rsp("contatore")) = 0, " notes", "") %>"><%= cIntero(rsp("contatore")) %></td>
						</tr>
							<% totUtenti = totUtenti + cIntero(rsp("contUtenti"))
							totCrawler = totCrawler + cIntero(rsp("contCrawler"))
							totAltro = totAltro + cIntero(rsp("contAltro"))
							totCont = totCont + cIntero(rsp("contatore"))
						end if
						rsp.close
					end if
				next
			rs.moveNext
		wend %>
		<tr>
			<td class="footer" colspan="3">
				totali pagine viste:
			</td>
			<td class="footer"><%= totUtenti %></td>
			<td class="footer"><%= totCrawler %></td>
			<td class="footer"><%= totAltro %></td>
			<td class="footer"><%= totCont %></td>
		</tr>
	</table>
</div>
</html>
<%
rs.close
set rsp = nothing
conn.close
set rs = nothing
set conn = nothing
%>