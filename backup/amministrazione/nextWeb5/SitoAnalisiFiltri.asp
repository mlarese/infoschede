<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->

<%
dim conn, rs, sql, ID, var
ID = CIntero(request("ID"))

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")


'SALVA
if request.form("salva") <> "" then
	if request.form("extt_fil_valore") = "" then
		session("ERRORE") = "Valore obbligatorio."
	else
		sql = "SELECT * FROM tb_contents_log_filtri"
		CALL SalvaCampiEsterni(conn, rs, sql, "fil_id", ID, "", "")
		
		if session("ERRORE") = "" then
			response.redirect "SitoAnalisiFiltri.asp"
		end if
	end if
end if


sql = " SELECT * FROM tb_contents_log_filtri ORDER BY fil_parametro"
rs.Open sql, conn, AdOpenStatic, adLockReadOnly, adAsyncFetch

dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Filtri di esclusione log - elenco"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoAnalisi.asp"
dicitura.scrivi_con_sottosez()
%>
<div id="content">
<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Gestione filtri di esclusione dei conteggi sui log</caption>
		<tr>
			<th style="width:24%;">PARAMETRO DA FILTRARE</th>
			<th style="width:12%;">CONFRONTO</th>
            <th>VALORE DA ESCLUDERE</th>
			<th class="center" style="width:20%;">OPERAZIONI</th>
		</tr>
		<%	while not rs.eof
				if ID = rs("fil_id") then 			'modifica %>
		<tr>
			<td class="content"><% SelectParameters(rs("fil_parametro")) %></td>
			<td class="content"><% SelectFiltro(rs("fil_tipo")) %></td>
			<td class="content">
				<input type="text" class="text" name="extt_fil_valore" value="<%= rs("fil_valore") %>" maxlength="255" style="width:95%;">&nbsp;(*)
			</td>
			<td class="Content_center" style="vertical-align: middle;">
				<input style="width:70px;" type="submit" class="button" name="salva" value="SALVA">
				<input style="width:70px;" type="button" class="button" name="annulla" value="ANNULLA" onclick="document.location='<%= GetPageName() %>';">
			</td>
		</tr>
		<% 		else 								'visualizza %>
		<tr>
			<td class="content"><%= rs("fil_parametro") %></td>
			<td class="content"><%= FiltroNome(rs("fil_tipo")) %></td>
            <td class="content"><%= rs("fil_valore") %></td>
			<td class="Content_center" style="font-size:1px;">
				<a class="button" href="?ID=<%= rs("fil_id") %>">
					MODIFICA
				</a>
				&nbsp;
				<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('FILTRI','<%= rs("fil_id") %>');" >
				    CANCELLA
				</a>
			</td>
		</tr>
		<% 		end if
				rs.moveNext
			wend %>
		<%	if ID = 0 then								'nuovo %>
		<tr>
			<td class="content"><% SelectParameters(request("extt_fil_parametro")) %></td>
			<td class="content"><% SelectFiltro(CIntero(request("extn_fil_tipo"))) %></td>
			<td class="content">
				<input type="text" class="text" name="extt_fil_valore" value="<%= request("extt_fil_valore") %>" maxlength="255" style="width:95%;">&nbsp;(*)
			</td>
			<td class="Content_center">
				<input style="width:70px;" type="submit" class="button" name="salva" value="AGGIUNGI">
				<input style="width:70px;" type="button" class="button" name="annulla" value="ANNULLA" onclick="document.form1.reset();">
			</td>
		</tr>
		<% 	end if %>
		<tr><td colspan="5" class="footer">(*) Campi obbligatori.</td></tr>
	</table>
</form>
</div>
</body>
</html>
<%
rs.close
conn.close
set rs = nothing
set conn = nothing


Sub SelectParameters(selected) %>
	<select name="extt_fil_parametro" style="width:100%;">
<%	for each var in Request.ServerVariables %>
		<option<%= IIF(var = selected, " selected", "") %>><%= var %></option>
<%	next %>
	</select>
<%
End Sub

Sub SelectFiltro(selected) %>
	<select name="extn_fil_tipo">
		<option value="1"<%= IIF(selected = FILTRO_TEXT_UGUALE, " selected", "") %>>UGUALE</option>
		<option value="2"<%= IIF(selected = FILTRO_TEXT_FULLTEXT, " selected", "") %>>FULL-TEXT</option>
		<option value="3"<%= IIF(selected = FILTRO_TEXT_INIZIO, " selected", "") %>>INIZIA CON</option>
		<option value="4"<%= IIF(selected = FILTRO_TEXT_FINE, " selected", "") %>>FINISCE CON</option>
	</select>
<%
End Sub

Function FiltroNome(selected)
	select case selected
		case 1
			FiltroNome = "UGUALE"
		case 2
			FiltroNome = "FULL-TEXT"
		case 3
			FiltroNome = "INIZIA CON"
		case 4
			FiltroNome = "FINISCE CON"
	end select
End Function
%>
