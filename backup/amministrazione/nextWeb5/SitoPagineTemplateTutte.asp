<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<% 
dim conn, sql, sqlT, rs, rst, PaginaSito, lingua, Pagine
dim Template, TemplateUnico, Esito

PaginaSito = request.Querystring("PAGINA") 

set conn = Server.CreateObject("ADODB.Connection")
set rs = server.CreateObject("ADODB.Recordset")
set rst = server.CreateObject("ADODB.Recordset")
conn.open Application("DATA_ConnectionString"),"",""

'check dei permessi dell'utente
if NOT index.content.ChkPrmF("tb_pagineSito", PaginaSito) then
	session("ERRORE") = "Non si possiedono i permessi per modificare la pagina." %>
	<script language="JavaScript">
		opener.location.reload(true);
		window.close();
	</script>
<% end if

sql = " SELECT * FROM tb_PagineSito INNER JOIN tb_webs " &_
	   " ON tb_PagineSito.id_web=tb_webs.id_webs WHERE id_paginesito=" & PaginaSito
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

if request("APPLICA")<>"" then
	'registrazione modifiche
	Select case request("selezione_template")
		case "unico"
			sql = " UPDATE tb_pages SET id_template=" & cIntero(request("sel_template_unico")) & _
				  " WHERE id_page IN (" & cIntero(rs("id_pagStage_IT")) & _
				  					  IIF(Session("LINGUA_EN"), ", " & cIntero(rs("id_pagStage_EN")), "") + _
									  IIF(Session("LINGUA_FR"), ", " & cIntero(rs("id_pagStage_FR")), "") + _
									  IIF(Session("LINGUA_DE"), ", " & cIntero(rs("id_pagStage_DE")), "") + _
									  IIF(Session("LINGUA_ES"), ", " & cIntero(rs("id_pagStage_ES")), "") + _
									  IIF(Session("LINGUA_RU"), ", " & cIntero(rs("id_pagStage_RU")), "") + _
									  IIF(Session("LINGUA_CN"), ", " & cIntero(rs("id_pagStage_CN")), "") + _
									  IIF(Session("LINGUA_PT"), ", " & cIntero(rs("id_pagStage_PT")), "") & ") " & _
				  " AND id_PaginaSito = " & PaginaSito
			CALL conn.execute(sql, , adExecuteNoRecords)
			Esito = "OK"
		case "lingue"
			for each lingua in Application("LINGUE")
				if Session("LINGUA_" & lingua) then
					sql = " UPDATE tb_pages SET id_template=" & cIntero(request("sel_template_" & lingua)) & _
						  " WHERE id_page=" & cIntero(rs("id_pagStage_" & lingua)) & _
						  " AND id_PaginaSito=" & PaginaSito
					CALL conn.execute(sql, , adExecuteNoRecords)
				end if
			next
			Esito = "OK"
	end select
end if

set Pagine = Server.CreateObject("Scripting.dictionary")
Pagine.CompareMode = vbTextCompare
for each lingua in Application("LINGUE")
	if Session("LINGUA_" & lingua) then
		sql = "SELECT * FROM tb_pages WHERE id_page=" & rs("id_pagStage_" & lingua)
		Pagine.Add lingua, conn.execute(sql)
		if lingua = LINGUA_ITALIANO then
			Template = cInteger(Pagine(lingua)("id_template").value)
			TemplateUnico = true
		elseif TemplateUnico then
			if Template <> cInteger(Pagine(lingua)("id_template").value) then
				Template = NULL
				TemplateUnico = false
			end if
		end if
	end if
next

'--------------------------------------------------------
sezione_testata = "Gestione siti - indice delle pagine - template di tutte le pagine" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

sql = "SELECT (0) AS id_page, ('Template vuoto (nessun template associato)') AS nomepage, (0) AS ordine FROM tb_pages UNION " &_
	  "SELECT id_page, nomepage, (1) AS ordine FROM tb_pages WHERE "& SQL_IsTrue(conn, "template") &" AND id_webs=" & Session("AZ_ID") &_
	  " Order by ordine, NomePage"
rst.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>
	  
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption class="border">
			Gestione template delle pagine in tutte le lingue
		</caption>
		<% if Esito<>"" then %>
			<tr>
				<td colspan="3" class="content_b">
					<br>
					Modifiche applicate correttamente.<br>
					<br>
				</td>
			</tr>
			<tr>
				<td colspan="3" class="note">
					Questa finestra si chiuder&agrave; automaticamente tra 5 secondi.
				</td>
			</tr>
			<script language="JavaScript">
				opener.location.reload(true);
				window.setTimeout("close();", 10000);
			</script>
		<% end if %>
		<tr>
			<td class="label" rowspan="<%= 2 + IIF(cInteger(Session("LINGUE_ATTIVE"))>1, 1 + Session("LINGUE_ATTIVE"), 0) %>">template:</td>
			<td class="content_center" rowspan="2">
				<input type="radio" class="noBorder" name="selezione_template" id="selezione_template_unico" onclick="Template_SetState()" value="unico" <%= chk(TemplateUnico) %>>
			</td>
			<td class="content">uguale per tutte le lingue</td>
		</tr>
		<tr>
			<td class="content">
				<% CALL DropDownRecordset(rst, "id_page", "nomepage", "sel_template_unico", IIF(TemplateUnico, Template, 0), TRUE, "", LINGUA_ITALIANO)%>
			</td>
		</tr>
		<% if cInteger(Session("LINGUE_ATTIVE"))>1 then %>
			<tr>
				<td class="content_center" rowspan="<%= 1 + cInteger(Session("LINGUE_ATTIVE")) %>">
					<input type="radio" class="noBorder" name="selezione_template" id="selezione_template_lingue" onclick="Template_SetState()" value="lingue" <%= chk(not TemplateUnico) %>>
				</td>
				<td class="content">per ogni lingua:</td>
			</tr>
			<%for each lingua in Application("LINGUE")
				if Session("LINGUA_" & lingua) then%>
				<tr>
					<td class="content">
						<img src="../grafica/flag_<%= lingua %>.jpg" alt="" border="0">
						<% CALL DropDownRecordset(rst, "id_page", "nomepage", "sel_template_" & lingua, Pagine(lingua)("id_template").value, TRUE, "", LINGUA_ITALIANO)%>
					</td>
				</tr>
				<%end if
			next
		end if %>
		
		<script language="JavaScript" type="text/javascript">
			function Template_SetState(){
				var template_unico = document.getElementById("selezione_template_unico");
				EnableIfChecked(template_unico, form1.sel_template_unico);
				
				<% if cInteger(Session("LINGUE_ATTIVE"))>1 then %>
					var template_lingue = document.getElementById("selezione_template_lingue");
					<%for each lingua in Application("LINGUE")
						if Session("LINGUA_" & lingua) then%>
							EnableIfChecked(template_lingue, form1.sel_template_<%= lingua %>);
						<%end if
					next
				end if %>
			}
			
			Template_SetState();
		</script>
		<tr>
			<td class="note" colspan="3">
				Le operazioni di copia dei contenuti dei template nelle pagine possono essere eseguite solo dalla gestione del template di ogni singola pagina.
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="3">
				<input type="submit" class="button" name="applica" value="APPLICA">
				<input type="button" class="button" name="chiudi" value="ANNULLA" onclick="window.close();">
			</td>
		</tr>
	</table>
	</form>
</div>
</body>
</html>

<% rs.close
rst.close
conn.close
set rs = nothing
set rst = nothing
set conn = nothing %>
<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>