<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<% 
dim conn, sql, rs, i, lingua

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""

set rs = Server.CreateObject("ADODB.RecordSet")
sql = "SELECT * FROM tb_pagineSito WHERE id_paginesito=" & cIntero(request.Querystring("ID"))
rs.open sql, conn, adOpenstatic, adLockOptimistic
%>

<%'--------------------------------------------------------
sezione_testata = "Gestione siti - indice delle pagine - visualizzazione" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre" border="0">
			<caption class="border">Visualizza pagine associate:</caption>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
				lingua = Application("LINGUE")(i)
				if cIntero(rs("id_pagStage_"& lingua)) > 0 AND cIntero(rs("id_pagDyn_"& lingua)) > 0 then%>
					<tr>
						<td class="label" rowspan="2" style="width:30px;"><img src="../grafica/flag_<%= lingua %>.jpg" width="26" height="15" alt="" border="0"></td>
						<td class="label" rowspan="2" style="width:50%;">Visualizza in lingua <%=  GetNomeLingua(lingua) %></td>
						<td class="content_center">
						<% 	if cIntero(rs("id_pagStage_"& lingua)) > 0 then %>
							<a HREF="dynalay.asp?PAGINA=<%=rs("id_pagStage_" & lingua)%>&lingua=<%= lingua %>" target="_blank" class="button_L2" onclick="window.close();" style="width:120px; line-height:12px;">
								PAGINA DI LAVORO
							</a>
						<% 	else %>
							<a class="button_L2_disabled" title="Pagina di lavoro non inizializzata." style="width:120px; line-height:12px;">
								PAGINA DI LAVORO
							</a>
						<% 	end if %>
						</td>
					</tr>
					<tr>
						<td class="content_center">
						<% 	if cIntero(rs("id_pagDyn_"& lingua)) > 0 then %>
							<a HREF="dynalay.asp?PAGINA=<%=rs("id_pagDyn_" & lingua)%>&lingua=<%= lingua %>" target="_blank" class="button_L2" onclick="window.close();" style="width:120px; line-height:12px;">
								PAGINA PUBBLICATA
							</a>
						<% 	else %>
							<a class="button_L2_disabled" title="Pagina pubblicata non inizializzata." style="width:120px; line-height:12px;">
								PAGINA PUBBLICATA
							</a>
						<% 	end if %>
						</td>
					</tr>
				<%end if
			next %>
			<tr>
				<td class="footer" colspan="3">
					<input type="button" class="button" name="chiudi" value="CHIUDI" onclick="window.close();">
				</td>
			</tr>
		</form>
	</table>
</div>
</body>
</html>

<% rs.close
conn.close
set rs = nothing
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>