<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_strumenti_accesso, 0))

dim conn, rs, rsp, sql ,lingua, i, UsedInStage, UsedInPublic
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")
%>

<%'--------------------------------------------------------
sezione_testata = "Elenco di pagine e template associati" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre" border="0">
			<caption class="border">
				Template che utilizzano l'immagine "<%= request("FILE") %>"
			</caption>
					<%
					sql = "SELECT DISTINCT tb_pages.id_page, tb_pages.nomepage FROM tb_layers INNER JOIN tb_pages ON tb_layers.id_pag=tb_pages.id_page "  &_
				  		  " WHERE tb_pages.id_webs=" & Session("AZ_ID") & " AND id_tipo=" & LAYER_IMAGE &_
				  	      " AND (tb_layers.nome LIKE '" & ParseSQL(request("FILE"), adChar) & "'  OR tb_layers.nome LIKE '" & right(ParseSQL(request("FILE"), adChar), len(request("FILE"))-1) & "') AND "& SQL_IsTrue(conn, "tb_pages.template")
					rs.open sql, conn, adOpenstatic, adLockReadOnly, adAsyncFetch
					if rs.eof then%>
						<tr>
							<td class="label" colspan="2">
								Nessun template trovato.
							</td>
						</tr>
					<% else 
						while not rs.eof%>
							<tr>
								<td class="content">
									<a HREF="dynalay.asp?PAGINA=<%=rs("id_page")%>&lingua=IT" target="_blank" 
										title="Visualizza la pagina in una nuova finestra">
										<%= rs("nomepage") %>
									</a>
								</td>
							</tr>
							<%rs.movenext
						wend
					end if
					rs.close %>
		</form>
	</table><br>
	
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre" border="0">
			<caption class="border">
				Pagine che utilizzano l'immagine "<%= request("FILE") %>"
			</caption>
					<%
					sql = " SELECT "
					for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
						lingua = Application("LINGUE")(i)
						sql = sql & " (D_" & lingua & ".id_template) AS template_D_" & lingua & ", " & _
									" (S_" & lingua & ".id_template) AS template_S_" & lingua & ", " & _
									" (D_" & lingua & ".id_page) AS page_D_" & lingua & ", " & _
									" (S_" & lingua & ".id_page) AS page_S_" & lingua & ", "
					next
					sql = sql & " tb_pagineSito.* FROM " & string(((ubound(Application("LINGUE")) + 1) * 2), "(") & "tb_paginesito "
					
					for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
						lingua = Application("LINGUE")(i)
						sql = sql & " LEFT JOIN tb_pages D_" & lingua & " ON tb_paginesito.id_pagDyn_" & lingua & "=D_" & lingua & ".id_page) " &_
						  			" LEFT JOIN tb_pages S_" & lingua & " ON tb_paginesito.id_pagStage_" & lingua & "=S_" & lingua & ".id_page) "
					next
					
					sql = sql & " WHERE tb_pagineSito.id_web=" & Session("AZ_ID") & " AND ("
					
					for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
						lingua = Application("LINGUE")(i)
						sql = sql & "D_" & lingua & ".id_page IN (SELECT DISTINCT id_pag FROM tb_layers WHERE id_tipo=" & LAYER_IMAGE & _
								    " AND (tb_layers.nome LIKE '" & ParseSQL(request("FILE"), adChar) & "' OR tb_layers.nome LIKE '" & right(ParseSQL(request("FILE"), adChar), len(request("FILE"))-1) & "')) OR " &_
									"S_" & lingua & ".id_page IN (SELECT DISTINCT id_pag FROM tb_layers WHERE id_tipo=" & LAYER_IMAGE & _
								    " AND (tb_layers.nome LIKE '" & ParseSQL(request("FILE"), adChar) & "' OR tb_layers.nome LIKE '" & right(ParseSQL(request("FILE"), adChar), len(request("FILE"))-1) & "')) OR "
					next
					
					sql = left(sql, len(sql)-3) & ")"
					rs.open sql, conn, adOpenstatic, adLockReadOnly, adAsyncFetch
					if rs.eof then%>
						<tr>
							<td class="label" colspan="2">
								Nessuna pagina trovata.
							</td>
						</tr>
					<% else 
						while not rs.eof%>
							<tr>
								<td class="content">
									<a href="SitoPagineMod.asp?ID=<%= rs("id_pagineSito") %>" target="modifica_pagina"
										title="Modifica la pagina in una nuova finestra">
										<%= PaginaSitoNome(rs, "") %>
									</a>
								</td>
								<td class="content">
									<table border="0" cellspacing="0" cellpadding="0">
										<% for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
											lingua = Application("LINGUE")(i) 
											UsedInStage = UsedInPage(rs("page_D_" & lingua), request("FILE"))
											UsedInPublic = UsedInPage(rs("page_D_" & lingua), request("FILE"))
											if UsedInStage OR UsedInPublic then%>
												<tr>
													<td valign="top"><img src="../grafica/flag_<%= lingua %>.jpg" width="20" alt="" border="0"></td>
													<td class="content">
														<% if UsedInPublic then%>
															<a HREF="dynalay.asp?PAGINA=<%=rs("page_D_" & lingua)%>&lingua=<%= lingua %>" target="_blank" 
																title="Visualizza la pagina in una nuova finestra">
																pagina pubblicata
															</a>
															<br>
														<%end if
														if UsedInStage then%>
															<a HREF="dynalay.asp?PAGINA=<%=rs("page_S_" & lingua)%>&lingua=<%= lingua %>" target="_blank" 
																title="Visualizza la pagina in una nuova finestra">
																pagina di lavoro
															</a>
															<br>
														<%end if %>
													</td>
												</tr>
											<% end if
										next %>
									</table>
								</td>
							</tr>
							<%rs.movenext
						wend
					end if
					rs.close %>
			<tr>
				<td class="footer" colspan="2">
					<input type="button" class="button" name="chiudi" value="CHIUDI" onclick="window.close();">
				</td>
			</tr>
		</form>
	</table>
</div>
</body>
</html>

<%
conn.close
set rs = nothing
set conn = nothing 

function UsedInPage(page, element)
	if cInteger(page)>0 then
		sql = "SELECT (COUNT(*)) AS DIPENDENZE FROM tb_layers WHERE id_tipo=" & LAYER_IMAGE &_
			  " AND tb_layers.nome LIKE '" & element & "' AND id_pag=" & page
		rsp.open sql, conn, adOpenstatic, adLockReadOnly, adAsyncFetch
		
		UsedInPage = rsp("DIPENDENZE")>0
		
		rsp.close
	else
		UsedInPage = false
	end if
end function
%>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>