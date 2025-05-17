<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
dim conn, sql, rs, rsp, i, lingua, title
 
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rsp = Server.CreateObject("ADODB.RecordSet")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT * FROM tb_pages WHERE id_page="& cIntero(request("ID"))
rs.open sql, conn, adOpenstatic, adLockReadOnly, adAsyncFetch
%>

<%'--------------------------------------------------------
sezione_testata = "Gestione siti - templates - pagine associate" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>

<div id="content_ridotto">
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>Elenco pagine associate al template '<%= rs("nomepage") %>'</caption>
					<tr>
						<th>TITOLO</th>
						<th class="center" style="width:4%;">LINGUA</th>
						<th class="center" style="width:35%;" colspan="2">VERSIONI ASSOCIATE</th>
					</tr>
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
					sql = sql & " WHERE ("
					
					for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
						lingua = Application("LINGUE")(i)
						sql = sql & " D_" & lingua & ".id_template=" & rs("id_page") & " OR S_" & lingua & ".id_template=" & rs("id_page") & " OR "
					next
					sql = left(sql, len(sql)-3) & ")"
					rsp.open sql, conn, adOpenstatic, adLockReadOnly, adAsyncFetch %>
					<tr>
						<% if rsp.recordcount<1 then %>
							<td class="label" colspan="4">
								Nessuna pagina associata a questo template
							</td>
						<% else 
							dim row
							while not rsp.eof
								row = false
								for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
									lingua = Application("LINGUE")(i) 
									if rsp("template_D_" & lingua) = rs("id_page") OR rsp("template_S_" & lingua) = rs("id_page") then %>
										<tr>
											<%if not row then%>
												<td class="content">
													<a href="SitoPagineMod.asp?ID=<%= rsp("id_pagineSito") %>" target="modifica_pagina"
														title="Modifica la pagina in una nuova finestra">
                                                        <%= PaginaSitoNome(rsp, "")%>
													</a>
													<% if cInteger(rsp("id_web")) <> cInteger(Session("AZ_ID")) then %>
														<br>ERRORE NELL'ASSOCIAZIONE DEL TEMPLATE
													<% end if %>
												</td>
												<% row = true
											else %>
												<td class="content">&nbsp;</td>
											<% end if %>
											<td class="content_center"><img src="../grafica/flag_mini_<%= lingua %>.jpg" alt="" border="0"></td>
											<td class="content">
												<% if rsp("template_S_" & lingua) = rs("id_page") then%>
													<a HREF="dynalay.asp?PAGINA=<%=rsp("page_S_" & lingua)%>&lingua=<%= lingua %>" target="_blank" 
														title="Visualizza la pagina in una nuova finestra" <%=ACTIVE_STATUS%>>
														pagina di lavoro
													</a>
												<% else %>	
													&nbsp;
												<%end if %>
											</td>
											<td class="content">
												<% if rsp("template_D_" & lingua) = rs("id_page") then%>
													<a HREF="dynalay.asp?PAGINA=<%=rsp("page_D_" & lingua)%>&lingua=<%= lingua %>" target="_blank" 
														title="Visualizza la pagina in una nuova finestra" <%=ACTIVE_STATUS%>>
														pagina pubblicata
													</a>
												<% else %>
													&nbsp;
												<%end if %>
											</td>
										</tr>
									<% end if
								next 
								rsp.movenext
								if not rsp.eof then%>
									</tr>
									<tr>
								<%end if
							wend
						end if %>
					</tr>	
					<% rsp.close %>
				<tr>
					<td class="footer" colspan="4">
						<a class="button" href="javascript:void(0);" onclick="window.close();">
							CHIUDI
						</a>
					</td>
				</tr>
			</table>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set conn = nothing
%>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>
