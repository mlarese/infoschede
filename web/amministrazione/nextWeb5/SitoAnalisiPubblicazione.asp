<%@ Language=VBScript CODEPAGE=65001%>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1000 %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_strumenti_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Analisi stato delle pagine"
	dicitura.puls_new = "INDIETRO"
if request("FROM")<>"" then
	dicitura.link_new = "SitoPagine.asp"
else
	dicitura.link_new = "SitoAnalisi.asp"
end if
dicitura.scrivi_con_sottosez()

dim conn, sql, rs, rsp, i, lingua
 
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rsp = Server.CreateObject("ADODB.RecordSet")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = " SELECT tb_paginesito.*, ("& SQL_If(conn, "tb_paginesito.id_paginesito=tb_webs.id_home_page", "1", "0") &") AS HOME " &_
	  " FROM (tb_PagineSito INNER JOIN tb_webs ON tb_pagineSito.id_web=tb_webs.id_webs) " &_
	  " WHERE tb_paginesito.id_web=" & Session("AZ_ID") & _
	  " ORDER BY tb_paginesito.nome_ps_IT, tb_paginesito.nome_ps_interno "
rs.open sql, conn, adOpenStatic, adLockOptimistic, adAsyncFetch
%>

<script language="JavaScript" type="text/javascript">
	//pubblicazione
	function azione_pagina(action, page_source, page_dest, lingua, nome_lingua){
		OpenAutoPositionedWindow('SitoPagineCopia.asp?ID_S=' + page_source + '&ID_D=' + page_dest + '&lingua=' + lingua + '&nome_lingua=' + nome_lingua + '&azione=' + action, 
				 				 'action', 500, 200);
    }
	
</script>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption class="border">Analisi delle pagine.</caption>
		<tr>
			<td class="label_no_width">Pubblicazione di <span style="text-transform:uppercase;">TUTTE</span> le pagine del sito.</td>
			<td class="content_center" style="width:25%;">
				<a HREF="javascript:void(0)" class="button_L2_block" onclick="OpenAutoPositionedWindow('SitoPaginePubblicaTutte.asp', 'pubblica', 500, 250)"
				   title="Permette la pubblicazione di tutte le pagine del sito." <%= ACTIVE_STATUS %>>
					PUBBLICA TUTTE LE PAGINE
				</a>
			</td>
		</tr>
	</table>
	
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:30px;">
		<caption>Analisi di n&deg; <%= rs.recordcount %> pagine.</caption>
		<tr>
			<th class="center" width="3%">ID</th>
			<th>TITOLO</th>
			<th>TEMPLATE</th>
			<th class="center" style="width:15%;">STATO</th>
			<th class="center" style="width:10%;">OPERAZIONI</th>
		</tr>
				<% while not rs.eof 
					CALL Ceck_page_exists(conn, rs)%>
					<tr>
						<td class="content_center" rowspan="<%= Session("LINGUE_ATTIVE") %>">
							<%= rs("id_paginesito") %>
						</td>
						<% for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
							lingua = Application("LINGUE")(i)
							if Session("LINGUA_" & lingua) then
								if lingua<>LINGUA_ITALIANO then%>
								<tr>
								<%end if%>
									<td class="content">
										<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
											<tr>
												<td width="14%" align="center"><img src="../grafica/flag_mini_<%= lingua %>.jpg" alt="" border="0"></td>
												<td class="content">    
                                                    <%= PaginaSitoNome(rs, lingua) %>
												</td>
											</tr>
										</table>
									</td>
									<td class="content">
										<% CALL get_Template_Name(conn, rsp, rs("id_pagStage_" & lingua)) %>
									</td>
									<% if cInteger(rs("id_pagDyn_" & lingua))<1 OR _
										  must_be_published(conn, rsp, cInteger(rs("id_pagStage_" & lingua)), cInteger(rs("id_pagDyn_" & lingua))) then %>
										<td class="content_center dapubblicare"><strong>da pubblicare</strong></td>
										<td class="content_center">
											<a HREF="javascript:void(0);" class="button_L2_block"" onclick="azione_pagina('PUBBLICA', <%= rs("id_pagStage_" & lingua) %>, <%= cInteger(rs("id_pagDyn_" & lingua)) %>, '<%= lingua %>', '<%= GetNomeLingua(lingua) %>')">
												PUBBLICA
											</a>
										</td>
									<%else 
										'pagina OK%>
										<td class="content" align="center">
											pubblicata
										</td>
										<td class="content">&nbsp;</td>
									<% end if %>
								</tr>
							<%end if
						next%>
						
					<% rs.MoveNext
				wend %>
	</table>
</div>
</html>
<%
rs.close
conn.close
set rs = nothing
set rsp = nothing
set conn = nothing



sub get_Template_Name(conn, rs, page_id)
	sql = "SELECT tb_templates.nomepage, tb_pages.id_template FROM tb_pages INNER JOIN tb_pages tb_templates ON tb_pages.id_template=tb_templates.id_page " &_
		  " WHERE tb_pages.id_page=" & page_id
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
	if not rs.eof then %>
		<a href="SitoTemplateMod.asp?ID=<%= rs("id_template") %>" target="_blank" title="apre la pagina di modifica del template" <%= ACTIVE_STATUS %>>
			<%= rs("nomepage") %>
		</a>
	<%else %>
		- - - -
	<%end if
	
	rs.close
end sub
%>