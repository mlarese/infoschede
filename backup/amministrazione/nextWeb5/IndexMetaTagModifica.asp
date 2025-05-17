<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<!--#INCLUDE FILE="IndexMetaTag_TOOLS.asp" -->
<%
'check dei permessi dell'utente
if NOT index.content.ChkPrm(index.content.GetID(request("co_F_table_id"), request("co_F_key_id"))) then %>
<script type="text/javascript">
	window.close()
</script>
<%
end if

dim conn, rsi, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rsi = Server.CreateObject("ADODB.RecordSet")

if request.form("salva")<>"" OR request.form("salva_chiudi")<>"" then
	'salva le modifiche al nodo dell'indice senza aggiornare l'albero.
	
	sql = "SELECT * FROM tb_contents_index "
	CALL SalvaCampiEsterniAdvanced(conn, rsi, sql, "idx_id", request("ID"), "", "", "", "idx_")
	
	if request("salva_chiudi")<>"" then %>
		<script type="text/javascript">
			window.close();
		</script>
	<% end if
end if

'--------------------------------------------------------
sezione_testata = "Gestione meta tag - modifica"
testata_show_back = false %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim Lingua, Visual, TagTitle, MetaDescription, MetaKeywords, Title, TextValue

sql = " SELECT " & _
		"v_indice.*, " &_
		"tb_paginesito.nome_ps_IT, tb_paginesito.nome_ps_EN, tb_paginesito.nome_ps_FR, tb_paginesito.nome_ps_DE, tb_paginesito.nome_ps_ES, tb_paginesito.nome_ps_RU, tb_paginesito.nome_ps_CN, tb_paginesito.nome_ps_PT," &_
		" tb_paginesito.PAGE_keywords_IT, tb_paginesito.PAGE_keywords_EN, tb_paginesito.PAGE_keywords_FR, tb_paginesito.PAGE_keywords_DE, tb_paginesito.PAGE_keywords_ES, tb_paginesito.PAGE_keywords_RU, tb_paginesito.PAGE_keywords_CN, tb_paginesito.PAGE_keywords_PT," & _
		" tb_webs.META_keywords_IT, tb_webs.META_keywords_EN, tb_webs.META_keywords_FR, tb_webs.META_keywords_DE, tb_webs.META_keywords_ES, tb_webs.META_keywords_RU, tb_webs.META_keywords_CN, tb_webs.META_keywords_PT," & _
		" tb_paginesito.PAGE_description_IT, tb_paginesito.PAGE_description_EN, tb_paginesito.PAGE_description_FR, tb_paginesito.PAGE_description_DE, tb_paginesito.PAGE_description_ES, tb_paginesito.PAGE_description_RU, tb_paginesito.PAGE_description_CN, tb_paginesito.PAGE_description_PT," & _
		" tb_webs.META_description_IT, tb_webs.META_description_EN, tb_webs.META_description_FR, tb_webs.META_description_DE, tb_webs.META_description_ES, tb_webs.META_description_RU, tb_webs.META_description_CN, tb_webs.META_description_PT" &_
		
	" FROM ((v_indice INNER JOIN tb_siti ON v_indice.tab_sito_id = tb_siti.id_sito) " + _
	  " LEFT JOIN tb_paginesito ON v_indice.idx_link_pagina_id = tb_paginesito.id_pagineSito ) " + _
	  " LEFT JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + _
	  " WHERE idx_id=" & cIntero(request("ID"))
	  
rsi.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>
<div id="content_ridotto">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<form action="" method="post" id="form1" name="form1" >
		<caption>Gestione meta tag della voce</caption>
		<tr><th colspan="3">VOCE DELL'INDICE</th></tr>
		<tr>
			<td class="label_no_width">voce:</td>
			<td class="content" colspan="2"><%= Index.NomeCompleto(rsi("idx_id")) %>&nbsp;<%= Index.Content.WriteTipoRS(rsi) %></td>
		</tr>
		<tr><th colspan="3">META TAG DELLA VOCE</th></tr>
		<tr>
			<th colspan="3" class="L2">
				<% CALL WriteSyncroLock(rsi("tab_field_titolo_alt_" & LINGUA_ITALIANO)) %>
				TITOLO DELLA PAGINA / TESTO ALTERNATIVO DELLA VOCE
			</th>
		</tr>
		<%for each lingua in Application("LINGUE")
			TagTitle = GetTitle(cString(rsi("idx_alt_" + Lingua)), cString(rsi("co_alt_" + Lingua)), cString(rsi("nome_ps_" + Lingua)), cString(rsi("co_titolo_" + Lingua)), _
								Lingua, Visual, TextValue, Title) %>
			<tr>
				<td class="content_center" style="width:20px;" rowspan="2" title="<%= Title %>">
					<img src="<%= GetAmministrazionePath() %>grafica/flag_<%= lingua %>.jpg">
				</td>
				<td class="label_no_width" style="width:17%;" title="<%= Title %>">Attualmente visibile:</td>
				<td class="content notes" title="<%= Title %>">
					<% writeIcon(visual) %>
					&nbsp;
					<%= TagTitle %> (<%= TextValue %>)
				</td>
			</tr>
			<tr>
				<td class="label_no_width" title="<%= Title %>">Titolo della voce:</td>
				<td class="content">
					<input type="text" class="text" name="extt_idx_alt_<%= lingua %>" value="<%= rsi("idx_alt_"& lingua) %>" maxlength="255" style="width:100%;" title="<%= Title %>">
				</td>
			</tr>
		<% next %>
		<tr>
			<th colspan="3" class="L2">
				<% CALL WriteSyncroLock(rsi("tab_field_meta_keywords_it")) %>
				KEYWORDS PER I MOTORI DI RICERCA
			</th>
		</tr>
		<%for each lingua in Application("LINGUE")
			MetaKeywords = GetKeywords(cString(rsi("idx_meta_keywords_" + Lingua)), cString(rsi("co_meta_keywords_" + Lingua)), cString(rsi("PAGE_keywords_" + Lingua)), cString(rsi("META_keywords_" + Lingua)), _
									   Lingua, Visual, TextValue, Title)
			%>
			<tr>
				<td class="content_center" rowspan="2" title="<%= Title %>">
					<img src="<%= GetAmministrazionePath() %>grafica/flag_<%= lingua %>.jpg">
				</td>
				<td class="label_no_width" title="<%= Title %>">Attualmente applicate:</td>
				<td class="content note" title="<%= Title %>">
					<% writeIcon(visual) %>
					&nbsp;
					<%= MetaKeywords %> (<%= TextValue %>)
				</td>
			</tr>
			<tr>
				<td class="label_no_width" title="<%= Title %>">Keywords:</td>
				<td class="content">
					<textarea style="width:100%;" title="<%= Title %>" rows="2" name="extt_idx_meta_keywords_<%= lingua %>"><%= rsi("idx_meta_keywords_" & lingua) %></textarea>
				</td>
			</tr>
		<% next %>
		<tr>
			<th colspan="3" class="L2">
				<% CALL WriteSyncroLock(rsi("tab_field_meta_description_it")) %>
				DESCRIPTION PER I MOTORI DI RICERCA
			</th>
		</tr>
		<%for each lingua in Application("LINGUE")
			MetaDescription = GetDescription(cString(rsi("idx_meta_description_" + Lingua)), cString(rsi("co_meta_description_" + Lingua)), cString(rsi("PAGE_description_" + Lingua)), cString(rsi("META_description_" + Lingua)), _
											 Lingua, Visual, TextValue, Title)
			%>
			<tr>
				<td class="content_center" rowspan="2" title="<%= Title %>">
					<img src="<%= GetAmministrazionePath() %>grafica/flag_<%= lingua %>.jpg">
				</td>
				<td class="label_no_width" title="<%= Title %>">Attualmente applicata:</td>
				<td class="content note" title="<%= Title %>">
					<% writeIcon(visual) %>
					&nbsp;
					<%= MetaDescription %> (<%= TextValue %>)
				</td>
			</tr>
			<tr>
				<td class="label_no_width" title="<%= Title %>">Description:</td>
				<td class="content">
					<textarea style="width:100%;" title="<%= Title %>" rows="2" name="extt_idx_meta_description_<%= lingua %>"><%= rsi("idx_meta_description_" & lingua) %></textarea>
				</td>
			</tr>
		<% next %>
		<tr>
			<th colspan="3">LEGENDA DEI COLORI UTILIZZATI:</th>
		</tr>
		<% CALL WriteLegenda( "", 3 ) %>
		<tr>
			<td class="footer" colspan="4">
				<input type="submit" style="width:15%;" class="button" name="salva" value="SALVA">
				<input type="submit" style="width:15%;" class="button" name="salva_chiudi" value="SALVA & CHIUDI">
				<input type="button" style="width:15%;" class="button" name="annulla" onclick="window.close()" value="ANNULLA">
			</td>
		</tr>
		</form>
	</table>
</div>
</body>
</html>
<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>
<% 
rsi.close
conn.close
set rsi = nothing
set conn = nothing
%>