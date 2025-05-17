<% 
'.................................................................................................
'.................................................................................................
'FUNZIONI E PROCEDURE
'.................................................................................................
'.................................................................................................


'.................................................................................................
'..		gestione dell'input per il collegamenti dei documenti alle pratiche
'..		id_list				elenco degli id del documento
'..		name_list			elenco dei nomi del documento
'.................................................................................................
sub GestioneDocumentiCollegati(conn, rs, att_id)
	dim IdList, NameList
	att_id = cInteger(att_id) 
	if att_id > 0 then
		sql = "SELECT * FROM tb_documenti INNER JOIN tb_allegati "& _
				 " ON tb_documenti.doc_id = tb_allegati.all_documento_id " & _
				 " WHERE "& AL_query(conn, AL_DOCUMENTI) & " AND all_attivita_id="& cIntero(att_id)
		rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		
		while not rs.eof
			IdList = IdList & rs("doc_id") &";"
			NameList = NameList & rs("doc_nome") &";"
			rs.movenext
		wend
		rs.close
	else
		IdList = request("documenti")
		NameList = request("visDoc")
	end if
%>
	<script language="JavaScript" type="text/javascript">
		function NuovoDocumento(){
			var pratica_id 
			if (form1.tfn_att_pratica_id)
				pratica_id = form1.tfn_att_pratica_id.value;
			else
				pratica_id = ""
			OpenAutoPositionedScrollWindow('DocumentoNew.asp?ATT=si&tfn_doc_pratica_id=' + pratica_id, 'nuovoDoc', 760, 500, true)
		}
	</script>
	<tr>
		<th colspan="4">DOCUMENTI ALLEGATI</th>
	</tr>
	<tr>
		<td class="content" colspan="4">
			<input type="hidden" name="documenti" value="<%= IdList %>">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td width="90%">
						<textarea READONLY style="width:100%; height:69px;" name="visDoc"><%= NameList %></textarea>
					</td>
					<td>
						<a class="button_textarea" href="javascript:void(0)" onclick="NuovoDocumento()">
							NUOVO
						</a>
						<a class="button_textarea" 
							href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('Attivita_allega.asp?page_No=1&elenco='+form1.documenti.value, 'allegati', 700, 550, true)">
							SCEGLI
						</a>
						<a class="button_textarea" 
							href="javascript:void(0)" onclick="form1.visDoc.value='';form1.documenti.value=''">
							RESET
						</a>
					</td>
				</tr>
			</table>
		</td>
	</tr>
<% end sub

%>