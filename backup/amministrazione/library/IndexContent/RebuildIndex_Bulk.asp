<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<% Server.ScriptTimeout = 2147483647 %>
<% 

Session("LOGIN_4_LOG") = "true"
Session("UTENTE_MANUTENZIONE") = "true"


sezione_testata = "Rebuilding Index" 
%>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<!--#INCLUDE FILE="Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../database/Tools4DataBase.asp" -->
<% 

'*****************************************************************************************************************
'verifica dei permessi
CALL VerificaPermessiUtente(true)
'*****************************************************************************************************************

dim conn, rs, sql, ricarica, totale_conta, totale_manca
dim nextaim_admin_id, dummy_admin_id, idx_id
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
ricarica = false

'admin a cui viene settata la modifica
nextaim_admin_id = GetAdminId(conn, "NEXTAIM")
if nextaim_admin_id = 0 then
	nextaim_admin_id = GetAdminId(conn, "COMBINARIO")
	if nextaim_admin_id = 0 then
		nextaim_admin_id = 52
	end if
end if

'admin a cui settare i record da ricostruire
dummy_admin_id = 1
%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>REBUILDING INDEX</caption>
		<tr>
			<th class="L2" colspan="2"> Esecuzione dell'aggiornamento completo dell'indice in modalità sequenziale.</th>
		</tr>
	</table>
		<%	
		
		sql = "SELECT COUNT(*) FROM tb_contents_index WHERE idx_modAdmin_id = " & nextaim_admin_id
		totale_conta = cIntero(GetValueList(conn, rs, sql))
		
		sql = "SELECT COUNT(*) FROM tb_contents_index WHERE idx_modAdmin_id = " & dummy_admin_id
		totale_manca = cIntero(GetValueList(conn, rs, sql))
		
		sql = " SELECT top 5 idx_id, idx_livello, co_id, co_F_key_id, co_f_table_id, tab_name, co_titolo_it, tab_titolo " + _
			  " FROM v_indice " + _
			  " WHERE idx_modAdmin_id = " & dummy_admin_id & _
			  " ORDER BY idx_livello, idx_id "
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		'response.write sql & "<BR>"
		if not rs.eof then
			conn.beginTrans
			%>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
				<tr>
					<td class="label_no_width" colspan="2">Numero record già aggiornati:</td>
					<td class="content"><%= totale_conta %></td>
					<td class="content" colspan="2">alle: <%= DateTimeIta(Now()) %></td>
				</tr><tr>
					<td class="label_no_width" colspan="2">Numero record da aggiornare:</td>
					<td class="content_b" colspan="2"><%= totale_manca %></td>
				</tr>
				<tr>
					<th style="width:5%">n°</th>
					<th style="width:15%">Id</th>
					<th style="width:45%">Contenuto</th>
					<th style="width:25%">Tabella</th>
					<th style="width:10%">Livello</th>
				</tr>
			</table>
			<%
			while not rs.eof
				%>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
					<tr>
						<td style="width:5%" class="content"><%= rs.Absoluteposition %></td>
						<td style="width:15%" class="content"><%= rs("idx_id") %></td>
						<td style="width:45%" class="content"><%= rs("co_titolo_it") %></td>
						<td style="width:25%" class="content"><%= rs("tab_titolo") %></td>
						<td style="width:10%" class="content"><%= rs("idx_livello") %></td>
					</tr>
				</table>
				<%
				'sql = "SELECT idx_modAdmin_id FROM tb_contents_index WHERE idx_id=" & rs("idx_id")
				'response.write "admin prima:" & getvaluelist(conn, NULL, sql)& "<BR>"
				
				Index.DisableRicorsione = true
				Session("ID_ADMIN") = nextaim_admin_id
				CALL Index_UpdateItem(conn, rs("tab_name"), rs("co_F_key_id"), false)
				
				sql = " UPDATE tb_contents_index SET idx_modAdmin_id=" & nextaim_admin_id & _
				      " FROM tb_contents_index INNER JOIN tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id " & _
					  " INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " & _
					  " WHERE co_F_key_id = " & rs("co_f_key_id") & " AND tab_name LIKE '" & rs("tab_name") & "'" &_
					  " AND idx_livello<=" & rs("idx_livello")
				'response.write sql& "<BR>"
				CALL conn.execute(sql)
				
				'sql = "SELECT idx_modAdmin_id FROM tb_contents_index WHERE idx_id=" & rs("idx_id")
				'response.write "admin dopo:" & getvaluelist(conn, NULL, sql)& "<BR>"
				'idx_id = rs("idx_id")
				rs.movenext
			wend

			conn.CommitTrans
			
			'sql = "SELECT idx_modAdmin_id FROM tb_contents_index WHERE idx_id=" & idx_id
			'response.write "admin ultimo:" & getvaluelist(conn, NULL, sql)& "<BR>"
			'response.end
			ricarica = true
		else
			ricarica = false
			%>
			<table cellspacing="1" cellpadding="0" class="tabella_madre">
				<tr>
					<td class="content">aggiornamento eseguito correttamente.</td>
					<td class="content_center">&nbsp;</td>
				</tr>
			</table>
			<%
		end if
		rs.close 
		%>
		
	</form>
</div>
<% 
if ricarica then
	%>
	<script language="JavaScript" type="text/javascript">
		document.location.reload(true);
	</script>
	<%
end if
%>
<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>