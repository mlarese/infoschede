<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<% 
dim conn, rs, sql, tipo
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	dim nuova_classe, aggiornamento
	if cInteger(request("classe_sconto"))>0 then
		nuova_classe = cInteger(request("classe_sconto"))
	else
		nuova_classe = NULL
	end if
	
	conn.beginTrans
	sql = "SELECT * FROM gtb_articoli WHERE art_id=" & cIntero(request("ID"))
	rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
	'aggiorna classe di sconto base dell'articolo
	rs("art_scontoQ_id") = nuova_classe
	rs.update
	rs.close
	
	'aggiorna classi di sconto delle varianti
	aggiornamento = cInteger(request("aggiornamento_varianti"))
	Select case aggiornamento
		case 1
			'sostituisce una classe particolare
			sql = " UPDATE grel_Art_valori SET rel_scontoQ_id = " & IIF(isNull(nuova_classe), "NULL", nuova_classe) & _
				  " WHERE rel_art_id=" & cIntero(request("ID")) & " AND "
			if cInteger(request("classe_variante")) = 0 then
				sql = sql & " (" & SQL_isNull(conn, "rel_scontoQ_id") & " OR rel_scontoq_id=0) "
			else
				sql = sql & " rel_scontoQ_id=" & cInteger(request("classe_variante"))
			end if
			CALL conn.execute(sql, , adExecuteNoRecords)
		case 2
			'sostituisce tutte le classi di sconto
			sql = " UPDATE grel_Art_valori SET rel_scontoQ_id = " & IIF(isNull(nuova_classe), "NULL", nuova_classe) & _
				  " WHERE rel_art_id=" & cIntero(request("ID"))
			CALL conn.execute(sql, , adExecuteNoRecords)
	end select
	
	'aggiorna classi di sconto nei listini
	aggiornamento = cInteger(request("aggiornamento_listini"))
	Select case aggiornamento
		case 1
			'sostituisce una classe particolare
			sql = " UPDATE gtb_prezzi SET prz_scontoQ_id = " & IIF(isNull(nuova_classe), "NULL", nuova_classe) & _
				  " WHERE prz_variante_id IN (SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & cIntero(request("ID")) & ") AND " 
			if cInteger(request("classe_listino")) = 0 then
				sql = sql & " (" & SQL_isNull(conn, "prz_scontoQ_id") & " OR prz_scontoQ_id=0) "
			else
				sql = sql & " prz_scontoQ_id=" & cInteger(request("classe_listino"))
			end if
			CALL conn.execute(sql, , adExecuteNoRecords)
		case 2
			'sostituisce tutte le classi di sconto
			sql = " UPDATE gtb_prezzi SET prz_scontoQ_id = " & IIF(isNull(nuova_classe), "NULL", nuova_classe) & _
				  " WHERE prz_variante_id IN (SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & cIntero(request("ID")) & ")" 
			CALL conn.execute(sql, , adExecuteNoRecords)
	end select
	
	conn.commitTrans %>
	<script language="JavaScript" type="text/javascript">
		opener.location.reload(true);
		window.close();
	</script>
<% end if


sql = "SELECT * FROM gtb_articoli WHERE art_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

if rs("art_se_bundle") then
	tipo = "del bundle"
elseif rs("art_se_confezione") then
	tipo = "della confezione"
elseif rs("art_varianti") then
	tipo ="dell'articolo con varianti"
else
	tipo ="dell'articolo singolo"
end if
%>

<%'--------------------------------------------------------
sezione_testata = "modifica classe di sconto " & tipo %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>
				Modifica della classe di sconto per quantit&agrave; <%= tipo %>
			</caption>
			<tr><th colspan="3">DATI BASE ARTICOLO</th></tr>
			<tr>
				<td class="label" colspan="2" rowspan="2" style="width:31%;">articolo</td>
				<td class="content"><%= rs("art_nome_it") %></td>
			</tr>
			<tr>
				<% if rs("art_se_bundle") then %>
					<td colspan="2" class="content bundle">bundle</td>
				<% elseif rs("art_se_confezione") then %>
					<td colspan="2" class="content confezione">confezione</td>
				<% elseif rs("art_varianti") then %>
					<td colspan="2" class="content varianti">articolo con varianti</td>
				<% else %>
					<td colspan="2" class="content">articolo singolo</td>
				<% end if %>
			</tr>
			<tr>
				<td class="label"colspan="2" >classe di sconto per quantit&agrave;</td>
				<td class="content">
					<% sql = " SELECT scc_id, scc_nome FROM gtb_scontiQ_classi ORDER BY scc_nome" 
					CALL dropDown(conn, sql, "scc_id", "scc_nome", "classe_sconto", rs("art_scontoQ_id"), false, "", LINGUA_ITALIANO)%>
				</td>
			</tr>
			<% if rs("art_varianti") then %>
				<tr><th colspan="3">AGGIORNAMENTO VARIANTI</th></tr>
				<tr>
					<td class="label" rowspan="3">tipo:</td>
					<td class="content_center">
						<input type="radio" class="checkbox" name="aggiornamento_varianti" value="0" <%= chk(cInteger(request("aggiornamento_varianti"))=0) %> onclick="varianti_abilita_classe();">
					</td>
					<td class="content">
						non aggiornare
					</td>
				</tr>
				<tr>
					<td class="content_center">
						<input type="radio" class="checkbox" name="aggiornamento_varianti" id="aggiornamento_varianti_1" value="1" <%= chk(cInteger(request("aggiornamento_varianti"))=1) %> onclick="varianti_abilita_classe();">
					</td>
					<td class="content">
						sostituisci solo dove la classe &egrave;:
						<% sql = "SELECT (NULL) AS scc_id, ('nessuna classe di sconto applicata') AS scc_nome " + _
								 " UNION " + _
								 " SELECT scc_id, scc_nome FROM gtb_scontiQ_classi ORDER BY scc_nome" 
						CALL dropDown(conn, sql, "scc_id", "scc_nome", "classe_variante", request("classe_variante"), true, IIF(cInteger(request("aggiornamento_varianti"))=1, " ", " disabled"), LINGUA_ITALIANO)%>
					</td>
				</tr>
				<tr>
					<td class="content_center">
						<input type="radio" class="checkbox" name="aggiornamento_varianti" value="2" <%= chk(cInteger(request("aggiornamento_varianti"))=2) %> onclick="varianti_abilita_classe();">
					</td>
					<td class="content">
						sostituisci tutto
					</td>
				</tr>
				<script language="JavaScript" type="text/javascript">
					function varianti_abilita_classe(){
						form1.classe_variante.disabled = !(document.all.aggiornamento_varianti_1.checked);
					}
				</script>
			<% else 
				'se non ci sono varianti la classe di sconto va sostituita nell'esploso%>
				<input type="hidden" name="aggiornamento_varianti" value="2">
			<% end if %>
			<tr><th colspan="3">AGGIORNAMENTO LISTINI</th></tr>
			<tr>
				<td class="label" rowspan="3">tipo:</td>
				<td class="content_center">
					<input type="radio" class="checkbox" name="aggiornamento_listini" value="0" <%= chk(cInteger(request("aggiornamento_listini"))=0) %> onclick="listini_abilita_classe();">
				</td>
				<td class="content">
					non aggiornare
				</td>
			</tr>
			<tr>
				<td class="content_center">
					<input type="radio" class="checkbox" name="aggiornamento_listini" id="aggiornamento_listini_1" value="1" <%= chk(cInteger(request("aggiornamento_listini"))=1) %> onclick="listini_abilita_classe();">
				</td>
				<td class="content">
					sostituisci solo dove la classe &egrave;:
					<% sql = "SELECT (NULL) AS scc_id, ('nessuna classe di sconto applicata') AS scc_nome " + _
							 " UNION " + _
							 " SELECT scc_id, scc_nome FROM gtb_scontiQ_classi ORDER BY scc_nome" 
					CALL dropDown(conn, sql, "scc_id", "scc_nome", "classe_listino", request("classe_listino"), true, IIF(cInteger(request("aggiornamento_listini"))=1, " ", " disabled"), LINGUA_ITALIANO)%>
				</td>
			</tr>
			<tr>
				<td class="content_center">
					<input type="radio" class="checkbox" name="aggiornamento_listini" value="2" <%= chk(cInteger(request("aggiornamento_listini"))=2) %> onclick="listini_abilita_classe();">
				</td>
				<td class="content">
					sostituisci tutto
				</td>
			</tr>
			<script language="JavaScript" type="text/javascript">
				function listini_abilita_classe(){
					form1.classe_listino.disabled = !(document.all.aggiornamento_listini_1.checked);
				}
			</script>
			<tr>
				<td class="footer" colspan="3">
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
					<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
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