<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	

dim conn, rs, rsc, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")
set rsc = server.CreateObject("ADODB.recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_CATEGORIE_SQL"), "tip_id", "CategorieSottocategorie.asp")
end if

sql =" SELECT *, (SELECT COUNT(*) FROM gtb_articoli WHERE art_tipologia_id=t.tip_id) AS N_ART, " & _
	  " (SELECT COUNT(*) FROM gtb_tipologie WHERE tip_padre_id=t.tip_id) AS N_FIGLI, " & _
	  " (SELECT COUNT(*) FROM gtb_tipologie_raggruppamenti WHERE rag_tipologia_id=t.tip_id) AS N_GRUPPI " & _
	  " FROM gtb_tipologie t WHERE tip_id="& cIntero(request("ID"))
rsc.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione categorie - modifica raggruppamenti della categoria"
dicitura.puls_new = "INDIETRO;SCHEDA"
dicitura.link_new = "Categorie.asp;CategorieMod.asp?ID=" & rsc("tip_id")
if cInteger(rsc("N_ART")) = 0 AND cInteger(rsc("N_GRUPPI"))=0 then
	dicitura.puls_new = dicitura.puls_new + ";SOTTOCATEGORIE"
	dicitura.link_new = dicitura.link_new + ";CategorieSottocategorie.asp?ID=" & rsc("tip_id")
end if
dicitura.scrivi_con_sottosez() 
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Raggruppamenti di "<%= categorie.NomeCompleto(rsc("tip_id")) %>"</caption>
		<tr><th colspan="7">ELENCO RAGGRUPPAMENTI</th></tr>
		<% sql = " SELECT *, " + _
				 " (SELECT COUNT(*) FROM gtb_articoli WHERE art_raggruppamento_id= gtb_tipologie_raggruppamenti.rag_id ) AS N_ART" + _
				 " FROM gtb_tipologie_raggruppamenti WHERE rag_tipologia_id=" & rsc("tip_id")
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext %>
		<tr>
			<td class="label_no_width">
				<% if rs.eof then %>
					Nessun raggruppamento presente.
				<% else %>
					Trovati n&ordm; <%= rs.recordcount %> raggruppamenti
				<% end if %>
			</td>
			<td colspan="4" class="content_right" style="padding-right:0px;">
				<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la selezione ed inserimento del nuovo raggruppamento" <%= ACTIVE_STATUS %>
				   onclick="OpenAutoPositionedScrollWindow('CategorieRaggruppamentiNew.asp?R_ID=<%= rsc("tip_id") %>', 'RAGGRUPPAMENTI', 530, 400, true)">
					NUOVO RAGGRUPPAMENTO
				</a>
			</td>
		</tr>
		<% if not rs.eof then %>
			<tr>
				<th class="L2">NOME</th>
				<th class="l2_center">ORDINE</th>
				<th class="l2_center" colspan="2" style="width:16%;">OPERAZIONI</th>
			</tr>
			<% while not rs.eof %>
				<tr>
					<td class="content"><%= rs("rag_nome_it") %></td>
					<td class="content_center"><%= rs("rag_ordine") %></td>
					<td class="content_center">
						<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la modifica dei dati del raggruppamento" <%= ACTIVE_STATUS %>
						   onclick="OpenAutoPositionedScrollWindow('CategorieRaggruppamentiMod.asp?ID=<%= rs("rag_id") %>', 'RAGGRUPPAMENTO', 510, 250, true)">
							MODIFICA
						</a>
					</td>
					<td class="content_center">
						<% if cInteger(rs("N_ART"))>0 then %>
							<a class="button_L2_DISABLED" href="javascript:void(0);" title="raggruppamento non cancellabile perch&egrave; associato ad almeno un articolo" <%= ACTIVE_STATUS %>
								CANCELLA
							</a>
						<% else %>
							<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione del raggruppamento" <%= ACTIVE_STATUS %>
							   onclick="OpenDeleteWindow('RAGGRUPPAMENTO','<%= rs("rag_id") %>');">
								CANCELLA
							</a>
						<% end if %>
					</td>
				</tr>
				<% rs.movenext
			wend
		end if
		rs.close %>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>

<%
rsc.close
set rs = nothing
set rsc = nothing
conn.Close
set conn = nothing
%>