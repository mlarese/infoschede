<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_indice_accesso, 0))

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Indice generale - gestione voci collegate"
dicitura.puls_new = "INDIETRO;SCHEDA"
dicitura.link_new = IIF(request("FROM")=FROM_ALBERO, "IndexAlbero.asp", "IndexGenerale.asp") &";IndexGestione.asp?ID=" & request("ID") & "&FROM=" & request("FROM")
dicitura.scrivi_con_sottosez()

dim rs, sql
set rs = server.CreateObject("ADODB.recordset")
    
sql = " SELECT *, (SELECT COUNT(*) FROM tb_contents_index WHERE idx_padre_id=i.idx_id) AS N_FIGLI " & _
	  " FROM (tb_contents_index i"& _
	  " INNER JOIN tb_contents c ON i.idx_content_id = c.co_id)"& _
	  " INNER JOIN tb_siti_tabelle t ON c.co_F_table_id = t.tab_id"& _
	  " WHERE idx_padre_id = "& cIntero(request("ID")) & _
	  " ORDER BY co_titolo_it"
Session("IDX_SQL") = sql
rs.open sql, Index.Conn, adOpenStatic, adLockReadOnly, adCmdText %>
	<div id="content">
    	<form action="" method="post" id="form1" name="form1">
    	<table cellspacing="1" cellpadding="0" class="tabella_madre">
    		<caption>Elenco voci collegate a "<%= Index.NomeCompleto(request("ID")) %>"</caption>
    		<tr><th colspan="7">ELENCO VOCI COLLEGATE</th></tr>
    		<tr>
    			<td colspan="2" class="label_no_width">
    				<% if rs.eof then %>
    					Nessuna voce presente.
    				<% else %>
    					Trovate n&ordm; <%= rs.recordcount %>&nbsp; voci
    				<% end if %>
    			</td>
    			<td colspan="3" class="content_right" style="padding-right:0px;">
   					<a class="button_L2" target="_blank" href="IndexGestione.asp?&FROM=<%= FROM_ELENCO %>&idx_padre_id=<%= request("ID") %>&SOTTO=1">
   						NUOVA VOCE
   					</a>
   				</td>
   			</tr>
    		<% if not rs.eof then %>
   				<tr>
   					<th>NOME</th>
   					<th class="center" colspan="3">OPERAZIONI</th>
   				</tr>
   				<% while not rs.eof %>
   					<tr>
   						<td class="content"><% CALL index.content.WriteNomeETipo(rs) %></td>
   						<td class="content_center" style="width:95px;">
   							<a class="button_L2" target="_blank" href="IndexSottosezioni.asp?ID=<%= rs("idx_id") %>" title="Apre elenco delle sotto-voci." <%= ACTIVE_STATUS %>>
   								VOCI COLLEGATE
   							</a>
   						</td>
   						<td class="content_center" style="width:50px;">
   							<a class="button_L2" target="_blank" href="IndexGestione.asp?FROM=<%= FROM_ELENCO %>&ID=<%= rs("idx_id") %>&SOTTO=1">
   								MODIFICA
   							</a>
   						</td>
    					<td class="content_center" style="width:50px;">
							<% if rs("idx_autopubblicato") then %>
	   	    					<a class="button_L2_disabled" href="javascript:void(0);" title="Impossibile cancellare la voce perch&egrave; fa parte delle seguenti pubblicazioni automatiche:<%= vbCrLF & index.GetPubblicazioniLockers(rs("idx_id")) %>."<%= ACTIVE_STATUS %>>
	   								CANCELLA
	   							</a>
   							<% else
                               CALL index.WriteDeleteButton("_L2", rs("idx_id"))
                           	end if %>
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
set rs = nothing
%>