<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
dim conn, rs, rsp, rsc, rstot, rscont, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.Recordset")
set rsc = Server.CreateObject("ADODB.Recordset")
set rsp = Server.CreateObject("ADODB.Recordset")
set rstot = Server.CreateObject("ADODB.Recordset")
set rscont = Server.CreateObject("ADODB.Recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_ARTICOLI_SQL"), "art_id", "ArticoliCommenti.asp")
end if

response.buffer = false
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="ListiniPrezzi_tools.asp" -->
<% 	

sql = " SELECT * FROM gtb_articoli INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " + _
	  " INNER JOIN gtb_iva ON gtb_articoli.art_iva_id = gtb_iva.iva_id " + _
	  " LEFT JOIN gtb_scontiq_classi ON gtb_articoli.art_scontoQ_id = gtb_scontiq_classi.scc_id " + _
	  " WHERE art_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

dim dicitura, tipo, listino
if rs("art_se_bundle") then
	tipo = "bundle"
elseif rs("art_se_confezione") then
	tipo = "confezione"
elseif rs("art_varianti") then
	tipo ="articolo con varianti"
else
	tipo ="articolo singolo"
end if

set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione articoli - commenti " & tipo
dicitura.puls_new = "INDIETRO;SCHEDA ARTICOLO;"
dicitura.link_new = "Articoli.asp;ArticoliMod.asp?ID=" & request("ID")
dicitura.puls_2a_riga.Add "PREZZI","ArticoliPrezzi.asp?ID=" & request("ID")
dicitura.puls_2a_riga.Add "GIACENZE","ArticoliGiacenze.asp?ID=" & request("ID")
if Session("ATTIVA_FAQ_ARTICOLI") then
	dicitura.puls_2a_riga.Add "FAQ","ArticoliFaq.asp?ID=" & request("ID")
end if
dicitura.scrivi_con_sottosez()

%>

<div id="content_abbassato">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati <%= tipo %> con codice &quot;<%= rs("art_cod_int") %>&quot;</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="articolo precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="articolo successiva" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="7">DATI DELL'ARTICOLO</th></tr>
		<% CALL ArticoloScheda (conn, rs, rsp) %>
	</table>
	<% sql = " SELECT * FROM v_indice INNER JOIN tb_comments ON v_indice.idx_id=tb_comments.com_idx_id " + _
			 " WHERE tab_name like 'gtb_articoli' AND co_F_key_id =" & cIntero(request("ID"))
	   
	   rsc.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	%>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
		<% sql = " SELECT COUNT(co_F_key_id) AS tot_comm_art FROM v_indice INNER JOIN tb_comments " + _
				 " ON v_indice.idx_id=tb_comments.com_idx_id WHERE tab_name like 'gtb_articoli' AND co_F_key_id =" & request("ID")
				 
		   rstot.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		%>
		<caption class="border">Elenco dei commenti collegati all'articolo</caption>
		<tr colspan="4">
			<% if not rstot.eof then %> 
				<td class="label_no_width" colspan="2">Trovati n&ordm; <%= rstot("tot_comm_art")%> Commenti </td>		
			<% else %>
				<td class="label_no_width" colspan="2">Nessun Commento trovato</td>		
			<% end if %>
			<td class="Content_right" colspan="2">
				<a class="button" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ArticoliCommentiNew.asp?ARTICOLOID=<%= request("ID") %>', '_blank', 780, 500, true)">
					NUOVO COMMENTO
				</a>
			</td>
		</tr>
		
		<tr colspan="4">
			<th colspan="1" style="width:20%">CONTATTO</th>
			<th colspan="1">DOMANDA</th>
			<th class="center" colspan="2" style="width: 20%;">OPERAZIONI</th>
		</tr>
		<% while not rsc.eof %>
			<tr colspan="4">
				<% sql = " SELECT * FROM tb_indirizzario i LEFT JOIN tb_utenti u " + _
						 " ON i.idElencoIndirizzi = u.ut_nextCom_id WHERE IdElencoIndirizzi = " & rsc("com_contatto_id")
				   rscont.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText %>
				<td class="label_no_width"><%= rscont("NomeOrganizzazioneElencoIndirizzi")%></td>
				<td class="label_no_width"><%= rsc("com_comment")%></td>
				<td class="Content_center">
					<a class="button" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ArticoliCommentiMod.asp?ARTICOLOID=<%= request("ID") %>&ID=<%=rsc("com_id")%>', '_blank', 780, 500, true)">
						MODIFICA
					</a>
				</td>
				<td class="Content_center">
					<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('ARTICOLICOMMENTI','<%= rsc("com_id") %>');" >
						CANCELLA
					</a>
				</td>
			</tr>
			<% rsc.movenext
			rscont.close
		wend 
		rsc.close %>
	</table>
	
	
