<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
dim conn, rs, rsp, rsf, rstot, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.Recordset")
set rsf = Server.CreateObject("ADODB.Recordset")
set rsp = Server.CreateObject("ADODB.Recordset")
set rstot = Server.CreateObject("ADODB.Recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("B2B_ARTICOLI_SQL"), "art_id", "ArticoliFaq.asp")
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
dicitura.sezione = "Gestione articoli - FAQ " & tipo
dicitura.puls_new = "INDIETRO;SCHEDA ARTICOLO;"
dicitura.link_new = "Articoli.asp;ArticoliMod.asp?ID=" & request("ID")
dicitura.puls_2a_riga.Add "PREZZI","ArticoliPrezzi.asp?ID=" & request("ID")
dicitura.puls_2a_riga.Add "GIACENZE","ArticoliGiacenze.asp?ID=" & request("ID")
if Session("ATTIVA_COMMENTI") then
	dicitura.puls_2a_riga.Add "COMMENTI","ArticoliCommenti.asp?ID=" & request("ID")
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
	<% sql = " SELECT * FROM grel_art_faq INNER JOIN tb_FAQ ON grel_art_faq.raf_faq_id = tb_FAQ.faq_id INNER JOIN " + _
             " gtb_articoli ON grel_art_faq.raf_art_id = gtb_articoli.art_id WHERE grel_art_faq.raf_art_id =" & cIntero(request("ID"))
	   
	   rsf.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	%>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
		<% sql = " SELECT COUNT(raf_id) AS tot_faq_art FROM grel_art_faq WHERE (raf_art_id = " & cIntero(request("ID")) & ")"
		
		   rstot.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		%>
		<caption>Elenco FAQ collegate all'articolo</caption>
		<tr colspan="3">
			<% if not rsf.eof then %>
				<td class="label_no_width" colspan="1">Trovate n&ordm; <%= rstot("tot_faq_art")%> FAQ </td>	
			<% else %>
				<td class="label_no_width" colspan="1">Nessuna FAQ trovata </td>
			<% end if %>
			<td class="Content_right" colspan="2">
				<a class="button" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('../nextFaq/Faq.asp?SOURCE=<%=NEXTB2B%>&ARTICOLOID=<%= request("ID") %>', '_blank', 780, 500, true)">
					NUOVA FAQ
				</a>
			</td>
		</tr>
		<tr colspan="3">
			<th>DOMANDA</th>
			<th class="center" colspan="2" style="width: 20%;">OPERAZIONI</th>
		</tr>
		<% while not rsf.eof %>
			<tr>
				<td class="label_no_width"><%= rsf("faq_domanda_IT")%></td>
				<td class="Content_center">
					<a class="button" href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('../nextFaq/FaqMod.asp?ID=<%=rsf("raf_faq_id")%>&SOURCE=<%=NEXTB2B%>', '_blank', 780, 500, true)">
						MODIFICA
					</a>
				</td>
				<td class="Content_center">
					<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('ARTICOLIFAQ','<%= rsf("raf_id") %>');" >
						CANCELLA
					</a>
				</td>
			</tr>
			<% rsf.movenext
		wend 
		rsf.close %>
	</table>
	
	
