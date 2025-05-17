<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="Tools_IndexContent.asp" -->
<% 
'imposta elenco di schemi da visualizzare
dim Conn, rs, rsi, rsp, rsc, sql, Pager
set Pager = new PageNavigator
set Conn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.Recordset")
set rsi = Server.CreateObject("ADODB.Recordset")
set rsp = Server.CreateObject("ADODB.Recordset")
set rsc = Server.CreateObject("ADODB.Recordset")
Conn.Open Application("DATA_ConnectionString"), "", ""
set Index.conn = conn
set Index.Content.conn = conn

if request.querystring("OPERAZIONE")<>"" AND cIntero(Session("ID_ADMIN"))>0  then
    SELECT CASE request.QueryString("OPERAZIONE")
        CASE "VERIFICA_CANCELLA"
            'verifica i contenuti e le voci dell'indice 
            
            conn.begintrans
            
            sql = index.QueryElenco(false, "")
            rs.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdtext
            while not rs.eof
                 if not Index.Content.IsValid(rs("co_id")) then
                    'contenuto non valido: lo cancella compreso voci collegate e figlie
					response.write "<!-- " + vbCrLf + _
								   "cancellazione contenuto: " & rs("co_id") & vbCrLF + _
								   rs("co_titolo_it") & vbCrLF + _
								   "-->"
                    CALL Index.Content.Delete(rs("co_id"))
                 end if
                 
                rs.movenext
            wend
            rs.close
            conn.committrans
		case "SBLOCCA_NODO"
			conn.begintrans
			
			'rimuove vincoli pubblicazioni automatiche da un nodo
			sql = "DELETE FROM rel_index_pubblicazioni WHERE rip_idx_id=" & cIntero(request("NODO_ID"))
			CALL conn.execute(sql)
			sql = "UPDATE tb_contents_index SET idx_autopubblicato=0 WHERE idx_id=" & cIntero(request("NODO_ID"))
			CALL conn.execute(sql)
			
            conn.committrans
    END SELECT
end if
%>
<html>
<head>
	<title>Verifica contenuti dell'indice</title>
	<link rel="stylesheet" type="text/css" href="../stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0" onload="window.focus();">
<table width="100%" cellspacing="1" cellpadding="0" style="margin-bottom:10px;">
    <caption style="border:0px;">
        <a style="float:right;" href="javascript:close();" class="menu" name="top">CHIUDI</a>
	</caption>
</table>
<% if cIntero(Session("ID_ADMIN"))>0 then %>
    <script type="text/javascript" language="JavaScript">
        function ConfermaCancellazioneContenutiNonValidi(){
            if (confirm("ATTENZIONE: verranno cancellati tutti i contenuti e le voci dell\'indice non piu\' valide."))
                document.location = "?OPERAZIONE=VERIFICA_CANCELLA";
        }
    </script>
    <table cellpadding="0" cellspacing="1" class="tabella_madre" style="margin-bottom:10px;">
        <caption class="border">Operazioni globali</caption>
        <tr>
            <td class="label_no_width" style="width:65%;">
                Verifica tutti gli elementi dell'indice, cancellando contenuti non validi comprensivi di voci collegate ed eventuali rami dipendenti dalle voci cancellate.
            </td>
            <td class="content_right">
                <a class="button" href="javascript:void(0);" onclick="ConfermaCancellazioneContenutiNonValidi()" title="Cancellazione elementi non validi">VERIFICA CONTENUTI E CANCELLA NON VALIDI</a>
            </td>
        </tr>
    </table>
    <table cellpadding="0" cellspacing="1" class="tabella_madre" style="margin-bottom:20px;">
        <% sql = index.QueryElenco(false, "")
        CALL Pager.OpenSmartRecordset(conn, rs, sql, 40)
        if not rs.eof then%>
            <caption>Voci dell'indice - Trovati n&ordm; <%= Pager.recordcount %> records in n&ordm; <%= Pager.PageCount %> pagine</caption>
            <tr>
               <th>Voce</th> 
               <th class="center" style="width:14%;">LINK</th>
               <th class="center">TIPO CONTENUTO</th>
               <th class="center" >COLLEGAMENTI</th>
               <th class="center" >VOCE</th>
               <th class="center" >CONTENUTO</th>
            </tr>
            <% rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo
                sql = "SELECT * FROM v_indice WHERE idx_id=" & cIntero(rs("idx_id"))
                rsi.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext%>
                <tr>
                    <td class="content" title="id contentuto = <%= rs("co_id") %><%= vbCrLF %> id indice = <%= rs("idx_id") %>" <% if rsi("idx_autopubblicato") then %> rowspan="2" <% end if %>><%= rs("NAME") %></td>
                    <% 'verifica del link
                    if cIntero(rsi("idx_link_pagina_id"))>0 then
                        'link interno: verifica se esiste la pagina.
                        sql = "SELECT * FROM tb_pagineSito WHERE id_paginesito=" & cIntero(rsi("idx_link_pagina_id"))
                        rsp.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
                        if rsp.eof then %>
                            <td class="content_center alert">link interno non valido.<br>Pagina non trovata.</td>
                        <% else %>
                            <td class="content_Center ok">link interno valido.</td>
                        <% end if
                        rsp.close
                    else
                        'link esterno: non fa verifica 
						%>
                        <td class="content_center">esterno</td>
                    <% end if
                    
                    'verifica del contenuto
                    if index.content.IsRaggruppamento(rsi("tab_name")) then 
                        'raggruppamento
						%>
                        <td class="content_center" colspan="3">raggruppamento</td>
                    <% else
                        'contenuto esterno 
						%>
                        <td class="content" style="color:<%= rsi("tab_colore") %>; width:15%;">
                            <%= rsi("tab_titolo") %>
                            <% if not Index.Content.IsValid(rsi("co_id")) then %>
                                <span class="alert" style="display:block;">Contenuto non trovato</span>
                            <% end if %>
                        </td>
                        <td class="content_center" style="width:140px;">
                            <% CALL index.WriteButton(rsi("tab_name"), rsi("co_F_key_id"), POS_ELENCO) %>
                        </td>
                        <td class="content_center" style="width:80px;">
                            <% CALL index.WriteDeleteButton("", rs("idx_id")) %>
                        </td>
                    <% end if %>
                    <td class="content_center" style="width:70px;">
                        <% CALL index.content.WriteDeleteButton("", rsi("idx_content_id")) %>
                    </td>
                </tr>
                <% if rsi("idx_autopubblicato") then %>
                    <tr>
						<td class="content warning" colspan="5">
							<span style="float:right">
								<a href="?OPERAZIONE=SBLOCCA_NODO&NODO_ID=<%= rsi("idx_id") %>" class="button_L2" title="sblocca il vincolo con le pubblicazioni automatiche." <%= ACTIVE_STATUS %>>
									SBLOCCA
								</a>
							</span>
							Bloccato da pubblicazione automatica
						</td>
					</tr>
                <%end if
                rsi.close
                rs.movenext
            wend %>
            <tr>
			    <td colspan="6" class="footer" style="text-align:left;">
			        <% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
		        </td>
			</tr>
        <% else %>
            <caption>Voci dell'indice</caption>
            <tr><td class="noRecords">Nessun record trovato</th></tr>
        <% end if
        rs.close %>
    </table>
<% else %>
    <table cellpadding="0" cellspacing="1" class="tabella_madre" style="margin-bottom:20px;">
        <tr>
            <td class="header alert">Utente non autenticato. Accedere all'area amministrativa.</td>
        </tr> 
    </table>
<% end if %> 

</body>
</html>
<% 
conn.close
set rs = nothing
set rsi = nothing
set rsp = nothing
set rsc = nothing
set conn = nothing
%>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>

