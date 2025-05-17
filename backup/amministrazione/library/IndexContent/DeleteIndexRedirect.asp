<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_IndexContent.asp" -->
<%
'check dei permessi dell'utente
if NOT index.ChkPrm(prm_indice_accesso, 0) then %>
    <script type="text/javascript">
	    window.close()
    </script>
<% end if

'--------------------------------------------------------
sezione_testata = "cancellazione indirizzo alternativo della voce dell'indice"
testata_show_back = false %>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim conn, rs, sql
set rs = server.CreateObject("ADODB.Recordset")
set conn = index.conn

if request("CONFERMA")<>"" then
    conn.Execute("DELETE FROM rel_index_url_redirect WHERE riu_id = "& cIntero(request("ID"))) %>
    
    <script type="text/javascript">
		opener.document.location.reload(true);
        window.close();
    </script>
    
    <%response.end
end if


sql = " SELECT * FROM rel_index_url_redirect WHERE riu_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
%>
<div id="content_ridotto">
    <table cellpadding="0" cellspacing="1" class="tabella_madre">
        <caption class="border">Cancellazione indirizzo alternativo della voce dell'indice:</caption>
        <tr>
    		<td class="content_center" style="padding:4px;"><img src="<%= GetAmministrazionePath() %>grafica/alert_anim.gif"></td>
        </tr>
    	<tr>
            <td class="content" style="padding:4px;">
                Cancellare l'indirizzo alternativo "<strong><%= rs("riu_url") %></strong>"?
            </td>
        </tr>
    	<tr>
            <td class="content_center">
                <table cellpadding="10">
                    <tr>
                        <td style="width:45%; text-align:right;">
                            <a accesskey="c" tabindex="1" href="?ID=<%= cIntero(request("ID")) %>&CONFERMA=1" class="button"
                               title="Esegue la cancellazione dell'indirizzo alternativo della voce" <%= ACTIVE_STATUS %> id="primo_elemento">
                                CONFERMA
							</a>
						</td>
						<td style="width:10%;">&nbsp;</td>
                        <td style="width:45%; text-align:left;">
                            <a accesskey="a" tabindex="2" href="javascript:window.close();" class="button"
                               title="Chiude la finestra ed annulla la cancellazione" <%= ACTIVE_STATUS %> >
                                ANNULLA
							</a>
						</td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td class="footer">
                <a class="button" href="javascript:close();">CHIUDI</a>
            </td>
        </tr>
    </table>
</div>
<%
rs.close
set rs = nothing
%>
</body>
</html>
<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>