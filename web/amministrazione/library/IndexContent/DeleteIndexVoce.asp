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
sezione_testata = "cancellazione voce dell'indice"
testata_show_back = false %>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim conn, rs, sql
set rs = server.CreateObject("ADODB.Recordset")
set conn = index.conn

if request("CONFERMA")<>"" then
    'cancellazione della voce e di tutti i figli.
    CALL index.Delete(request("ID")) %>
    
    <script type="text/javascript">
		//opener.document.location.reload(true);
		opener.document.location.href = opener.document.location.href + '&MODE=standard';
        window.close();
    </script>
    
    <%response.end
end if


sql = " SELECT *, " + _
      " (SELECT COUNT(*) FROM tb_contents_index WHERE idx_padre_id = " & cIntero(request("ID")) & ") AS N_FIGLI, " + _
      " (SELECT COUNT(*) FROM tb_contents_index WHERE " + SQL_IdListSearch(conn, "idx_tipologie_padre_lista", cIntero(request("ID"))) + ") AS N_DISCENDENTI " + _
      " FROM v_indice WHERE idx_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
%>
<div id="content_ridotto">
    <table cellpadding="0" cellspacing="1" class="tabella_madre">
        <caption class="border">Cancellazione voce dell'indice:</caption>
        <tr>
    		<td class="content_center" style="padding:4px;"><img src="<%= GetAmministrazionePath() %>grafica/alert_anim.gif"></td>
        </tr>
    	<tr>
            <td class="content" style="padding:4px;">
                Cancellare la voce dell'indice "<strong><%= rs("co_titolo_it") %></strong>"?
            </td>
        </tr>
        <% if cIntero(rs("N_FIGLI"))>0 then %>
            <tr>
                <td class="content" style="padding:4px;">
                    <strong>ATTENZIONE:</strong><br>
                    Sono presenti n&ordm;<strong><%= rs("N_FIGLI") %></strong> voci collegate alla voce corrente.<br>
                    Cancellando la voce corrente verranno cancellate anche le n&ordm;<strong><%= rs("N_FIGLI") %></strong> voci collegate 
                    <% if (cIntero(rs("N_DISCENDENTI")) - cIntero(rs("N_FIGLI")) - 1) > 0 then %>e le relative n&ordm;<strong><%= (cIntero(rs("N_DISCENDENTI")) - cIntero(rs("N_FIGLI")) - 1) %></strong> sotto-voci<% end if%>.
                </td>
            </tr>
        <% end if %>
        <script language="JavaScript" type="text/javascript">
            function conferma_onclick(){
                <% if cIntero(rs("N_FIGLI"))>0 then %>if (window.confirm("Verranno cancellate anche le voci collegate.\nCONTINUARE?"))<% end if %>{
					document.location = "DeleteIndexVoce.asp?ID=<%= cIntero(request("ID")) %>&CONFERMA=1";
                }
            }
        </script>
    	<tr>
            <td class="content_center">
                <table cellpadding="10">
                    <tr>
                        <td style="width:45%; text-align:right;">
                            <a accesskey="c" tabindex="1" href="#" class="button"
                               onclick="conferma_onclick();"
                               title="Esegue la cancellazione della voce" <%= ACTIVE_STATUS %> id="primo_elemento">
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
            <td class="content_center">
				<% CALL WriteCopiaIndirizziAlternativi(request("ID"),"vuoi recuperare gli URL relativi a questa voce dell'indice?","RECUPERA URL") %>
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
