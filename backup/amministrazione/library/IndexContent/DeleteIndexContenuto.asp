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
sezione_testata = "cancellazione contenuto dell'indice"
testata_show_back = false %>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

dim conn, rs, rsv, rsd, sql
set rs = server.CreateObject("ADODB.Recordset")
set rsv = server.CreateObject("ADODB.Recordset")
set rsd = server.CreateObject("ADODB.Recordset")
set conn = index.conn


if request("CONFERMA")<>"" then
    'cancellazione della voce e di tutti i figli.
    CALL index.content.Delete(request("ID")) %>
    
    <script language="JavaScript">
        opener.location.reload(true);
        window.close();
    </script>
    
    <%response.end
end if

'legge contenuto
sql = " SELECT * FROM tb_contents WHERE co_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext

'legge pubblicazioni del contenuto
sql = "SELECT * FROM tb_contents_index WHERE idx_content_id=" & cIntero(request("ID"))
rsv.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext

if not rsv.eof then
    'legge discendenti delle pubblicazioni del contenuto
    sql = "SELECT * FROM tb_contents_index WHERE ("
    while not rsv.eof
        sql = sql + SQL_IdListSearch(conn, "idx_tipologie_padre_lista", rsv("idx_id"))
        rsv.movenext
        if not rsv.eof then
            sql = sql + " OR "
        end if
    wend
    sql = sql & ")"
    rsv.movefirst
    rsd.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
end if
%>
<div id="content_ridotto">
    <table cellpadding="0" cellspacing="1" class="tabella_madre">
        <caption class="border">Cancellazione contenuto dell'indice:</caption>
        <tr>
    		<td class="content_center" style="padding:4px;"><img src="<%= GetAmministrazionePath() %>grafica/alert_anim.gif"></td>
        </tr>
    	<tr>
            <td class="content" style="padding:4px;">
                Cancellare il contenuto "<strong><%= rs("co_titolo_it") %></strong>"?
            </td>
        </tr>
        <% if not rsv.eof then %>
             <tr>
                <td class="content" style="padding:4px;">
                    <strong>ATTENZIONE:</strong><br>
                    Sono presenti n&ordm;<strong><%= rsv.recordcount %></strong> pubblicazioni sull'indice del contenuto.<br>
                    Cancellando la voce corrente verranno cancellate anche le n&ordm;<strong><%= rsv.recordcount %></strong> pubblicazioni
                    <% if (rsd.recordcount - rsv.recordcount)>0 then %>e tutte le loro n&ordm;<strong><%= rsd.recordcount - rsv.recordcount %></strong> sottovoci<% end if %>.
                </td>
            </tr>
        <% end if %>
        <script language="JavaScript" type="text/javascript">
            function conferma_onclick(){
                <% if not rsv.eof then %>if (window.confirm("Verranno cancellate anche le pubblicazioni del contenuto.\nCONTINUARE?"))<% end if %>{
                    document.location = "DeleteIndexContenuto.asp?ID=<%= request("ID") %>&CONFERMA=1";
                }
            }
        </script>
    	<tr>
            <td class="content_center">
                <table cellpadding="10">
                    <tr>
                        <td style="width:45%; text-align:right;">
                            <a accesskey="c" tabindex="1" href="javascript:void(0);" class="button"
                               onclick="conferma_onclick();"
                               title="Esegue la cancellazione del contenuto" <%= ACTIVE_STATUS %> id="primo_elemento">
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
<br>
</div>
<%

rs.close
if not rsv.eof then
    rsd.close
end if
rsv.close
set rs = nothing
set rsv = nothing
set rsd = nothing

%>

</body>
</html>
<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>