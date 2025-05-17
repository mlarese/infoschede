<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ClientiProfiliSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione profili - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "ClientiProfili.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, rsv, sql, i, rsm
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")
set rsv = server.CreateObject("ADODB.recordset")
set rsm = server.CreateObject("ADODB.recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("CLI_PROFILI_SQL"), "pro_id", "ClientiProfiliMod.asp")
end if

sql = "SELECT * FROM gtb_profili WHERE pro_id="& cIntero(request("ID"))
set rs = conn.Execute(sql)
%>


<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
        <caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica profilo</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="profilo precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="profilo successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="3">DATI DEL PROFILO</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
    		<tr>
    		<% 	if i = 0 then %>
    			<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
    		<% 	end if %>
    			<td class="content" colspan="2">
    				<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
    				<input type="text" class="text" name="tft_pro_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("pro_nome_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
    				<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
    			</td>
    		</tr>
		<%next %>
		<tr>
			<td class="label" style="width:20%;">codice:</td>
			<td class="content" colspan="2">
		    	<input type="text" class="text" name="tft_pro_codice" value="<%= rs("pro_codice")%>" maxlength="100" size="25">
			</td>
		</tr>
        <tr><th colspan="3">AREA RISERVATA SUL PORTALE</th></tr>
		<tr>
			<td class="label">home page:</td>
            <td colspan="2" class="content">
                <% CALL DropDownPages(conn, "form1", "250", Session("AZ_ID"), "nfn_pro_pagina_id", rs("pro_pagina_id"), false, false) %>
            </td>
        </tr>
    </table>
    <table cellspacing="1" cellpadding="0" class="tabella_madre">
		<% CALL Form_DatiModifica(conn, rs, "pro_") %>
		<tr>
			<td colspan="4" class="footer">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>