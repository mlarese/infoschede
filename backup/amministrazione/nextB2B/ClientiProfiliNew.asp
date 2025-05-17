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
dicitura.sezione = "Gestione profili - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "ClientiProfili.asp"
dicitura.scrivi_con_sottosez() 

dim i, sql, conn, rs
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" id="">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>Inserimento nuovo profilo</caption>
		<tr><th colspan="3">DATI DEL PROFILO</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
    		<tr>
    		<% 	if i = 0 then %>
    			<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
    		<% 	end if %>
    			<td class="content" colspan="2">
    				<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
    				<input type="text" class="text" name="tft_pro_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_pro_nome_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
    				<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
    			</td>
    		</tr>
		<%next %>
		<tr>
			<td class="label" style="width:20%;">codice:</td>
			<td class="content" colspan="2">
		    	<input type="text" class="text" name="tft_pro_codice" value="<%= request("tft_pro_codice")%>" maxlength="100" size="25">
			</td>
		</tr>
        <tr><th colspan="3">AREA RISERVATA SUL PORTALE</th></tr>
		<tr>
			<td class="label">home page:</td>
            <td colspan="2" class="content">
                <% CALL DropDownPages(conn, "form1", "250", Session("AZ_ID"), "nfn_pro_pagina_id", request("tfn_pro_pagina_id"), false, false) %>
            </td>
        </tr>
    </table>
    <table cellspacing="1" cellpadding="0" class="tabella_madre">
		<tr>
			<td class="footer">
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
<% 
conn.close
set conn = nothing 
%>