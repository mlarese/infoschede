<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoRssSalva.asp")
end if
%>

<%'--------------------------------------------------------
sezione_testata = "modifica dati rss" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<% 
dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM tb_rss WHERE rss_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>


<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_rss_web_id" value="<%= request("WEB_ID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Modifica dati rss</caption>
			<tr><th colspan="3">DATI DELL'RSS</th></tr>
			
			<tr>
				<td class="label" colspan="2">lingua:</td>				
				<td class="content" nowrap>
					<% CALL DropLingue(conn, NULL, "tft_rss_lingua", rs("rss_lingua"), true, false, "width:100px;") %>
					(*)
				</td>
			</tr>			
			<tr>
				<td class="label" colspan="2">titolo:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_rss_titolo" value="<%= CBR(rs, "rss_titolo", "tft_") %>" maxlength="255" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">file:</td>
				<td class="content" colspan="3">
					<input type="text" class="text" name="tft_rss_file" value="<%= CBR(rs, "rss_file", "tft_") %>" maxlength="255" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">abilitato:</td>
				<td class="content">
					<input type="radio" class="checkbox" value="1" name="tfn_rss_abilitato" <%= chk(rs("rss_abilitato")) %>>
					si
					<input type="radio" class="checkbox" value="0" name="tfn_rss_abilitato" <%= chk(not rs("rss_abilitato")) %>>
					no
				</td>
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">metatag:</td>
				<td class="content">
					<input type="radio" class="checkbox" value="1" name="tfn_rss_metatag" <%= chk(rs("rss_metatag")) %>>
					si
					<input type="radio" class="checkbox" value="0" name="tfn_rss_metatag" <%= chk(not rs("rss_metatag")) %>>
					no
				</td>
				</td>
			</tr>
			<tr>				
				<td class="label" colspan="2">descrizione:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tft_rss_descrizione" value="<%= CBR(rs, "rss_descrizione", "tft_") %>" maxlength="255" size="100">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">immagine:</td>
				<td class="content" colspan="3">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_rss_image", rs("rss_image") , "width:425px;", false) %>
				</td>
			</tr>
			<tr>
				<td class="label" valign="top" colspan="2">query:</td>
				<td class="content"><textarea style="width:528px;" rows="6" name="tft_rss_query" ><%= rs("rss_query")%></textarea>
				(*)
				</td>
				
			</tr>
			<tr>				
				<td class="label" colspan="2">frequenza di generazione:</td>
				<td class="content" nowrap>
					<input type="text" class="text" name="tfn_rss_freq_generazione" value="<%= rs("rss_freq_generazione") %>" size="10">
					(*)
				</td>
			</tr>
			<tr>				
				<td class="label" colspan="2">data generazione:</td>
				<td class="content" nowrap>
					<% dim val
					if(cString(request("RESET"))<>"") then
						val=DateAdd("n",cIntero(rs("rss_freq_generazione"))*(-1),Now())
					else
						val=CBR(rs, "rss_data_generazione", "tfd_")						
					end if
					%>
					<input type="text" class="text_disabled" name="tfd_rss_data_generazione" value="<%=val%>" size="50">
					<a class="button_L2" href="javascript:void(0)" title="reimposta la data di generazione all'istante attuale meno il valore della frequenza di generazione" <%= ACTIVE_STATUS %>
							   onclick="location.href=location.href+'&RESET=1'">
								RESET
					</a>
				</td>				
					
			</tr>	
			<tr>
				<td class="footer" colspan="4">
					(*) Campi obbligatori.
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
					<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
				</td>
			</tr>			
	</form>
		</table>
</div>
</body>
</html>
<% rs.close
conn.close
set rs = nothing
set conn = nothing %>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>