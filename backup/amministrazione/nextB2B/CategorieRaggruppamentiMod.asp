<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("CategorieRaggruppamentiSalva.asp")
end if

dim i, conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")
sql = "SELECT * FROM gtb_tipologie_raggruppamenti WHERE rag_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>

<%'--------------------------------------------------------
sezione_testata = "modifica raggruppamento"
testata_show_back = true %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Modifica raggruppamento per la categoria</caption>
			<tr><th colspan="3">DATI RAGGRUPPAMENTO</th></tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
				<% 	if i = 0 then %>
					<td class="label" style="width:22%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
				<% 	end if %>
					<td class="content" colspan="2">
						<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
						<input type="text" class="text" name="tft_rag_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("rag_nome_"& Application("LINGUE")(i)) %>" maxlength="50" size="55">
						<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
					</td>
				</tr>
			<%next %>
			<tr>
				<td class="label" rowspan="2">dati pubblicazione:</td>
				<td class="label">pubblicato:</td>
				<td class="content">
					<input type="radio" class="checkbox" value="1" name="tfn_rag_visibile" <%= chk(rs("rag_visibile")) %>>
					si
					<input type="radio" class="checkbox" value="0" name="tfn_rag_visibile" <%= chk(not rs("rag_visibile")) %>>
					no
				</td>
			</tr>
			<tr>
				<td class="label">ordine</td>
				<td class="content">
					<input type="text" class="text" name="tfn_rag_ordine" value="<%= rs("rag_ordine") %>" maxlength="4" size="4">
				</td>
			</tr>
			<tr>
				<td class="label">foto:</td>
				<td class="content" colspan="2">
					<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_rag_foto", rs("rag_foto"), "", FALSE) %>
				</td>
			</tr>
			<tr><th colspan="3">DESCRIZIONE</th></tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
					<td class="content" colspan="3">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>
								<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<td><textarea style="width:100%;" rows="3" name="tft_rag_descr_<%= Application("LINGUE")(i) %>"><%= rs("rag_descr_" & Application("LINGUE")(i)) %></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			<%next %>
			<tr>
				<td class="footer" colspan="3">
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

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>
<%
set rs = nothing
conn.Close
set conn = nothing
%>