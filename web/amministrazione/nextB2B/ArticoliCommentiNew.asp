<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ArticoliCommentiSalva.asp")
end if

dim i, conn, rs, sql, valore
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")

%>

<%'--------------------------------------------------------
sezione_testata = "inserimento nuovo commento" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

sql = "SELECT TOP 1 idx_id FROM v_indice WHERE tab_name like 'gtb_articoli' AND co_F_key_id =" & cIntero(request("ARTICOLOID"))
valore = GetValueList(conn, NULL, sql )

%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_com_idx_id" value="<%=valore%>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Inserimento nuovo commento</caption>
			<tr><th colspan="3">DATI DEL COMMENTO</th></tr>
			<tr>
				<td class="label_no_width" colspan="2">contatto (*)</td>
				<td class="content">
					<% CALL WriteContactPicker_Input(conn, rs, "", "", "form1", "tfn_com_contatto_id", request("tfn_com_contatto_id"), "", false, false, false, "") %>		
				</td>
			</tr>
			<tr>
				<td class="label_no_width" colspan="2">valutazione</td>
				<td class="content">
					<%CALL dropDown(conn, "SELECT * FROM tb_comments_valutazioni ", _
							"val_id", "val_nome_it", "tfn_com_val_id", request("tfn_com_val_id") , true, " style=""width=250""", LINGUA_ITALIANO)%>
				</td>
			</tr>
			
			<tr>
				<td class="label_no_width" colspan="2">commento</td>
				<td class="content">
					<textarea rows="5" style="width:95%;" name="tft_com_comment"><%= request("tft_com_comment") %></textarea>
					(*)
				</td>
			</tr>
			
			<tr>
				<td class="label_no_width" colspan="2">validato</td>
				<td class="content">
						<input type="checkbox" class="checkbox" name="chk_com_validate" <%= chk(request("chk_com_validate")<>"") %>>
				</td>
			</tr>
			
			<tr>
				<td class="footer" colspan="3">
					(*) Campi obbligatori.
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA" >
					<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
				</td>
			</tr>
		</table>
	</form>
</div>
</body>
</html>
<%
set rs = nothing
%>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>