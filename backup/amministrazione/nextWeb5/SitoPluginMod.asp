<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_plugin_accesso, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoPluginSalva.asp")
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - plugin - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoPlugin.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, i, hide, display
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("WEB_PLUGIN_SQL"), "id_objects", "SitoPluginMod.asp")
end if

sql = "SELECT * FROM tb_objects WHERE id_objects="& cIntero(request("ID"))
set rs = conn.Execute(sql)

hide="visibility:hidden;display:none;"
display="visibility:visible;"
%>

<script language="JavaScript" type="text/javascript">
	function tipo_plugin()
	{		
		var SelectedIndex = form1.tft_obj_type.selectedIndex;
		var html = document.getElementById('html')
		var plugin = document.getElementById('plugin')		
		var sorgente = document.getElementById('sorgente')
		if (SelectedIndex==2){
			html.style.visibility = "visible";
			html.style.display = "";
			plugin.style.visibility = "hidden";
			plugin.style.display = "none";
			sorgente.style.visibility = "hidden";
			sorgente.style.display = "none";
			if(form1.sorg.value=="")
			{
				form1.sorg.value="NextHTMLController";
			}
			if(form1.tft_param_list.value=="")
			{
			    form1.tft_param_list.value=";";
			}
		}
		else{						
			plugin.style.visibility = "visible";
			plugin.style.display = "";
			sorgente.style.visibility = "visible";
			sorgente.style.display = "";
			html.style.visibility = "hidden";
			html.style.display = "none";
		}
	}
</script>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati del plugin</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="plugin precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="plugin successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="2">DATI DEL PLUGIN</th></tr>
		<tr>
			<td class="label">nome:</td>
			<td class="content">
				<input type="text" class="text" name="tft_name_objects" value="<%= rs("name_objects") %>" maxlength="255" size="80">
				(*)
			</td>
		</tr>
		<tr id="sorgente" style="<%= IIF(rs("obj_type") = "html", hide, display) %>">
			<td class="label" >classe sorgente:</td>
			<td class="content">
				<input type="text" id="sorg" class="text" name="tft_identif_objects" value="<%= rs("identif_objects") %>" maxlength="70" size="80">
			</td>
		</tr>
		<tr>
			<td class="label">tipo:</td>
			<td class="content">
				<select name="tft_obj_type" onchange="tipo_plugin()">
					<option value="ascx" <%= IIF(rs("obj_type") = "ascx", "selected", "") %>>Plugin</option>
					<option value="class" <%= IIF(rs("obj_type") = "class", "selected", "") %>>Classe</option>
					<option value="html" <%= IIF(rs("obj_type") = "html", "selected", "") %>>Html</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2" style="padding:0px;">
				<table id="plugin" style="width:100%; <%= IIF(rs("obj_type") = "html", hide, display) %>" border="0" cellspacing="1" cellpadding="0" align="left">
					<tr><th>PROPRIETA'</th></tr>
					<tr>
						<td class="content">
							<textarea class="codice" id="tft_param_list" rows="15" name="tft_param_list"><%= IIF(cstring(rs("param_list"))<>"", rs("param_list"), ";") %></textarea>
						</td>
					</tr>
				</table>
				<table id="html" style="width:100%; <%= IIF(rs("obj_type") <> "html", hide, display) %>"  border="0" cellspacing="1" cellpadding="0" align="left">
					<tbody style="width:100%;">
						<tr><th>HTML</th></tr>
						<% for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE")) %>
							<tr>
								<td class="content">								
									<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
										<tbody style="width:100%;">
											<tr>
												<td width="5%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
												<td><textarea style="width:100%;" rows="15" name="tft_obj_html_<%= Application("LINGUE")(i)%>"><%= rs("obj_html_" & Application("LINGUE")(i)) %></textarea></td>
											</tr>
										</tbody>
									</table>
								</td>
							</tr>
						<% next %>
					</tbody>
				</table>
			</td>		
		</tr>
		<tr>
			<td class="footer" colspan="3">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="mod" value="SALVA">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>

<%
set rs = nothing
conn.Close
set conn = nothing
%>