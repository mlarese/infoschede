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
dicitura.sezione = "Gestione siti - plugin - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoPlugin.asp"
dicitura.scrivi_con_sottosez() 

dim conn, sql,i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
%>
<style type="text/css">
/*
	#content table tr
	{
		display:block;
		width:100%;
	}
	#content table,
	#content table tr th,
	#content table tr td
	{
		width:100%;
	}
	*/
</style>
<script language="JavaScript" type="text/javascript">
	function tipo_plugin()
	{		
		var SelectedIndex = form1.tft_obj_type.selectedIndex;
		var html = document.getElementById('html')
		var plugin = document.getElementById('plugin')		
		var sorgente = document.getElementById('sorgente')
		if (SelectedIndex==2){			
			html.style.visibility = "visible";
			html.style.display = "block";
			plugin.style.visibility = "hidden";
			plugin.style.display = "none";
			sorgente.style.visibility = "hidden";
			sorgente.style.display = "none";
			form1.sorg.value="NextHTMLController";
			form1.tft_param_list.value=";";
		}
		else{			
			html.style.visibility = "hidden";
			html.style.display = "none";
			if (plugin.style.visibility == "hidden")
			{
				plugin.style.visibility = "visible";
				plugin.style.display = "block";
			}
			if (sorgente.style.visibility == "hidden")
			{
				sorgente.style.visibility = "visible";
				sorgente.style.display = "block";
			}
		}
	}
</script>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_id_webs" value="<%= Session("AZ_ID") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" border="0">
		<caption>Inserimento nuovo plugin</caption>
		<tr><th colspan="2">DATI DEL PLUGIN</th></tr>
		<tr>
			<td class="label">nome:</td>
			<td class="content">
				<input type="text" class="text" name="tft_name_objects" value="<%= request("tft_name_objects") %>" maxlength="255" size="80">
				(*)
			</td>
		</tr>
		<tr id="sorgente">
			<td class="label">classe sorgente:</td>
			<td class="content">
				<input type="text" id="sorg" class="text" name="tft_identif_objects" value="<%= request("tft_identif_objects") %>" maxlength="70" size="80">
			</td>
		</tr>
		<tr>
			<td class="label">tipo:</td>
			<td class="content">
				<select name="tft_obj_type" onchange="tipo_plugin()">
					<option value="ascx" <%= IIF(request.form("tft_obj_type") = "ascx", "selected", "") %>>Plugin</option>
					<option value="class" <%= IIF(request.form("tft_obj_type") = "class", "selected", "") %>>Classe</option>
					<option value="html" <%= IIF(request.form("tft_obj_type") = "html", "selected", "") %>>Html</option>
				</select>
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2" style="padding:0px;">
				<table id="plugin" border="0" cellspacing="1" cellpadding="0" align="left" style="width:100%;">
					<tbody style="width:100%;">
						<tr><th colspan="2">PROPRIETA'</th></tr>
						<tr>
							<td class="content" colspan="2">
								<textarea class="codice" id="tft_param_list" rows="15" name="tft_param_list"><%= request("tft_param_list") %></textarea>
							</td>
						</tr>
					</tbody>
				</table>
				<table id="html" style="width:100%; visibility:hidden; display:none" border="0" cellspacing="1" cellpadding="0" align="left">
					<tbody style="width:100%;">
						<tr><th colspan="2">HTML</th></tr>
						<% for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE")) %>
							<tr>
								<td class="content" colspan="2">								
									<table border="0" cellspacing="0" cellpadding="0" align="left" style="width:100%;">
										<tr>
											<td style="width:5%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
											<td><textarea style="width:100%;" rows="15" name="tft_obj_html_<%= Application("LINGUE")(i)%>"><%= request("tft_html_" & Application("LINGUE")(i)) %></textarea></td>
										</tr>
									</table>
								</td>
							</tr>
						<% next %>
					</tbody>
				</table>
			</td>		
		</tr>
		<tr>
			<td class="footer" colspan="2">
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
set conn = nothing%>