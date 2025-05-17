<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_menu_accesso, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoMenuSalva.asp")
end if
 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - menu - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoMenu.asp"
dicitura.scrivi_con_sottosez() 

dim i, lingua
dim conn, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfn_m_id_webs" value="<%=Session("AZ_ID")%>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" border="0">
		<caption>Inserimento nuovo menu</caption>
		<tr><th colspan="4">DATI DEL MENU</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
			lingua = Application("LINGUE")(i)
			if Session("LINGUA_" & lingua) then%>
				<tr>
				<% 	if i = 0 then %>
					<td class="label" style="width:10%;" colspan="2" rowspan="<%= Session("LINGUE_ATTIVE") %>">nome:</td>
				<% 	end if %>
					<td class="content" colspan="2">
						<img src="../grafica/flag_<%= lingua %>.jpg" width="26" height="15" alt="" border="0">
						<input type="text" class="text" name="tft_m_nome_<%= lingua %>" value="<%= request("tft_m_nome_"& lingua) %>" maxlength="255" size="100">
						<% 	if lingua = LINGUA_ITALIANO then response.write "(*)" end if %>
					</td>
				</tr>
			<%end if
		next %>
		<tr><th colspan="4">LINKS DEL MENU</th></tr>
		<script type="text/javascript">
			function AbilitaMaster(v) {
				var o
				o = document.getElementById("rad_copia_0")
				DisableControl(o, (v != "0"))
				o = document.getElementById("rad_copia_1")
				DisableControl(o, (v != "0"))
				
				if (v != "0") {
					o = document.getElementById("copia_menu")
					DisableControl(o, true)
					o = document.getElementById("copia_index")
					DisablePicker(o, true)
					o = document.getElementById("chk_mi_figli")
					DisableControl(o, true)
				} else {
					document.getElementById("rad_copia_0").checked = false
					document.getElementById("rad_copia_1").checked = false
				}
				
				o = document.getElementById("tfn_m_index_id")
				DisablePicker(o, (v != "1"))
			}
			
			function Abilita(v) {
				var o
				o = document.getElementById("copia_menu")
				DisableControl(o, (v == "1"))
				
				o = document.getElementById("copia_index")
				DisablePicker(o, (v == "0"))
				o = document.getElementById("chk_mi_figli")
				DisableControl(o, (v == "0"))
			}
		</script>
		<tr>
			<td class="content">
				<input type="radio" class="noBorder" name="rad_op" onclick="AbilitaMaster(this.value)" value="" <%= chk(request("rad_op")="") %>>
			</td>
			<td class="label" colspan="3">
				inserimento manuale dei link
			</td>
		</tr>
		<tr>
			<td class="content" rowspan="5">
				<input type="radio" class="noBorder" name="rad_op" onclick="AbilitaMaster(this.value)" value="0" <%= chk(request("rad_op")="0") %>>
			</td>
			<td class="label" rowspan="5">
				copia da:
			</td>
			<td class="content_center" rowspan="2">
				<input type="radio" disabled class="noBorder" name="rad_copia" id="rad_copia_0" onclick="Abilita(this.value)" value="0" <%= chk(request("rad_copia")="0") %>>
			</td>
			<td class="content">altro menu</td>
		</tr>
		<tr>
			<td class="content">
				<% sql = "SELECT m_id, m_nome_it FROM tb_menu WHERE m_id_webs=" & session("AZ_ID") & " ORDER BY m_nome_it"
				if cString(GetValueList(conn, NULL, sql)) <> "" then
					CALL dropDown(conn, sql, "m_id", "m_nome_it", "copia_menu", "", FALSE, "disabled", LINGUA_ITALIANO)
				else
					response.write "&nbsp;"					
				end if
				%>
			</td>
		</tr>
		<tr>
			<td class="content_center" rowspan="3">
				<input type="radio" disabled class="noBorder" name="rad_copia" id="rad_copia_1" onclick="Abilita(this.value)" value="1" <%= chk(request("rad_copia")="1") %>>
			</td>
			<td class="content">voce dell'indice</td>
		</tr>
		<tr>
			<td class="content">
				<% 	CALL index.WritePicker("", "", "form1", "copia_index", "", Session("AZ_ID"), false, false, "86", true, false) %>
			</td>
		</tr>
		<tr>
			<td class="content">
				<input disabled type="checkbox" name="chk_mi_figli" id="chk_mi_figli" value="1" class="checkbox">
				abilita visualizzazione delle voci figlie
			</td>
		</tr>
		
		<tr>
			<td class="content" rowspan="2">
				<input type="radio" class="noBorder" name="rad_op" onclick="AbilitaMaster(this.value)" value="1"<%= chk(request("rad_op")="1") %>>
			</td>
			<td class="label" rowspan="2">aggancia a:</td>
			<td class="content" colspan="2">voce dell'indice</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<% 	CALL index.WritePicker("", "", "form1", "tfn_m_index_id", "", Session("AZ_ID"), false, false, "90", true, false) %>
			</td>
		</tr>
		
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA &gt;&gt;">
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