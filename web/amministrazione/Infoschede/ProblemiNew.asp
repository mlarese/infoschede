<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if request("salva")<>"" AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ProblemiSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura, data
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione guasti - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Problemi.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, i, ordine
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")

%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo guasto</caption>
		<tr><th colspan="6">DATI PRINCIPALI</th></tr>
		<script language="JavaScript" type="text/javascript">
			function natura_guasto(){
				var isProtetto = document.getElementById('chk_prb_riscontrato_true');
				var tab = document.getElementById('protetto')
				if (!(isProtetto.checked)){
					tab.style.visibility = "visible";
					tab.style.display = "block";
				}
				else{
					tab.style.visibility = "hidden";
					tab.style.display = "none";
				}
			}
		</script>
		<tr>
			<td class="label">guasto:</td>
			<td class="content" colspan="5">
				<table border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td><input class="noBorder riscontrato" type="radio" name="prb_riscontrato" id="chk_prb_riscontrato_true" value="1" <%=chk(cBoolean(request("prb_riscontrato"), true))%> onClick="natura_guasto()"></td>
						<td width="60%" style="padding-left:5px;">riscontrato</td>
						<td><input class="noBorder segnalato" type="radio" name="prb_riscontrato" id="chk_prb_riscontrato_false" value="" <%=chk(cBoolean(request("prb_riscontrato"), false))%> onClick="natura_guasto()"></td>
						<td width="30%" style="padding-left:5px;">segnalato</td>
					</tr>
				</table>
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" style="width:18%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo:</td>
			<% 	end if %>
				<td class="content" colspan="5">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_prb_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_prb_nome_"& Application("LINGUE")(i)) %>" maxlength="255" size="100">
					<% 	if i = 0 then %>(*)<% end if %>
				</td>
			</tr>
		<% next %>
		<tr>
			<td class="label">ordine:</td>
			<td class="content" colspan="5">
				<% 
				if request("tfn_prb_ordine") <> "" then
					ordine = request("tfn_prb_ordine")
				else
					sql = " SELECT MAX(prb_ordine) FROM sgtb_problemi "
					ordine = cIntero(GetValueList(conn,NULL,sql)) + 10
				end if
				%>
				<input type="text" class="text" name="tfn_prb_ordine" value="<%= ordine %>" maxlength="8" size="8">
			</td>
		</tr>
		<tr>
			<td class="label" style="width:20%;">visibile:</td>
			<td class="content" colspan="5">
				<input checked type="checkbox" class="noBorder" name="prb_visibile" value="1" <% if request("prb_visibile")<>"" then %> checked <% end if %>>
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="2">immagini</td>
			<td class="label_no_width">thumb:</td>
			<td class="content" colspan="4">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_prb_img_th", request("tft_prb_img_th"), "width:350px;", false) %>
			</td>
		</tr>
		<tr>
			<td class="label_no_width">zoom:</td>
			<td class="content" colspan="4">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_prb_img_zo", request("tft_prb_img_zo"), "width:350px;", false) %>
			</td>
		</tr>
		
		
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i=lbound(Application("LINGUE")) then %>
					<td class="label" rowspan="<%=ubound(Application("LINGUE")) + 1%>">descrizione:</td>
				<% end if %>
				<td class="content" colspan="5">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="5%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="3" name="tft_prb_descrizione_<%= Application("LINGUE")(i)%>"><%= request("tft_prb_descrizione_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		<%next %>
		
		<tr><th colspan="6">ASSOCIAZIONI</th></tr>
		<tr>
			<td class="label" colspan="6">Le associazioni saranno possibili dopo aver registrato il guasto.<br>&nbsp;</td>
		</tr>

		<script language="JavaScript" type="text/javascript">
			function abilita(){
				var modalitaEasy = document.getElementById('chk_prb_modalita_easy_true');
				<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					var tab = document.getElementById('tft_prb_avviso_per_conferma_<%= Application("LINGUE")(i)%>')
					DisableControl(tab, modalitaEasy.checked)
				<%next%>
			}
			
		</script>
		
		<tr>
			<td class="content" colspan="6" style="padding:0px;">
				<table id="protetto" border="0" cellspacing="1" cellpadding="0" align="left" style="width:100%;">
					<tr><th colspan="6">CARATTERISTICHE DEL GUASTO</th></tr>
					<tr>
						<td class="label" style="width:18%;">modalit&agrave;:</td>
						<td class="content" colspan="5">
							<table border="0" cellspacing="0" cellpadding="0" align="left">
								<tr>
									<td><input class="noBorder" type="radio" name="prb_modalita_easy" id="chk_prb_modalita_easy_true" value="1" <%=chk(cBoolean(request("prb_modalita_easy"), true))%> onClick="abilita()"></td>
									<td width="60%"><% WriteColor(MODALITA_EASY_COLOR)%><%=MODALITA_EASY%></td>
									<td><input class="noBorder" type="radio" name="prb_modalita_easy" id="chk_prb_modalita_easy_false" value="" <%=chk(cBoolean(request("prb_modalita_easy"), false))%> onClick="abilita()"></td>
									<td width="30%"><% WriteColor(MODALITA_NON_EASY_COLOR)%><%=MODALITA_NON_EASY%></td>
								</tr>
							</table>
						</td>
					</tr>
					
					<tr><th class="L2" colspan="6">AVVISO PER CONFERMA</th></tr>
					<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
						<tr>
							<td class="content" colspan="6">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
									<tr>
										<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
										<td><textarea style="width:100%;" rows="10" id="tft_prb_avviso_per_conferma_<%= Application("LINGUE")(i)%>" name="tft_prb_avviso_per_conferma_<%= Application("LINGUE")(i)%>"><%= request("tft_prb_avviso_per_conferma_" & Application("LINGUE")(i)) %></textarea></td>
									</tr>
								</table>
							</td>
						</tr>
					<%next %>
				</table>
			</td>
		</tr>
			
		<tr>
			<td class="footer" colspan="6">
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

<script language="JavaScript" type="text/javascript">
	abilita();
	natura_guasto();
</script>


