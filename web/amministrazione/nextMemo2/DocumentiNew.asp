<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if request("salva")<>"" AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("DocumentiSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura, data
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione documenti - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Documenti.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, i, catalogo_attivo
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")

catalogo_attivo = cBoolean(Session("MEMO2_CATALOGHI_ATTIVI"),false)

%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo documento</caption>
		<tr><th colspan="6">DATI PRINCIPALI</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" style="width:18%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo:</td>
			<% 	end if %>
				<td class="content" colspan="5">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_doc_titolo_<%= Application("LINGUE")(i) %>" value="<%= request("tft_doc_titolo_"& Application("LINGUE")(i)) %>" maxlength="500" size="75">
					<% 	if i = 0 then %>(*)<% end if %>
				</td>
			</tr>
		<% next %>
		<% if cBoolean(Session("CATEGORIE_NEXTMEMO2_ABILITATE"), false) then %>
			<tr>
				<td class="label">categoria:</td>
				<td class="content" colspan="5">
					<%CALL dropDown(conn, categorie.QueryElenco(false, ""), "catC_id", "NAME", "tfn_doc_categoria_id", request("tfn_doc_categoria_id"), false, " onchange='form1.submit()'", LINGUA_ITALIANO)%>
					(*)
				</td>
			</tr>
		<% end If %>
		<tr>
			<td class="label">numero / protocollo:</td>
			<td class="content" colspan="5">
				<%
				dim num
				num = request("tft_doc_numero")
				if cString(num) = "" then
					num = GetValueList(conn, NULL, "SELECT MAX(CAST(doc_numero AS int)) + 1 FROM mtb_documenti WHERE ISNUMERIC(doc_numero) = 1")
				end if
				%>
				<input type="text" class="text" name="tft_doc_numero" value="<%= num %>" maxlength="50" size="30">
			</td>
		</tr>
		<tr>
			<td class="label">data di pubblicazione:</td>
			<td class="label" style="width:4%;text-align:right;">dal:</td>
			<td class="content" style="width:20%;">
				<% if isDate(request("tfd_doc_pubblicazione")) then
					data = request("tfd_doc_pubblicazione")
				else
					data = Date()
				end if
				CALL WriteDataPicker_Input_Manuale("form1", "tfd_doc_pubblicazione", data, "", "/", false, true, LINGUA_ITALIANO, "", true, "") %>
			</td>
			<td class="content" style="width:3%;text-align:left;">
				(*)
			</td>
			<td class="label" style="width:6%;text-align:right;">al:</td>
			<td class="content">
				<% CALL WriteDataPicker_Input_Manuale("form1", "tfd_doc_scadenza", "", "", "/", true, true, LINGUA_ITALIANO, "", true, "") %>
			</td>
		</tr>
		<tr>
			<td class="label" style="width:20%;">visibile:</td>
			<td class="content" colspan="5">
				<input checked type="checkbox" class="noBorder" name="doc_visibile" value="1" <% if request("doc_visibile")<>"" then %> checked <% end if %>>
			</td>
		</tr>
		<% if catalogo_attivo then %>
			<tr>
				<td class="label">tipo visualizzazione:</td>
					<td class="content"><input class="noBorder" type="radio" name="chk_doc_catalogo_sfogliabile" value="" <%=chk(cBoolean(cString(request("chk_doc_catalogo_sfogliabile"))="", true))%>></td>
					<td  class="content" style="padding-left:5px;">standard</td>
					<td class="content"><input class="noBorder" type="radio" name="chk_doc_catalogo_sfogliabile" value="1" <%=chk(cBoolean(cString(request("chk_doc_catalogo_sfogliabile"))="1", false))%> onclick="alert('Attenzione: il salvataggio del catalogo sfogliabile pu&ograve; durare qualche minuto.')"></td>
					<td  class="content" style="padding-left:5px;" colspan="2">catalogo sfogliabile</td>
				</td>
			</tr>	
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>	
				<input type="hidden" name="doc_data_modifica_catalogo_<%= Application("LINGUE")(i) %>" value="">
			<% next %>
		<% end if %>
		
		<tr><th colspan="6">DOCUMENTO</th></tr>
		<tr>
			<td class="label" rowspan="<%= ubound(Application("LINGUE"))+2 %>">file</td>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
					<td class="content"><img src="../grafica/flag_<%= Application("LINGUE")(i)%>.jpg" width="26" height="15" alt="" border="0"></td>
					<td class="content" colspan="4">
						<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_doc_file_" & Application("LINGUE")(i), request("tft_doc_file_" & Application("LINGUE")(i)), "width:450px;", false) %>
					</td>
				</tr>
			<%next %>
		</tr>
		
		
		<% if (cBoolean(Session("CONDIVISIONE_INTERNA"), false) OR cBoolean(Session("CONDIVISIONE_PUBBLICA"), false)) then %>
			<script language="JavaScript" type="text/javascript">
				function show_permessi(){
					var isProtetto = document.getElementById('chk_doc_protetto_true');
					var tab = document.getElementById('protetto')
					if (isProtetto.checked){
						tab.style.visibility = "visible";
						tab.style.display = "block";
					}
					else{
						tab.style.visibility = "hidden";
						tab.style.display = "none";
					}
				}
				
			</script>
		
			<tr><th colspan="6">PERMESSI PER LA VISIBILIT&Agrave; DEL DOCUMENTO</th></tr>
			<tr>
				<td class="label">protetto:</td>
				<td class="content" colspan="5">
					<table border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td><input class="noBorder" type="radio" name="doc_protetto" id="chk_doc_protetto_true" value="1" <%=chk(cBoolean(request("doc_protetto"), false))%> onClick="show_permessi()"></td>
							<td width="30%">si</td>
							<td><input class="noBorder" type="radio" name="doc_protetto" id="chk_doc_protetto_false" value="" <%=chk(cBoolean(request("doc_protetto"), true))%> onClick="show_permessi()"></td>
							<td>no</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td class="content" colspan="6" style="padding:0px;">
					<table id="protetto" border="0" cellspacing="1" cellpadding="0" align="left" style="width:100%;">
						<tr><th colspan="6" class="L2">SCEGLIERE A CHI RENDERE VISIBILE IL DOCUMENTO</th></tr>
						<% sql = "SELECT pro_id FROM mtb_profili"%>
						<% if cString(GetValueList(conn, NULL, sql)) <> "" then %>
							<tr>
								<td class="label" style="width:20%;">profili:</td>
								<td class="content" colspan="4">
									<% dim rsp
									   set rsp = Server.CreateObject("ADODB.RecordSet")
									   sql = "SELECT *, (NULL) as valore FROM mtb_profili ORDER BY pro_nome_it"
									   CALL Write_Relations_Checker(conn, rsp, sql, 3, "pro_id", "pro_nome_it", "valore", "profili_associati")%>
								</td>
							</tr>
						<% end if %>
						<% if cBoolean(Session("CONDIVISIONE_PUBBLICA"), false) then %>
							<tr>
								<td class="label">utenti area riservata:</td>
								<td class="content" colspan="4">
									<% CALL WriteContactPicker_Input(conn, NULL, " ut_ID IN (SELECT rel_ut_id FROM rel_utenti_sito WHERE rel_permesso = 1 OR rel_permesso = 2) ", "", "form1", "utenti_associati", "", "LOGINMANDATORY", true, false, false, "")  %>
								</td>
							</tr>
						<% end if %>
						<% if cBoolean(Session("CONDIVISIONE_INTERNA"), false) then %>
							<tr>
								<td class="label">utenti area amministrativa:</td>
								<td class="content" colspan="4">
									<% CALL WriteAdminPicker_Input(conn, NULL, " ID_admin IN (SELECT admin_id FROM rel_admin_sito WHERE sito_id = 36) ", "form1", "admin_associati", "", "", true, false, false, "")  %>
								</td>
							</tr>
						<% end if %>
					</table>	
				</td>
			</tr>
		<% else %>
			<input type="hidden" value="" name="doc_protetto">
		<% end if %>


		<% sql = " SELECT TOP 1 ct_id FROM mtb_carattech INNER JOIN mrel_categ_ctech ON mtb_carattech.ct_id = mrel_categ_ctech.rcc_ctech_id " & _
			  " WHERE rcc_categoria_id = " & cIntero(request("tfn_doc_categoria_id"))
		if cIntero(GetValueList(conn, NULL, sql)) > 0 then 
			%>
			<tr>
				<td class="content" colspan="6" style="padding:0px;">
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr><th colspan="7">CARATTERISTICHE</th></tr>
						<% if cInteger(request("tfn_doc_categoria_id"))>0 then 
							
							sql = " SELECT *" + _
								  " FROM mtb_carattech"& _
								  " INNER JOIN mrel_categ_ctech ON (mtb_carattech.ct_id = mrel_categ_ctech.rcc_ctech_id AND rcc_categoria_id=" & request("tfn_doc_categoria_id") & ")" + _
								  " LEFT JOIN mrel_doc_ctech ON (mtb_carattech.ct_id = mrel_doc_ctech.rdc_ctech_id AND mrel_doc_ctech.rdc_doc_id=" & CInteger(request("ID")) & ")"& _
								  " LEFT JOIN mtb_carattech_raggruppamenti ON mtb_carattech.ct_raggruppamento_id = mtb_carattech_raggruppamenti.ctr_id " & _
								  " ORDER BY ctr_ordine, rcc_ordine, ct_nome_it, ctr_id"
							CALL DesForm  (conn, sql, "mtb_carattech", "ct_id", "ct_nome_it", "ct_tipo", "ct_unita_it", "", "rdc_valore_", "rdc_valore_", "ctr_titolo_it", cIntero(request("ID")) = 0, 7)
							%>
						<% else %>
							<tr><td class="label" colspan="7">Per descrivere le caratteristiche del documento selezionare prima la sua categoria.</td></tr>
						<% end if %>
					</table>
				</td>
			</tr>
		<% end if %>

		
		<tr><th colspan="6">ESTRATTO / DESCRIZIONE</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="6">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="5" name="tft_doc_estratto_<%= Application("LINGUE")(i)%>"><%= request("tft_doc_estratto_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
					<% 
					CALL activateCKEditor("tft_doc_estratto_"&Application("LINGUE")(i)) 
					%>
				</td>
			</tr>
		<%next %>
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

<% if (cBoolean(Session("CONDIVISIONE_INTERNA"), false) OR cBoolean(Session("CONDIVISIONE_PUBBLICA"), false)) then %>
	<script language="JavaScript" type="text/javascript">
		show_permessi();		
	</script>
<% end if %>

