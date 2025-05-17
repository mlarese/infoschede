<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("CaratteristicheSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim conn, rs, rsc, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = server.CreateObject("ADODB.recordset")
set rsc = server.CreateObject("ADODB.recordset")

if request("goto")<>"" then
	CALL GotoRecord(conn, rsc, session("MEMO2_CTECH_SQL"), "ct_id", "CaratteristicheMod.asp")
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(1)
dicitura.sottosezioni(1) = "GRUPPI"
dicitura.links(1) = "CaratteristicheGruppi.asp"
dicitura.sezione = "Gestione caratteristiche - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Caratteristiche.asp"
dicitura.scrivi_con_sottosez() 


sql = "SELECT * FROM mtb_carattech WHERE ct_id="& cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati della caratteristica</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="caratteristica precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="caratteristica successiva" <%= ACTIVE_STATUS %>>
							SUCCESSIVA &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="2">DATI DELLA CARATTERISTICA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% 	if i = 0 then %>
					<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
				<% 	end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_ct_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("ct_nome_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">tipo di dato:</td>
			<td class="content">
				<% DesDropTipi "tfn_ct_tipo", "", rs("ct_tipo") %>
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% 	if i = 0 then %>
					<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">unit&agrave; di misura:</td>
				<% 	end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_ct_unita_<%= Application("LINGUE")(i) %>" value="<%= rs("ct_unita_"& Application("LINGUE")(i)) %>" maxlength="50" size="50">
				</td>
			</tr>
		<%next %>
		<tr><th colspan="2">GESTIONE DELLA CARATTERISTICA</th></tr>
		<tr>
			<td class="label" >codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_ct_codice" value="<%= rs("ct_codice") %>" maxlength="250" size="26">
			</td>
		</tr>
		<% if cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM mtb_carattech_raggruppamenti"))>0 then %>
			<tr>
				<td class="label">gruppo:</td>
				<td class="content" colspan="3">
					<% sql = "SELECT * FROM mtb_carattech_raggruppamenti ORDER BY ctr_titolo_it"
	                CALL dropDown(conn, sql, "ctr_id", "ctr_titolo_it", "tfn_ct_raggruppamento_id", rs("ct_raggruppamento_id"), false, "", LINGUA_ITALIANO) %>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label">immagine:</td>
			<td class="content" colspan="3">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_ct_img", rs("ct_img") , "width:403px;", false) %>
			</td>
		</tr>
		<tr>
			<td class="label">Confrontabile:</td>
			<td class="content"><input type="checkbox" class="checkbox" name="chk_ct_per_confronto" <%= chk(rs("ct_per_confronto")) %>></td>
		</tr>
		<tr>
			<td class="label">Ricercabile:</td>
			<td class="content"><input type="checkbox" class="checkbox" name="chk_ct_per_ricerca" <%= chk(rs("ct_per_ricerca")) %>></td>
		</tr>
		<tr>
			<td class="label">Principale:</td>
			<td class="content"><input type="checkbox" class="checkbox" name="chk_ct_principale" <%= chk(rs("ct_principale")) %>></td>
		</tr>
		<tr><th colspan="2">CATEGORIE DI PRODOTTI A CUI &Egrave; ASSOCIATA</th></tr>
		<tr>
			<td colspan="2">
				<% dim value
				sql = ", (SELECT rcc_ordine FROM mrel_categ_ctech WHERE rcc_categoria_id = TIP_L0.catC_id AND rcc_ctech_id=" & rs("ct_id") & ") AS ORDINE " + _
						 ", (SELECT COUNT(*) FROM mrel_doc_ctech INNER JOIN mtb_documenti ON mrel_doc_ctech.rdc_doc_id = mtb_documenti.doc_id " + _
						 "	 WHERE mtb_documenti.doc_categoria_id = TIP_L0.catC_id AND mrel_doc_ctech.rdc_ctech_id=" & rs("ct_id") & ") AS N_DOCUMENTI " + _
						 " FROM " %>
				<% sql = replace(categorie.QueryElenco(false, ""), " FROM ", sql) 
				rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<% if rsc.eof then %>
						<tr>
							<td class="label_no_width">
								Nessuna categoria di prodotti definita.
							</td>
						</tr>
					<% else %>
						<script language="JavaScript" type="text/javascript">
							function set_state_all(chk){
								var scrivi = "";
								for(i=0; i<document.form1.elements.length; i++)
								{
									if (document.form1.elements[i].name.indexOf("categorie_associate") >= 0)
									{	
										if (chk.checked)
											document.form1.elements[i].checked = true;
										else
											document.form1.elements[i].checked = false;
									}
									if (document.form1.elements[i].name.indexOf("rel_ordine_") >= 0)
										EnableIfChecked(chk, document.form1.elements[i]);
								}
							}
							
							function set_order_all(){
								var ordine = 0;
								ordine = prompt('Scrivi l\'ordine:','');
								for(i=0; i<document.form1.elements.length; i++)
								{
									if (document.form1.elements[i].name.indexOf("rel_ordine_") >= 0 && document.form1.elements[i].name.indexOf("rel_ordine_all") < 0)
									{
										document.form1.elements[i].value = ordine;
									}
								}
							}
						</script>
						<tr>
							<td colspan="3">
								<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
									<tr>
										<td class="label" style="width:10%;">
											Operazioni:
										</td>
										<td class="content" style="width:25%;">
											<input type="checkbox" class="checkbox" onclick="set_state_all(this)" id="associa_tutte_categorie" title="Sono presenti valori negli articoli di questa categoria.">
											<span>seleziona tutte le categorie</span>
										</td>
										<td class="content">
											<input type="button" disabled class="button_l2_disabled" onclick="set_order_all()" name="rel_ordine_all" id="ordine_tutte_categorie" value="imposta l'ordine per tutte le categorie">
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<th class="l2_center" width="6%">associa</th>
							<th class="l2_center" width="7%">ordine</th>
							<th class="L2">categoria</th>
						</tr>
						<% while not rsc.eof %>
							<tr>
								<td class="content_center">
									<% if cInteger(rsc("N_DOCUMENTI"))>0 then 
										value = true%>
										<input type="checkbox" checked class="checked" id="categorie_associate_<%= rsc("catC_id") %>" disabled onclick="set_state_<%= rsc("catC_id") %>(this)" title="Sono presenti valori nei documenti di questa categoria.">
										<input type="hidden" name="categorie_associate" value=" <%= rsc("catC_id") %> ">
									<% else 
										value = not IsNull(rsc("ORDINE"))%>
										<input type="checkbox" name="categorie_associate" id="categorie_associate_<%= rsc("catC_id") %>" value=" <%= rsc("catC_id") %> " <%= chk(value) %> class="<%= IIF(value, "checked", "checkbox") %>" onclick="set_state_<%= rsc("catC_id") %>(this)">
									<% end if %>
								</td>
								<td class="content_center"><input <%= disable(not value) %> type="text" class="<%= IIF(not value, "text_disabled", "text") %>" size="2" name="rel_ordine_<%= rsc("catC_id") %>" value="<%= rsc("ORDINE") %>"></td>
								<td class="content"><%= rsc("NAME") %></td>
							</tr>
							<script language="JavaScript" type="text/javascript">
								function set_state_<%= rsc("catC_id") %>(chk){
									EnableIfChecked(chk, form1.rel_ordine_<%= rsc("catC_id") %>);
									if (chk.checked){
										form1.rel_ordine_<%= rsc("catC_id") %>.title = "Inserisci l'ordine di visualizzazione nella scheda del documento";
									}
									else{
										form1.rel_ordine_<%= rsc("catC_id") %>.title = "Selezionare il flag di associazione prima di inserire l'ordine di visualizzazione nella scheda documento";
									}
								}
							</script>
							<% rsc.movenext
						wend %>
					<% end if %>
				</table>
				<% rsc.close %>
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
set rs = nothing
set rsc = nothing
conn.Close
set conn = nothing
%>