<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_DocumentiFiles.asp" -->
<%

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ContattiDescrittoreSalva.asp")
end if

dim conn, rs, rsc, sql, i

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rsc, session("NEXTCOM_CONTATTI_CTECH_SQL"), "ict_id", "ContattiDescrittoreMod.asp")
end if

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = "Descrittori contatti - modifica"
'Indirizzo pagina per link su sezione 
		HREF = "ContattiDescrittori.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************

sql = "SELECT * FROM tb_indirizzario_carattech WHERE ict_id = " & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
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
					<input type="text" class="text" name="tft_ict_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("ict_nome_"& Application("LINGUE")(i)) %>" maxlength="255" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">tipo di dato:</td>
			<td class="content">
				<% DesDropTipi "tfn_ict_tipo", "", rs("ict_tipo") %>
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% 	if i = 0 then %>
					<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">unit&agrave; di misura:</td>
				<% 	end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_ict_unita_<%= Application("LINGUE")(i) %>" value="<%= rs("ict_unita_"& Application("LINGUE")(i)) %>" maxlength="50" size="50">
				</td>
			</tr>
		<%next %>
		<tr><th colspan="2">GESTIONE DELLA CARATTERISTICA</th></tr>
		<tr>
			<td class="label" >codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_ict_codice" value="<%= rs("ict_codice") %>" maxlength="250" size="26">
			</td>
		</tr>
		<% if cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM tb_indirizzario_carattech_raggruppamenti"))>0 then %>
			<tr>
				<td class="label">gruppo:</td>
				<td class="content" colspan="3">
					<% sql = "SELECT * FROM tb_indirizzario_carattech_raggruppamenti ORDER BY icr_titolo_it"
	                CALL dropDown(conn, sql, "icr_id", "icr_titolo_it", "tfn_ict_raggruppamento_id", rs("ict_raggruppamento_id"), false, "", LINGUA_ITALIANO) %>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label">immagine:</td>
			<td class="content" colspan="3">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_ict_img", rs("ict_img") , "width:403px;", false) %>
			</td>
		</tr>
		<tr>
			<td class="label">Confrontabile:</td>
			<td class="content"><input type="checkbox" class="checkbox" name="chk_ict_per_confronto" <%= chk(rs("ict_per_confronto")) %>></td>
		</tr>
		<tr>
			<td class="label">Ricercabile:</td>
			<td class="content"><input type="checkbox" class="checkbox" name="chk_ict_per_ricerca" <%= chk(rs("ict_per_ricerca")) %>></td>
		</tr>
		<tr><th colspan="2">CATEGORIE DI PRODOTTI A CUI &Egrave; ASSOCIATA</th></tr>
		<tr>
			<td colspan="2">
				<% dim value
				sql = ", (SELECT rcc_ordine FROM rel_categ_ctech WHERE rcc_categoria_id = TIP_L0.icat_id AND rcc_ctech_id=" & rs("ict_id") & ") AS ORDINE " + _
						 ", (SELECT COUNT(*) FROM rel_cnt_ctech INNER JOIN tb_indirizzario ON rel_cnt_ctech.ric_cnt_id = tb_indirizzario.IDElencoIndirizzi " + _
						 "	 WHERE tb_indirizzario.cnt_categoria_id = TIP_L0.icat_id AND rel_cnt_ctech.ric_ctech_id=" & rs("ict_id") & ") AS N_CONTATTI " + _
						 " FROM " %>
				<% sql = replace(CatContatti.QueryElenco(false, ""), " FROM ", sql) 
				rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<% if rsc.eof then %>
						<tr>
							<td class="label_no_width">
								Nessuna categoria di prodotti definita.
							</td>
						</tr>
					<% else %>
						<tr>
							<th class="l2_center" width="6%">associa</th>
							<th class="l2_center" width="7%">ordine</th>
							<th class="L2">categoria</th>
						</tr>
						<% while not rsc.eof %>
							<tr>
								<td class="content_center">
									<% if cInteger(rsc("N_CONTATTI"))>0 then 
										value = true%>
										<input type="checkbox" checked class="checked" id="categorie_associate_<%= rsc("icat_id") %>" disabled onclick="set_state_<%= rsc("icat_id") %>(this)" title="Sono presenti valori nei contatti di questa categoria.">
										<input type="hidden" name="categorie_associate" value=" <%= rsc("icat_id") %> ">
									<% else 
										value = not IsNull(rsc("ORDINE"))%>
										<input type="checkbox" name="categorie_associate" id="categorie_associate_<%= rsc("icat_id") %>" value="<%= rsc("icat_id") %>" <%= chk(value) %> class="<%= IIF(value, "checked", "checkbox") %>" onclick="set_state_<%= rsc("icat_id") %>(this)">
									<% end if %>
								</td>
								<td class="content_center"><input <%= disable(not value) %> type="text" class="<%= IIF(not value, "text_disabled", "text") %>" size="2" name="rel_ordine_<%= rsc("icat_id") %>" value="<%= rsc("ORDINE") %>"></td>
								<td class="content"><%= rsc("NAME") %></td>
							</tr>
							<script language="JavaScript" type="text/javascript">
								function set_state_<%= rsc("icat_id") %>(chk){
									EnableIfChecked(chk, form1.rel_ordine_<%= rsc("icat_id") %>);
									if (chk.checked){
										form1.rel_ordine_<%= rsc("icat_id") %>.title = "Inserisci l'ordine di visualizzazione nella scheda del documento";
									}
									else{
										form1.rel_ordine_<%= rsc("icat_id") %>.title = "Selezionare il flag di associazione prima di inserire l'ordine di visualizzazione nella scheda documento";
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
rs.close
conn.close 
set rsc = nothing
set rs = nothing
set conn = nothing%>