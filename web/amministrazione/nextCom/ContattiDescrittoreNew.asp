<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_DocumentiFiles.asp" -->
<%

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ContattiDescrittoreSalva.asp")
end if

dim conn, rs, sql, i

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")


'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
	Titolo_sezione = "Descrittori contatti - nuovo"
'Indirizzo pagina per link su sezione 
		HREF = "ContattiDescrittori.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo descrittore</caption>
		<tr><th colspan="2">DATI PRINCIPALI</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i = 0 then %>
					<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
				<% end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_ict_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_ict_nome_"& Application("LINGUE")(i))%>" maxlength="50" size="50">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label">tipo:</td>
			<td class="content">
				<% CALL DesDropTipi("tfn_ict_tipo", "", request.form("tfn_ict_tipo")) %>
				<span id="nome">(*)</span>
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% 	if i = 0 then %>
					<td class="label" style="width:20%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">unit&agrave; di misura:</td>
				<% 	end if %>
				<td class="content">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_ict_unita_<%= Application("LINGUE")(i) %>" value="<%= request("tft_ict_unita_"& Application("LINGUE")(i)) %>" maxlength="50" size="20">
				</td>
			</tr>
		<%next %>
		<tr><th colspan="2">GESTIONE DELLA CARATTERISTICA</th></tr>
		<tr>
			<td class="label">codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_ict_codice" value="<%= request("tft_ict_codice")%>" maxlength="50" size="20">
			</td>
		</tr>
		<% if cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM tb_indirizzario_carattech_raggruppamenti"))>0 then %>
			<tr>
				<td class="label">gruppo:</td>
				<td class="content">
					<% sql = "SELECT * FROM tb_indirizzario_carattech_raggruppamenti ORDER BY icr_titolo_it"
	                CALL dropDown(conn, sql, "icr_id", "icr_titolo_it", "tfn_ict_raggruppamento_id", request.form("tfn_ict_raggruppamento_id"), false, "", LINGUA_ITALIANO) %>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label">immagine:</td>
			<td class="content">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_ict_img", request.form("tft_ict_img"), "width:430px;", FALSE) %>
			</td>
		</tr>
		<tr>
			<td class="label">confrontabile:</td>
			<td class="content"><input type="checkbox" class="checkbox" name="chk_ict_per_confronto" <%= chk(request("chk_ict_per_confronto")<>"") %>></td>
		</tr>
		<tr>
			<td class="label">ricercabile:</td>
			<td class="content"><input type="checkbox" class="checkbox" name="chk_ict_per_ricerca" <%= chk(request("chk_ict_per_ricerca")<>"") %>></td>
		</tr>
		<tr><th colspan="2">CATEGORIE DI PRODOTTI A CUI &Egrave; ASSOCIATA</th></tr>
		<tr>
			<td colspan="2">
				<%
				dim value
				sql = ", (NULL) AS ORDINE " + _
						 ", (0) AS N_CONTATTI " + _
						 " FROM "%>
				<% sql = replace(CatContatti.QueryElenco(false, ""), " FROM ", sql) 
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<% if rs.eof then %>
						<tr>
							<td class="label_no_width">
								Nessuna categoria di contatti &egrave; stata trovata.
							</td>
						</tr>
					<% else %>
						<tr>
							<th class="l2_center" width="6%">associa</th>
							<th class="l2_center" width="7%">ordine</th>
							<th class="L2">categoria</th>
						</tr>
						<% while not rs.eof %>
							<tr>
								<td class="content_center">
									<% if cInteger(rs("N_CONTATTI"))>0 then 
										value = true%>
										<input type="checkbox" checked class="checked" id="categorie_associate_<%= rs("icat_id") %>" disabled onclick="set_state_<%= rs("icat_id") %>(this)" title="Sono presenti valori nei contatti di questa categoria.">
										<input type="hidden" name="categorie_associate" value=" <%= rs("icat_id") %> ">
									<% else 
										value = instr(1, request("categorie_associate"), " " & rs("icat_id") & " ", vbTextCompare)>0 %>
										<input type="checkbox" name="categorie_associate" id="categorie_associate_<%= rs("icat_id") %>" value=" <%= rs("icat_id") %> " <%= chk(value) %> class="<%= IIF(value, "checked", "checkbox") %>" onclick="set_state_<%= rs("icat_id") %>(this)">
									<% end if %>
								</td>
								<td class="content_center"><input <%= disable(not value) %> type="text" class="<%= IIF(not value, "text_disabled", "text") %>" size="2" name="rel_ordine_<%= rs("icat_id") %>" value="<%= request("rel_ordine_" & rs("icat_id")) %>"></td>
								<td class="content"><%= rs("NAME") %></td>
							</tr>
							<script language="JavaScript" type="text/javascript">
								function set_state_<%= rs("icat_id") %>(chk){
									EnableIfChecked(chk, form1.rel_ordine_<%= rs("icat_id") %>);
									if (chk.checked){
										form1.rel_ordine_<%= rs("icat_id") %>.title = "Inserisci l'ordine di visualizzazione nella scheda del documento";
									}
									else{
										form1.rel_ordine_<%= rs("icat_id") %>.title = "Selezionare il flag di associazione prima di inserire l'ordine di visualizzazione nella scheda documento";
									}
								}
							</script>
							<% rs.movenext
						wend %>
					<% end if %>
				</table>
				<% rs.close %>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="2">
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
set rs = nothing
set conn = nothing%>