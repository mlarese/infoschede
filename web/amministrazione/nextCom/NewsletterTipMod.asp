<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->

<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("NewsletterTipSalva.asp")
end if

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Titolo_sezione, Sezione, HREF, Action
Titolo_sezione = "Tipologie Newsletter - modifica"
HREF = "NewsletterTip.asp"
Action = "INDIETRO;"

%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************

dim conn, rs, rsa, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "SELECT * FROM tb_newsletters WHERE nl_id = " & request("ID")
rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 

%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova tipologia di newsletter</caption>
		<tr><th colspan="4">DATI DELLA TIPOLOGIA</th></tr>
		<tr>
			<td class="label" style="width:20%;">nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_nl_nome_it" value="<%= rs("nl_nome_it")%>" maxlength="250" size="79">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">pagina da spedire:</td>
			<td class="content" colspan="2">
				<% CALL DropDownPages(conn, "form1", "410", 0, "tfn_nl_pagina_id", rs("nl_pagina_id"), true, false) %>
				<span style="line-height:24px;">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">gestione dinamica contenuti:</td>
			<td class="content">
				<input type="radio" class="checkbox" value="1" name="tfn_nl_gestione_dinamica_contenuti" <%= chk(rs("nl_gestione_dinamica_contenuti")) %>>
				si
				<input type="radio" class="checkbox" value="0" name="tfn_nl_gestione_dinamica_contenuti" <%= chk(not rs("nl_gestione_dinamica_contenuti")) %>>
				no
			</td>
		</tr>
		<tr>
			<td class="label_no_width" style="width:17%;">lingua:</td>
			<td class="content" colspan="2"><% CALL DropLingue(conn, rsa, "tft_nl_lingua", rs("nl_lingua"), true, false, "") %></td>
		</tr>
		<tr><th colspan="4">Scegli i destinatari di default</th></tr>
		<tr>
			<td class="label">rubriche:</td>
			<input type="hidden" name="tft_nl_rubriche_default" value="<%= rs("nl_rubriche_default") %>">
			<input type="hidden" name="contatti_email_newsletter_true" value="">
			<input type="hidden" name="invia_anche_contatti_interni" value="">
			<td class="content" colspan="3">
				<table border="0" cellspacing="0" cellpadding="0" style="width:100%">
					<tr>
						<td style="width:95%;" colspan="2">
							<% 
							dim messageType, rubricheNameList, contattiIdList 
							contattiIdList = rs("nl_contatti_default")
							messageType = MSG_EMAIL
							
							if cString(rs("nl_rubriche_default")) = "" then
								sql = "0"
							else
								sql = Replace(rs("nl_rubriche_default"), ";", ",")
								sql = Trim(sql)
								sql = left(sql, len(sql) - 1)
							end if
							sql = "SELECT * FROM tb_rubriche WHERE id_rubrica IN (" & sql & ") "
							rs.close
							rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
							while not rs.eof
								rubricheNameList = rubricheNameList & " " & JSReplacerEncode(rs("nome_Rubrica")) &";"
								rs.moveNext
							wend
							rs.close
							%>
							<textarea READONLY style="width:100%;" rows="3" name="visRubriche"><%= rubricheNameList %></textarea>
						</td>
						<td style="width:60px; vertical-align: top; padding-top: 1px;">
							<a class="button_textarea"
								href="javascript:void(0)" onclick="OpenAutoPositionedScrollWindow('ComunicazioniNew_Wizard_2_Rubriche.asp?input_hid=tft_nl_rubriche_default&messageType=<%= messageType %>&page_No=1&elenco='+form1.tft_nl_rubriche_default.value, 'selezione_rubriche', 600, 400, false);">
								SCEGLI
							</a>
							<a class="button_textarea" 
								href="javascript:void(0)" onclick="form1.visRubriche.value='';form1.tft_nl_rubriche_default.value=''">
								RESET
							</a>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td class="label">singoli contatti:</td>
			<td class="content" colspan="3">
				<% 
							
				CALL WriteContactPicker_Input(conn, rs, "", "", "form1", "tft_nl_contatti_default", contattiIdList, "EMAILMANDATORY;CNTREL", true, false, false, "") %>
			</td>
		</tr>
		
		<tr>
			<td class="footer" colspan="4">
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
