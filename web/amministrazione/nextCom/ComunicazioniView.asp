<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->
<%
dim conn, rs, rsd, rse, rsr, sql, allegato, FromWizard
dim LocalTitle

if request("FromSend")<>"" then
	'passo conclusivo di riepilogo del wizard di spedizione delle email
	FromWizard = request("FromSend")
else
	'visualizzazione normale
	FromWizard = ""
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.RecordSet")
set rse = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Replace(Session("SQL_COMUNICAZIONI"),"*","email_id"), "email_id", "ComunicazioniView.asp")
end if

sql = "SELECT tb_email.*, " &_
	  "(SELECT admin_cognome + ' ' + admin_nome FROM tb_admin WHERE id_admin=email_dipgenera) AS da " &_
	  "FROM tb_email WHERE email_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdtext

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
if FromWizard<>"" then
	'titolo della pagina
	Titolo_sezione = lcase(ComunicazioniNew_Wizard_Titolo("Comunicazioni in uscita - ", 4, FromWizard))
	'Azione sul link: {BACK | NEW}
		Action = "FINE"
else
	'Titolo della pagina
	Titolo_sezione = "Comunicazioni in uscita - Visualizzazione " + Comunicazioni_LabelByType(rs("email_tipi_messaggi_id"), "email inviata", "fax inviato", "sms inviato")
	
	'Azione sul link: {BACK | NEW}
		Action = "INDIETRO"
end if
'Indirizzo pagina per link su sezione 
HREF = "Comunicazioni.asp"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
'recupera errori di spedizione
sql = GetQuery_LogContatti(conn, rs("email_id"), true)
rse.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
'recupera rubriche
sql =  GetQuery_LogRubriche(conn, rs("email_id"))
rsr.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
'recupera destinatari
sql = GetQuery_LogContatti(conn, rs("email_id"), false)
rsd.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<% if FromWizard<>"" then %>
			<caption class="<%= IIF(rse.recordcount>0, "alert", "ok")%>">
				<% CALL Comunicazioni_Icona(rs("email_tipi_messaggi_id")) %>
				<%= ComunicazioniNew_Wizard_Titolo("", 4, FromWizard) %> - report di spedizione
			</caption>
		<% else %>
			<caption>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td class="caption">
							<% CALL Comunicazioni_Icona(rs("email_tipi_messaggi_id")) %>
							<%if rs("email_archiviata") then %>
								&nbsp;
								<img src="../grafica/archiviata.gif" border="0" alt="Comunicazione archiviata">
							<% end if %>
						</td>
						<td class="caption" title="<%=cString(rs("email_control_key"))%>">Visualizzazione <%= Comunicazioni_LabelByType(rs("email_tipi_messaggi_id"), "email inviata", "fax inviato", "sms inviato") %></td>
						<td align="right" style="font-size: 1px;">
							<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="comunicazione precedente">
								&lt;&lt; PRECEDENTE
							</a>
							&nbsp;
							<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="comunicazione successiva">
								SUCCESSIVA &gt;&gt;
							</a>
						</td>
					</tr>
				</table>
			</caption>
		<% end if %>
		<tr><th colspan="3">MITTENTE</th></tr>
		<tr>
			<td class="label" style="width:15%;">da:</td>
			<td class="content" colspan="2"><%= rs("da") %></td>
		</tr>
		<% if rs("email_tipi_messaggi_id") <> MSG_SMS then %>
			<script language="JavaScript" type="text/javascript">
				function SaveAsNewsletter(){
					var width, height;
					width = 300;
					height = 250;
					OpenAutoPositionedScrollWindow('ComunicazioniNew_Wizard_2_newsletter_save_as.asp?EMAIL_ID=<%= rs("email_id") %>&TIPO_ID=<%= cIntero(rs("email_newsletter_tipo_id")) %>', 
													'saveemailasnewsletter', width, height, true);

				}
			</script>
			<tr><th colspan="3">INTESTAZIONE</th></tr>
			<tr>
				<td class="label">oggetto:</td>
				<td class="content"><%= rs("email_object") %></td>
				<td class="label" align="right" rowspan="2">
					<span style="white-space:nowrap;">
						<% if cIntero(rs("email_newsletter_tipo_id")) = 0 then %>
							Salva l'email come newsletter:
						<% else %>
							Tipo newsletter:
							<%= GetValueList(conn, NULL, "SELECT nl_nome_it FROM tb_newsletters WHERE nl_id = " & cIntero(rs("email_newsletter_tipo_id")))%>
						<% end if %>
					</span>
					<br>
					<a class="button_L2" href="javascript:void(0);" onclick="SaveAsNewsletter()">
						<% if cIntero(rs("email_newsletter_tipo_id")) = 0 then %>
							&nbsp;SCEGLI TIPO NEWSLETTER&nbsp;		
						<% else %>
							&nbsp;CAMBIA TIPO NEWSLETTER&nbsp;		
						<% end if %>
					</a>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="label">inviat<%= IIF(rs("email_tipi_messaggi_id") = MSG_EMAIL, "a", "o") %>:</td>
			<td class="content"><%= DateTimeIta(rs("email_data")) %></td>
		</tr>
		<tr><th colspan="3">DESTINATARI</th></tr>
		<tr>
			<td colspan="3">
				<% CALL Write_LogCompleto_Destinatari(rs, rse, rsr, rsd, false) %>
			</td>
		</tr>	
		<tr><th colspan="3">CORPO DEL MESSAGGIO</th></tr>
		<tr>
			<td colspan="3">
				<table width="100%" cellSpacing="1" cellPadding="0">
					<tr>
						<% CALL Write_MessageViewFrame(rs("email_tipi_messaggi_id"), "ComunicazioniViewMessage.asp?ID=" & rs("email_id")) %>
					</tr>
				</table>
			</td>
		</tr>
		<% if rs("email_tipi_messaggi_id") = MSG_EMAIL then %>
			<tr><th colspan="3">ALLEGATI</th></tr>
			<tr><td class="content" colspan="3"><% CALL Write_Allegati(rs("email_docs"), rs("email_id")) %></td></tr>
		<% end if %>
		<tr>
			<td class="footer" colspan="3">
				<a href="Comunicazioni.asp" class="button">
					<%= IIF(FromWizard<>"", "FINE", "INDIETRO") %>
				</a>
			</td>
		</tr>
	</table>
	&nbsp;
	
</div>
</body>
</html>
<% 
rse.close
rsr.close
rsd.close

rs.close
conn.close 
set rs = nothing
set rsd = nothing
set conn = nothing
%>