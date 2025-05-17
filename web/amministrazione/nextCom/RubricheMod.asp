<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<%
'controllo accesso
if Session("COM_ADMIN")="" AND Session("COM_POWER")="" then
	response.redirect "Contatti.asp"
end if

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("RubricheSalva.asp")
end if

dim conn, rs, rsg, sql

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsg = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("SQL_RUBRICHE_ELENCO"), "id_rubrica", "RubricheMod.asp")
end if

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action, i
'Titolo della pagina
	Titolo_sezione = "Rubriche - modifica"
'Indirizzo pagina per link su sezione 
		HREF = "Rubriche.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************

sql = "SELECT * FROM tb_rubriche WHERE id_rubrica=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati della rubrica</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="rubrica precedente">
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="rubrica successiva">
							SUCCESSIVA &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="2">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label" style="width:22%;">nome rubrica:</td>
			<td class="content">
				<input type="text" class="text" name="tft_nome_rubrica" value="<%= rs("nome_rubrica") %>" maxlength="250" size="75">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<% 	if IsAdminCurrent(conn) then %>
		<tr>
			<td class="label">rubrica utilizzata dal sistema:</td>
			<td class="content">
				<input type="checkbox" class="noborder" name="chk_locked_rubrica" <%if rs("locked_rubrica") then%>checked <% end if %>>
			</td>
		</tr>
		<tr>
			<td class="label">rubrica esterna:</td>
			<td class="content">
				<input type="checkbox" class="noborder" name="chk_rubrica_esterna" <%if rs("rubrica_esterna") then%>checked <% end if %>>
			</td>
		</tr>
		<% 	else %>
		<input type="hidden" name="chk_rubrica_esterna" value="<%= IIF(rs("rubrica_esterna"), "1", "") %>">
		<input type="hidden" name="chk_locked_rubrica" value="<%= IIF(rs("locked_rubrica"), "1", "") %>">
		<% 		if rs("rubrica_esterna") then %>
		<tr>
			<td class="label" colspan="2">
				La cancellazione e l'inserimento di questa rubrica vengono gestiti automaticamente dal sistema. La rubrica &egrave; collegata
				ad una applicazione esterna che ne gestisce automaticamente le propriet&agrave;.
			</td>
		</tr>
		<% 		end if
			end if %>
		<% if Cinteger(Application("NextCom_DefaultWorkGroup"))=0 then %>
			<% 
			sql = "SELECT tb_gruppi.id_gruppo, tb_gruppi.nome_Gruppo, tb_rel_gruppiRubriche.id_rel_grupprub " &_
			  				 " FROM tb_gruppi LEFT JOIN tb_rel_gruppiRubriche ON (tb_gruppi.id_gruppo = tb_rel_gruppiRubriche.id_Gruppo_assegnato " &_
			  				 " AND tb_rel_gruppiRubriche.id_dellaRubrica=" & cIntero(request("ID")) & ")" & _
			  				 " ORDER BY nome_Gruppo"
			dim rs_group
			set rs_group = Server.CreateObject("ADODB.RecordSet")
			rs_group.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText 
			
			if rs_group.recordcount > 1 then %>
				<tr><th colspan="2">GRUPPI DI LAVORO COMPETENTI (*)</th></tr>
				<tr>
					<td colspan="2">
						<% 
						CALL Write_Relations_Checker(conn, rsg, sql, 2, "id_gruppo", "nome_Gruppo", "id_rel_grupprub", "gruppi") %>
					</td>
				</tr>
			<% else %>
				<input type="hidden" name="gruppi" value="<%= rs_group("id_gruppo") %>"
			<% end if 
			rs_group.close 
			set rs_group = nothing
			%>
			
		<% else %>
			<input type="hidden" name="gruppi" value="<%= Application("NextCom_DefaultWorkGroup") %>">
		<% end if %>
		<tr><th colspan="2">CONTATTI ASSOCIATI</th></tr>
		<tr>
			<td class="label">singoli contatti:</td>
			<td class="content">
                <% CALL WriteContactPicker_Input(conn, rsg, "", "", "form1", "contatti", "SELECT id_indirizzo FROM rel_rub_ind WHERE id_rubrica=" & cIntero(request("ID")), "", true, false, false, "") %>
			</td>
		</tr>
		<tr><th colspan="2">DATI VISIBILI NELLA PARTE PUBBLICA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" style="width:18%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome pubblico:</td>
			<% 	end if %>
				<td class="content" colspan="5">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_nome_pubblico_rubrica_<%= Application("LINGUE")(i) %>" value="<%= rs("nome_pubblico_rubrica_"& Application("LINGUE")(i)) %>" maxlength="500" size="75">
				</td>
			</tr>
		<% next %>
		<tr><th colspan="2">NOTE</th></tr>
		<tr>
			<td class="content" colspan="2">
				<textarea style="width:100%;" rows="5" name="tft_Note_rubrica"><%=rs("Note_rubrica")%></textarea>
			</td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<% if Session("COM_ADMIN")<>"" then %>
			<tr><th colspan="2">STRUMENTI</th></tr>
			<tr>
				<td class="content">
					<a HREF="javascript:void(0);" onClick="OpenAutoPositionedScrollWindow('RubricheFusione.asp?RUBRICA_SORGENTE=<%= rs("id_rubrica") %>', 'fondi_rubrica', 500, 250, true)" class="button_L2_block"
					   title="apri lo strumento di fusione di due rubriche: permette di copiare l'associazione dei contatti nella rubrica di destinazione." <%= ACTIVE_STATUS %>>
						FUSIONE RUBRICHE
					</a>
				</td>
				<td class="label_no_width">apri lo strumento di fusione di due rubriche: permette di copiare l'associazione dei contatti nella rubrica di destinazione.</td>
			</tr>
		<% end if %>
		<tr>
			<td class="footer" colspan="2">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva_avanti" value="SALVA">
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
set rsg = nothing
set rs = nothing
set conn = nothing
%>