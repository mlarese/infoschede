<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<%
'controllo accesso
if Session("COM_ADMIN")="" AND Session("COM_POWER")="" then
	response.redirect "Campagne.asp"
end if

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("CampagneSalva.asp")
end if

dim conn, rs, rsc, rsi, sql

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")
set rsi = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("SQL_CAMPAGNE_ELENCO"), "inc_id", "CampagneMod.asp")
end if

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action, i
'Titolo della pagina
	Titolo_sezione = "Campagne marketing - modifica"
'Indirizzo pagina per link su sezione 
		HREF = "Campagne.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************

sql = "SELECT * FROM tb_indirizzario_campagne WHERE inc_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfd_inc_modData" value="<%=Now()%>">
	<input type="hidden" name="tfn_inc_modAdmin_id" value="<%=Session("ID_ADMIN")%>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati della campagna marketing</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="campagna precedente">
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="campagna successiva">
							SUCCESSIVA &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="2">DATI PRINCIPALI</th></tr>
		<tr>
			<td class="label" style="width:22%;">nome campagna:</td>
			<td class="content">
				<input type="text" class="text" name="tft_inc_nome" value="<%= rs("inc_nome") %>" maxlength="250" size="70">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<tr><th colspan="2">CONTATTI ASSOCIATI</th></tr>
		<tr>
			<td class="label">contatti per i quali la campagna &egrave; stata effettuata:</td>
			<td class="content">
                <% dim id_contatti_nascosti
				sql = " SELECT rcc_cnt_id FROM rel_cnt_campagne WHERE rcc_data_conclusione IS NOT NULL " & _
						 " AND rcc_campagna_id=" & cIntero(request("ID"))
				rsc.open sql, conn
				while not rsc.eof
					sql = "SELECT * FROM tb_indirizzario WHERE IDElencoIndirizzi = " & rsc("rcc_cnt_id")
					rsi.open sql, conn
					response.write ContactFullName(rsi) & "; "
					id_contatti_nascosti = id_contatti_nascosti & rsi("IDElencoIndirizzi") & ";"
					rsi.close
					rsc.moveNext
				wend
				rsc.close
				%>
				<!--<input type="hidden" name="contatti_nascosti" value="<%=id_contatti_nascosti%>">-->
			</td>
		</tr>
		
		<tr>
			<td class="label">singoli contatti:</td>
			<td class="content">
                <% sql = " SELECT rcc_cnt_id FROM rel_cnt_campagne WHERE rcc_data_conclusione IS NULL " & _
						 " AND rcc_campagna_id=" & cIntero(request("ID"))
				CALL WriteContactPicker_Input(conn, rsc, "", "", "form1", "contatti", sql, "EMAIL", true, false, false, "") %>
			</td>
		</tr>
		<tr><th colspan="2">NOTE</th></tr>
		<tr>
			<td class="content" colspan="2">
				<textarea style="width:100%;" rows="5" name="tft_inc_note"><%=rs("inc_note")%></textarea>
			</td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
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
set rsi = nothing
set rsc = nothing
set rs = nothing
set conn = nothing
%>