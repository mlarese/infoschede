<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("OrdiniStatiLavorazioneSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione stati di lavorazione dell'ordine - nuovo"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "OrdiniStatiLavorazione.asp"
dicitura.scrivi_con_sottosez()  

dim conn, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")

%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo stato di lavorazione dell'ordine</caption>
		<tr><th colspan="3">DATI DELLO STATO DI LAVORAZIONE</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
			<% 	end if %>
				<td class="content" colspan="2">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_so_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_so_nome_"& Application("LINGUE")(i)) %>" maxlength="200" size="75">
					<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label" rowspan="4">stato degli ordini collegabili:</td>
			<td class="content<%= STILI_STATI_ORDINE(ORDINE_NON_CONFERMATO) %>" colspan="2">
				<input type="radio" class="checkbox" name="tfn_so_stato_ordini" value="<%= ORDINE_NON_CONFERMATO %>" <%= chk(request("tfn_so_stato_ordini")=ORDINE_NON_CONFERMATO OR request("tfn_so_stato_ordini")="") %>>
				ordini non confermati
			</td>
		</tr>
		<tr>
			<td class="content<%= STILI_STATI_ORDINE(ORDINE_CONFERMATO) %>" colspan="2">
				<input type="radio" class="checkbox" name="tfn_so_stato_ordini" value="<%= ORDINE_CONFERMATO %>" <%= chk(request("tfn_so_stato_ordini")=ORDINE_CONFERMATO) %>>
				ordini confermati
			</td>
		</tr>
		<tr>
			<td class="content<%= STILI_STATI_ORDINE(ORDINE_EVASO) %>" colspan="2">
				<input type="radio" class="checkbox" name="tfn_so_stato_ordini" value="<%= ORDINE_EVASO %>" <%= chk(request("tfn_so_stato_ordini")=ORDINE_EVASO) %>>
				ordini evasi
			</td>
		</tr>
		<tr>
			<td class="content<%= STILI_STATI_ORDINE(ORDINE_ARCHIVIATO) %>" colspan="2">
				<input type="radio" class="checkbox" name="tfn_so_stato_ordini" value="<%= ORDINE_ARCHIVIATO %>" <%= chk(request("tfn_so_stato_ordini")=ORDINE_ARCHIVIATO) %>>
				ordini archiviati
			</td>
		</tr>
		<tr>
			<td class="label">ordine:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tfn_so_ordine" value="<%= request("tfn_so_ordine") %>" maxlength="10" size="3">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">stato di ingresso ordini internet:</td>
			<% sql = "SELECT COUNT(*) FROM gtb_stati_ordine WHERE so_internet=1"
			if cInteger(GetValueList(Conn, NULL, sql))=0 then%>
				<td class="content">
					<input type="checkbox" class="noBorder" checked disabled>
					<input type="hidden" name="chk_so_internet" value="1">
				</td>
				<td class="note">
					Non &egrave; stato trovato lo stato di ingresso degli ordini internet: lo stato che si sta inserendo verr&agrave; impostato come stato iniziale per tali ordini.
				</td>
			<% else %>
				<td class="content">
					<input type="checkbox" class="noBorder" name="chk_so_internet" <%= chk(request("chk_so_internet")<>"") %>>
				</td>
				<td class="note">
					&Eacute; obbligatorio impostare lo stato di arrivo degli ordini in ingresso da internet.
				</td>
			<% end if %>
		</tr>
		<tr><th colspan="3">NOTE</th></tr>
		<tr>
			<td class="content" colspan="3">
				<textarea style="width:100%;" rows="3" name="tft_so_descrizione"><%= request("tft_so_descrizione") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="3">
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
conn.close
set conn = nothing
%>