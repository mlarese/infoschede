<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ListeCodiciSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione liste codici - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "ListeCodici.asp"
dicitura.scrivi_con_sottosez() 

dim conn
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova lista di codici</caption>
		<tr><th colspan="3">DATI DELLA LISTA DI CODICI</th></tr>
		<tr>
			<td class="label">Codice lista:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_LstCod_cod" value="<%= request("tft_LstCod_cod") %>" maxlength="50" size="40">
			</td>
		</tr>
		<tr>
			<td class="label">Nome:</td>
			<td class="content" colspan="2">
				<input type="text" class="text" name="tft_LstCod_nome" value="<%= request("tft_LstCod_nome") %>" maxlength="50" size="75">
				(*)
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="5">copia codici da:</td>
			<td class="content" colspan="2">
				<input type="radio" class="checkbox" name="copia_da" id="copia_da_v" value="" <%= chk(request("copia_da")="") %> onclick="EnableIfChecked(document.all.copia_da_L, form1.copia_da_lista);">
				codici vuoti
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="radio" class="checkbox" name="copia_da" id="copia_da_i" value="i" <%= chk(request("copia_da")="i") %> onclick="EnableIfChecked(document.all.copia_da_L, form1.copia_da_lista);">
				codici interni
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="radio" class="checkbox" name="copia_da" id="copia_da_a" value="a" <%= chk(request("copia_da")="a") %> onclick="EnableIfChecked(document.all.copia_da_L, form1.copia_da_lista);">
				codici alternativi
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input type="radio" class="checkbox" name="copia_da" id="copia_da_p" value="p" <%= chk(request("copia_da")="p") %> onclick="EnableIfChecked(document.all.copia_da_L, form1.copia_da_lista);">
				codici produttore
			</td>
		</tr>
		<tr>
			<td class="content" width="16%">
				<input type="radio" class="checkbox" name="copia_da" id="copia_da_L" value="l" <%= chk(request("copia_da")="l") %> onclick="EnableIfChecked(document.all.copia_da_L, form1.copia_da_lista);">
				altra lista codici:
			</td>
			<td class="content">
				<%CALL dropDown(conn, "SELECT lstCod_id, lstCod_nome FROM gtb_Lista_codici ORDER BY lstCod_nome", _
							    "lstCod_id", "lstCod_nome", "copia_da_lista", request("copia_da_lista"), true, "", LINGUA_ITALIANO)%>
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="2">tipo di lista:</td>
			<td class="content" colspan="2">
				<input class="checkbox" type="radio" name="tfn_lstCod_sistema" value="0" <%= chk(request("tfn_lstCod_sistema")="0" OR request("tfn_lstCod_sistema")="") %>> dei clienti
			</td>
		</tr>
		<tr>
			<td class="content" colspan="2">
				<input class="checkbox" type="radio" name="tfn_lstCod_sistema" value="1" <%= chk(request("tfn_lstCod_sistema")="1") %>> di sistema
			</td>
		</tr>
		<tr><th colspan="3">NOTE</th></tr>
		<tr>
			<td class="content" colspan="3">
				<textarea style="width:100%;" rows="3" name="tft_lstCod_note"><%= request("tft_lstCod_note") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="3">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva_avanti" value="SALVA &gt;&gt;">
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

<script language="JavaScript" type="text/javascript">
	EnableIfChecked(document.all.copia_da_L, form1.copia_da_lista);
</script>