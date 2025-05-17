<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_DocumentiFiles.asp" -->  
<%
if request("salva")<>"" then
	Server.Execute("DocumentoSalva.asp")
end if

dim conn, rs, sql
set conn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.RecordSet")
conn.open Application("DATA_ConnectionString")

'Blocco di codice da copiare in tutte le pagine
'************************************************************************************************************
'Dichiarazione ed impostazione parametri per menu e intestazione
dim Logo_azienda, Titolo_sezione, Sezione, HREF, Action
'Titolo della pagina
if Session("DOC_PRA_ID")<>"" then
	Titolo_sezione = "Pratiche - documenti della pratica - nuovo"
else
	Titolo_sezione = "Documenti - nuovo"
end if
'Indirizzo pagina per link su sezione 
		HREF = "Documenti.asp"
'Azione sul link: {BACK | NEW}
	Action = "INDIETRO"
%>
<!--#INCLUDE FILE="INTESTAZIONE.ASP" -->  

<%'Fine blocco da copiare in tutte le pagine
'************************************************************************************************************
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="tfd_doc_dataC" value="NOW">
	<input type="hidden" name="tfd_doc_mod_Data" value="NOW">
	<input type="hidden" name="tfn_doc_creatore_id" value="<%= Session("ID_ADMIN") %>">
	<input type="hidden" name="tfd_doc_mod_utente" value="<%= Session("ID_ADMIN") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuovo documento</caption>
		<tr><th colspan="4">DATI PRINCIPALI</th></tr>
		<% if Session("DOC_PRA_ID")<>"" then
			CALL SelezionaPratica(conn, rs, "DOC", Session("DOC_PRA_ID"), false) 
		else
			CALL SelezionaPratica(conn, rs, "DOC", request("tfn_doc_pratica_id"), true) 
		end if
		%>
		<tr>
			<td class="label">nome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_doc_nome" value="<%= request("tft_doc_nome") %>" maxlength="255" style="width:70%">
				<span id="nome">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">files:</td>
			<td class="content" colspan="3">
				<% CALL GestioneFilesAssociati(conn, rs, "")%>
			</td>
		</tr>
		<tr><th colspan="4">DESCRIZIONE DOCUMENTO</th></tr>
		<tr>
			<td class="label">tipologia:</td>
			<td class="content" colspan="3">
				<% dropDown conn, "SELECT * FROM tb_tipologie", "tipo_id", "tipo_nome", "tfn_doc_tipologia_id", request("tfn_doc_tipologia_id"), false, "onchange='form1.submit()'", LINGUA_ITALIANO %>
				<span id="tipologia">(*)</span>
			</td>
		</tr>
		<% If request("tfn_doc_tipologia_id") <> "" then %>
			<tr><th class="L2" colspan="4">DESCRITTORI</th></tr>
			<%sql = " SELECT *, (NULL) AS rdd_valore FROM tb_descrittori d INNER JOIN rel_tipologie_descrittori r "& _
			      	" ON d.descr_id=r.rtd_descrittore_id WHERE rtd_tipologia_id="& request("tfn_doc_tipologia_id") & _
				  	" ORDER BY descr_ordine, descr_nome "
			CALL DesElenco(conn, sql, "tb_descrittori", "descr_id", "descr_nome", "descr_tipo", "", "rdd_valore", true, 4)
	   	End If %>
		<tr><th colspan="4">PERMESSI DEL DOCUMENTO</th></tr>
		<tr>
			<td colspan="4"><% CALL AL_disegna(conn, "", AL_DOCUMENTI) %></td>
		</tr>
		<tr><th colspan="4">NOTE</th></tr>
		<tr>
			<td class="content" colspan="4">
				<textarea style="width:100%;" rows="7" name="tft_doc_note"><%=request("tft_doc_note")%></textarea>
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
<%
conn.close
set rs = nothing
set conn = nothing
%>
