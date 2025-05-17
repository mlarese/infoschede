<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->

<%
dim conn, rs, ID, tipo
set conn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.Recordset")
conn.open Application("DATA_ConnectionString")

ID = cIntero(request.querystring("ID"))
tipo = ParseSQL(request.querystring("tipo"), adChar)

'controllo accesso
dim ko
SELECT CASE tipo
	CASE AL_DEFAULT
		ko = (NOT AL(conn, cIntero(request("ID")), AL_PRATICHE) OR Session("COM_POWER") = "" AND _
		     Session("ID_ADMIN") <> CInt(GetValueList(conn, rs, "SELECT pra_creatore_id FROM tb_pratiche WHERE pra_id="& cIntero(request("ID"))))) _
		     AND Session("COM_ADMIN") = ""
	CASE AL_ATTIVITA
		dim aux
		set aux = conn.execute("SELECT att_inSospeso, att_mittente_id FROM tb_attivita WHERE att_id="& ID)
		ko = NOT AL(conn, cIntero(request("ID")), AL_ATTIVITA) OR _
			 Session("COM_ADMIN") = "" AND NOT (aux("att_inSospeso") AND aux("att_mittente_id") = Session("ID_ADMIN"))
		set aux = nothing
	CASE AL_DOCUMENTI
		ko = (NOT AL(conn, cIntero(request("ID")), AL_DOCUMENTI) OR Session("COM_POWER") = "" AND _
		     Session("ID_ADMIN") <> CInt(GetValueList(conn, rs, "SELECT doc_creatore_id FROM tb_documenti WHERE doc_id="& cIntero(request("ID"))))) _
	    	 AND Session("COM_ADMIN") = ""
	CASE ELSE
		ko = TRUE
END SELECT
if ko then%>
	<script language="javascript">
		alert('Modifica dei <%= IIF(tipo=AL_ATTIVITA, "destinatari", "permessi") %> non consentita');
		this.close()
	</script>
	<%response.end
else

'GESTIONE INSERIMENTO AL
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	if request.form("ere") <> "" then
		CALL AL_ins(conn, tipo, ID, true)
	else
		CALL AL_ins(conn, tipo, ID, false)
	end if
%>
<script language="javascript">
<% If request.querystring("ctrl") <> "" then %>
	opener.location.reload()
<% Else  %>
	<% If tipo = AL_DEFAULT then %>
	opener.location.href = "Pratiche.asp?ID=<%= Session("COM_PRA_CLIENTE") %>"
	<% Else  %>
	opener.location.href = "<%= tipo %>.asp"
	<% End If %>
<% End If %>
	this.close()
</script>
<% End If %>

<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<%'--------------------------------------------------------
sezione_testata = "Impostazione / modifica access list" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>

<div id="content_ridotto">
<table cellspacing="1" cellpadding="0" class="tabella_madre">
	<% 	If tipo = AL_DOCUMENTI AND CInteger(GetValueList(conn, rs, "SELECT COUNT(*) FROM tb_allegati a INNER JOIN tb_attivita t ON a.all_attivita_id=t.att_id WHERE NOT "& SQL_IsTrue(conn, "att_conclusa") &" AND all_documento_id="& ID)) > 0 then %>
		<form action="" method="post" id="form1" name="form1" onsubmit="alert('Attenzione: il documento &egrave; allegato ad una attivit&agrave; aperta e potrebbe non essere pi&ugrave; visibile')">
	<% 	Else %>
		<form action="" method="post" id="form1" name="form1">
	<% 	End If %>
		<caption style="border-bottom-width:1px;">
			<% if tipo <> AL_ATTIVITA then%>
				Gestione permessi
			<% else %>
				Impostazione lista destinatari
			<% end if %>
		</caption>
		<tr>
			<td>
				<% CALL AL_disegna(conn, ID, tipo) %>
			</td>
		</tr>
		<tr>
			<td class="footer">
				(*) Campi obbligatori.
				<input type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close()">
				<input type="submit" class="button" name="salva" value="SALVA">
			</td>
		</tr>
	</form>
	</table>
</div>
</body>
</html>
<% 
end if			'fine controllo permessi
conn.close 
set rs = nothing
set conn = nothing
%>
