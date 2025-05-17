<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ArticoliAccessoriSalva.asp")
end if
%>
<%'--------------------------------------------------------
sezione_testata = "modifica collegamento con altro articolo"%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- %>

<% 
dim conn, rs, rsc, sql, TipiConVincolo
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsc = Server.CreateObject("ADODB.Recordset")

sql = " SELECT * FROM gtb_articoli INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " + _
	  " INNER JOIN grel_art_acc ON gtb_articoli.art_id = grel_art_acc.aa_acc_id " + _
	  " INNER JOIN gtb_accessori_tipo ON grel_art_acc.aa_tipo_id = gtb_accessori_tipo.at_id " + _
	  " WHERE aa_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_aa_art_id" value="<%= rs("aa_art_id") %>">
		<input type="hidden" name="tfn_aa_acc_id" value="<%= rs("aa_acc_id") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Modifica collegamento dell'articolo</caption>
			<tr><th colspan="7">DATI ARTICOLO COLLEGATO</th></tr>
			<% CALL ArticoloScheda (conn, rs, rsc) %>
			<% 
			sql = " SELECT *, (CASE at_vincolo_vendita WHEN 1 THEN at_nome_it + ' (vendita singola vincolabile)' ELSE at_nome_it END) AS NOME " + _
			      " FROM gtb_accessori_tipo " + _
				  " WHERE (at_id NOT IN (SELECT aa_tipo_id FROM grel_Art_acc WHERE aa_art_id=" & rs("aa_art_id") & " AND aa_acc_id=" & rs("aa_acc_id") & " AND aa_tipo_id=gtb_accessori_tipo.at_id) OR at_id=" & rs("aa_tipo_id") & ")" + _
				  " AND (at_id NOT IN (SELECT aa_tipo_id FROM grel_art_acc WHERE aa_acc_id=" & rs("aa_art_id") & " AND aa_art_id=" & rs("aa_acc_id") & " AND aa_tipo_id=gtb_accessori_tipo.at_id) OR at_vincolo_vendita=0) " + _
				  " ORDER BY at_nome_it"
			rsc.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		
			TipiConVincolo = false
			while not rsc.eof 
				if rsc("at_vincolo_vendita") then
					TipiConVincolo = true
				end if
				rsc.movenext
			wend
			rsc.movefirst
			
			if rs("at_vincolo_vendita") AND (TipiConVincolo) then%>
				<tr>
					<td class="label" colspan="2" rowspan="2">non vendibile sing.</td>
					<td class="content" colspan="5">
						<input type="radio" class="checkbox" value="1" name="VendibileSingolarmente" <%= chk(rs("art_NoVenSingola")) %>>
						si
						<input type="radio" class="checkbox" value="" name="VendibileSingolarmente" <%= chk(not rs("art_NoVenSingola")) %>>
						no
					</td>
				</tr>
				<tr>
					<td class="note" colspan="5">
						il prodotto sar&agrave; vendibile SOLO singolarmente se il tipo di collegamento permette la "vendita singola vincolabile".
					</td>
				</tr>
			<% elseif rs("art_NoVenSingola") then %>
				<input type="hidden" name="VendibileSingolarmente" value="1">
			<% end if %>
			<tr>
				<td class="label" colspan="2">tipo di collegamento</td>
				<td class="content" colspan="5">
					<%if rsc.recordcount=1 then%>
						<%= rs("at_nome_it") %> 
						<% if rs("at_vincolo_vendita") then %> (vendita singola vincolabile)<% end if %>
					<%else
						CALL dropDownRecordset(rsc, "at_id", "NOME", "tfn_aa_tipo_id", rs("aa_tipo_id"), true, "", LINGUA_ITALIANO)
					end if
					rsc.close%>
				</td>
			</tr>
			<tr>
				<td class="label" colspan="2">ordine nell'elenco:</td>
				<td class="content" colspan="5">
					<input type="text" class="text" tabindex="1" name="tfn_aa_ordine" value="<%= rs("aa_ordine") %>" maxlength="3" size="3">
				</td>
			</tr>
			<tr><th colspan="7">NOTE / DESCRIZIONE</th></tr>
			<%dim i
			for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
					<td class="content" colspan="7">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>
								<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" alt="" border="0"></td>
								<td><textarea style="width:100%;" rows="2" name="tft_aa_note_<%= Application("LINGUE")(i) %>"><%= rs("aa_note_" & Application("LINGUE")(i)) %></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			<%next %>
			<tr>
				<td class="footer" colspan="7">
					(*) Campi obbligatori.
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
					<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
				</td>
			</tr>
	</form>
		</table>
</div>
</body>
</html>
<% rs.close
conn.close
set rs = nothing
set rsc = nothing
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>