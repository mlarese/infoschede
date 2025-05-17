<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_Pubblicazioni_accesso, 0))

if request.form("salva_elenco")<>"" then
	Server.Execute("IndexPubblicazioniSalva.asp")
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Indice generale - Pubblicazioni automatiche dei dati - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "IndexPubblicazioni.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rst, sql, i, lingua
set conn = Server.CreateObject("ADODB.Connection")
set rst = Server.CreateObject("ADODB.Recordset")
conn.open Application("DATA_ConnectionString")
%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" border="0">
		<caption>Inserimento nuova pubblicazione automatica dei dati</caption>
		<tr><th colspan="4">PUBBLICAZIONE DEI DATI</th></tr>
		<tr>
			<td class="label">titolo</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_pub_titolo" value="<%= request("tft_pub_titolo") %>" maxlength="255" size="79">
				(*)
			</td>
		</tr>
		<tr><th colspan="4">SORGENTE DEI DATI</th></tr>
		<tr>
			<td class="label">tabella</td>
			<td class="content" colspan="3">
				<% sql = " SELECT tab_id, (sito_nome + ' - ' + tab_titolo) AS NAME FROM tb_siti_tabelle INNER JOIN tb_siti ON tb_siti_tabelle.tab_sito_id = tb_siti.id_sito " + _
						 " ORDER BY sito_nome, tab_titolo "
				CALL dropDown(conn, sql, "tab_id", "NAME", "tfn_pub_tabella_id", request("tfn_pub_tabella_id"), false, " onchange=""form1.submit()""", LINGUA_ITALIANO)%>
				(*)
			</td>
		</tr>
		<tr>
			<td class="label">filtro sql sui dati</td>
			<td class="content" colspan="3">
				<textarea name="tft_pub_filtro_pubblicazione" class="codice" rows="3" cols="60"><%= request("tft_pub_filtro_pubblicazione") %></textarea>
			</td>
		</tr>
		<tr><th colspan="4">INFORMAZIONI DI PUBBLICAZIONE</th></tr>
		<tr>
			<td class="label" nowrap>posizione nell'indice</td>
			<td class="content" colspan="3">
				<% CALL Index.WritePicker("", "", "form1", "tfn_pub_padre_index_id", request("tfn_pub_padre_index_id"), 0, false, false, 90, false, true) %>
			</td>
		</tr>
		<tr>
			<td class="label" nowrap>campo "principale"</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_pub_field_principale" value="<%= request("tft_pub_field_principale") %>" maxlength="255" size="79">
			</td>
		</tr>
		<tr><th class="L2" colspan="4">CATEGORIZZAZIONE DEI DATI</th></tr>
		<tr>
			<td class="label">tabella</td>
			<td class="content" colspan="3">
				<% sql = " SELECT tab_id, (sito_nome + ' - ' + tab_titolo) AS NAME FROM tb_siti_tabelle INNER JOIN tb_siti ON tb_siti_tabelle.tab_sito_id = tb_siti.id_sito " + _
						 " ORDER BY sito_nome, tab_titolo "
				CALL dropDown(conn, sql, "tab_id", "NAME", "tfn_pub_categoria_tabella_id", request("tfn_pub_categoria_tabella_id"), false, "", LINGUA_ITALIANO)%>
				(**)
			</td>
		</tr>
		<tr>
			<td class="label">campo "categoria"</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_pub_categoria_field" value="<%= request("tft_pub_categoria_field") %>" maxlength="255" size="79">
				(**)
			</td>
		</tr>
		<tr><th class="L2" colspan="4">INFORMAZIONI PER CALCOLO URL DELLA VOCE DELL'INDICE</th></tr>
		<% if cInteger(request("tfn_pub_tabella_id"))=0 then %>
			<tr>
				<td class="content_b" colspan="4">Scegliere la tabella da pubblicare automaticamente.</td>
			</tr>
		<% else
			'verifica la tabella: se &egrave; impostato il campo url non permette la selezione della pagina 
			sql = "SELECT * FROM tb_siti_tabelle WHERE tab_id=" & cIntero(request("tfn_pub_tabella_id"))
			rst.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			i = ""
			for each lingua in Application("LINGUE")
 				if rst("tab_field_url_" + lingua) <> "" then
					i = rst("tab_field_url_" + lingua)
					exit for
				end if
			next
			if i <> "" then%>
				<input type="hidden" name="tipo_url" value="esterno">
				<tr>
					<td class="label" rowspan="2">url dell'elemento:</td>
					<td class="content_b" colspan="3">
						gestito automaticamente dal contenuto
					</td>
				</tr>
				<tr>
					<td class="label">campo sorgente</td>
					<td class="content_b" colspan="2"><%= i %></td>
				</tr>
			<% else %>
				<input type="hidden" name="tipo_url" value="">
				<tr>
					<td class="label" rowspan="2">url dell'elemento:</td>
					<td class="label">pagina collegata</td>
					<td class="content" colspan="2">
						<% CALL DropDownPages(conn, "form1", "420", 0, "tfn_pub_pagina_id", request("tfn_pub_pagina_id"), true, false) %>
						(*)
					</td>
				</tr>
				<tr>
					<td class="label">parametro nell'url</td>
					<td class="content_b" colspan="2"><%= rst("tab_parametro") %></td>
				</tr>
			<% end if
			rst.close
		end if %>
		<tr><th colspan="4">PERSONALIZZAZIONE DEI CONTENUTI</th></tr>
		<tr>
			<td class="label" style="width:23%;" rowspan="<%= ubound(Application("LINGUE")) + 1 %>">titolo:</td>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
				if i>1 then%>
					<tr>
				<% end if %>
					<td class="content" colspan="3">
						<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
						<input type="text" class="text" name="tft_pub_field_titolo_<%= Application("LINGUE")(i) %>" value="<%= request("tft_pub_field_titolo_" & Application("LINGUE")(i)) %>" maxlength="250" size="75">
					</td>
				</tr>
			<%next %>
		</tr>
		<tr>
			<td class="label" style="width:23%;" rowspan="<%= ubound(Application("LINGUE")) + 1 %>">titolo alternativo:</td>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
				if i>1 then%>
					<tr>
				<% end if %>
					<td class="content" colspan="3">
						<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
						<input type="text" class="text" name="tft_pub_field_titolo_alt_<%= Application("LINGUE")(i) %>" value="<%= request("tft_pub_field_titolo_alt_" & Application("LINGUE")(i)) %>" maxlength="250" size="75">
					</td>
				</tr>
			<%next %>
		</tr>
		<tr>
			<td class="label" style="width:23%;" rowspan="<%= ubound(Application("LINGUE")) + 1 %>">descrizione:</td>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))
				if i>1 then%>
					<tr>
				<% end if %>
					<td class="content" colspan="3">
						<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
						<input type="text" class="text" name="tft_pub_field_descrizione_<%= Application("LINGUE")(i) %>" value="<%= request("tft_pub_field_descrizione_" & Application("LINGUE")(i)) %>" maxlength="250" size="75">
					</td>
				</tr>
			<%next %>
		</tr>
		<tr>
			<td class="label">foto thumb:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_pub_field_foto_thumb" value="<%= request("tft_pub_field_foto_thumb") %>" maxlength="250" size="75">				
			</td>
		</tr>
		<tr>
			<td class="label">foto zoom:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_pub_field_foto_zoom" value="<%= request("tft_pub_field_foto_zoom") %>" maxlength="250" size="75">				
			</td>
		</tr>
		<tr>
			<td class="label" colspan="4">meta-keywords</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="4">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="6" name="tft_pub_field_meta_keywords_<%= Application("LINGUE")(i) %>"><%= request("tft_pub_field_meta_keywords_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		<%next %>
		<tr>
			<td class="label" colspan="4">meta-description</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<td class="content" colspan="4">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="6" name="tft_pub_field_meta_description_<%= Application("LINGUE")(i) %>"><%= request("tft_pub_field_meta_description_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		<%next %>		
		<tr>
			<td class="footer" colspan="4">
                <span style="float:left; text-align:left;">
                    (*) campi obbligatori.<br>
                    (**) permette l'individuazione del record "padre" all'interno del ramo di base scelto.
                </span>
				<input type="submit" class="button" name="salva_elenco" value="SALVA"<% if cInteger(request("tfn_pub_tabella_id"))=0 then %> disabled title="per salvare la pubblicazione &egrave; necessario scegliere la tabella." <% end if %>>
			</td>
		</tr>
	</table>
	</form>
</div>
</body>
</html>

<%
conn.Close
set rst = nothing
set conn = nothing
%>