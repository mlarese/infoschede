<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ArticoliVariantiSalva.asp")
end if

'--------------------------------------------------------
sezione_testata = "modifica della variante" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>
<SCRIPT LANGUAGE="javascript"  src="Tools_B2B.js" type="text/javascript"></SCRIPT>

<% 
dim conn, rs, rsa, rsv, sql, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsv = Server.CreateObject("ADODB.Recordset")
set rsa = Server.CreateObject("ADODB.Recordset")

sql = " SELECT *, " + _
	  " (SELECT COUNT(*) FROM gtb_bundle WHERE gtb_bundle.bun_articolo_id=gv_articoli.rel_id) AS COMPONENTE, " +_
	  " (SELECT COUNT(*) FROM gtb_dettagli_ord WHERE det_art_var_id= gv_articoli.rel_id) AS ORDINI " + _
	  " FROM gv_articoli WHERE rel_id=" & cIntero(request("ID"))
rs.open sql, conn, adopenStatic, adLockReadOnly, adCmdText
%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova variante per l'articolo</caption>
		<tr><th colspan="6">DATI VARIANTE</th></tr>
		<%sql = " SELECT val_nome_it, var_nome_it FROM grel_art_vv " + _
		  		" INNER JOIN gtb_valori ON grel_art_vv.rvv_val_id = gtb_valori.val_id " + _
		  		" INNER JOIN gtb_varianti ON gtb_valori.val_var_id = gtb_varianti.var_id " + _
		  		" WHERE grel_art_vv.rvv_art_var_id=" & cIntero(request("ID")) & _
		  		" ORDER BY var_ordine, val_ordine"
		rsv.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		while not rsv.eof%>
			<tr>
				<td class="label"><%= rsv("var_nome_it") %></td>
				<td colspan="5" class="content"><%= rsv("val_nome_it") %></td>
			</tr>
			<% rsv.movenext
		wend 
		rsv.close
		%>
		<tr>
			<td class="label" rowspan="3">codici</td>
			<td class="label" style="width:31%;">interno:</td>
		<%if cInteger(rs("COMPONENTE")) = 0 AND cInteger(rs("ORDINI")) = 0 then%>
				<input type="hidden" class="text" name="codice_modificabile" value="1">
				<td class="content" colspan="4">
					<input type="text" class="text" name="tft_rel_cod_int" value="<%= rs("rel_cod_int")%>" maxlength="50" size="20">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label">alternativo:</td>
				<td class="content" colspan="4">
					<input type="text" class="text" name="tft_rel_cod_alt" value="<%= rs("rel_cod_alt")%>" maxlength="50" size="20">
				</td>
			</tr>
			<tr>
				<td class="label">produttore:</td>
				<td class="content" colspan="4">
					<input type="text" class="text" name="tft_rel_cod_pro" value="<%= rs("rel_cod_pro")%>" maxlength="50" size="20">
				</td>
			</tr>
		<% else %>
				<td class="content"><%= rs("rel_cod_int") %>&nbsp;</td>
				<td class="note" rowspan="3" colspan="3">Variante gi&agrave; utilizzata nella composizione di un bundle e/o di un ordine.</td>
			</tr>
			<tr>
				<td class="label">alternativo:</td>
				<td class="content" colspan="5"><%= rs("rel_cod_alt") %>&nbsp;</td>
			</tr>
			<tr>
				<td class="label">produttore:</td>
				<td class="content" colspan="5"><%= rs("rel_cod_pro") %></td>
			</tr>
		<% end if %>
			<tr><th colspan="6">PREZZO VARIANTE</th></tr>
			<tr>
				<td class="label" rowspan="5">prezzo:</td>
				<td class="label">calcola prezzo variante:</td>
				<td class="content" colspan="4">
					<% if rs("rel_prezzo_indipendente") then %>
						da prezzo articolo sulla base delle variazioni applicate
					<% else %>
						indipendente dal prezzo articolo
					<% end if %>
				</td>
			</tr>
			<tr>
				<td class="label">della variante</td>
				<td class="content" width="20%"><%= FormatPrice(rs("rel_prezzo") , 2, true) %>&nbsp;&euro;</td>
				<td class="note" rowspan="4" colspan="4">
					La gestione dei prezzi &egrave; disponibile dalla sezione 
					<a class="button_L2" target="_blank" href="ArticoliPrezzi.asp?ID=<%= rs("art_id") %>" title="Apre la gestione dei prezzi dell'articolo in una nuova finestra" <%= ACTIVE_STATUS %>>PREZZI</a> 
					dell'articolo.
				</td>
			</tr>
			<tr>
				<td class="label" rowspan="2">
					variazioni da prezzo articolo:
				</td>
				<td class="content" style="padding-left:19px" colspan="4"><%= FormatPrice(cReal(rs("rel_var_sconto")) , 2, false) %> %</td>
			</tr>
			<tr>
				<td class="content" style="padding-left:19px" colspan="6"><%= FormatPrice(cReal(rs("rel_var_euro")) , 2, false) %> &euro;</td>
			</tr>
			<tr>
				<td class="label">classe di sconto per quantit&agrave;:</td>
				<td class="content" colspan="5">
					<% if cInteger(rs("rel_scontoQ_id"))>0 then 
						sql = "SELECT scc_nome FROM gtb_scontiQ_classi WHERE scc_id=" & rs("rel_scontoQ_id") & " ORDER BY scc_nome"%>
						<%= GetValueList(conn, rsv, sql) %>
					<% end if %>
				</td>
			</tr>
			<tr>
				<th class="L2" colspan="6">specifiche aggiuntive sul prezzo:</th>
			</tr>
			<% for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
					<td class="content" colspan="6">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>
								<td width="6%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<td><textarea style="width:100%;" rows="2" name="tft_rel_descr_prezzo_<%= Application("LINGUE")(i) %>"><%= rs("rel_descr_prezzo_" & Application("LINGUE")(i)) %></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			<% next %>
			<tr><th colspan="6">DATI DI GESTIONE</th></tr>
			<tr>
				<td class="label" colspan="2">non a catalogo:</td>
				<td class="content" colspan="5"><input type="checkbox" class="checkbox" name="chk_rel_disabilitato" <%= chk(rs("rel_disabilitato")) %>></td>
			</tr>
			<tr>
				<td class="label" rowspan="3">gestione:</td>
				<td class="label">giacenza minima</td>
				<td class="content" colspan="4"><input type="text" class="number" name="tfn_rel_giacenza_min" value="<%= rs("rel_giacenza_min")%>" size="4"></td>
			</tr>
			<tr>
				<td class="label">quantit&agrave; minima ordinabile</td>
				<td class="content" colspan="5"><input type="text" class="number" name="tfn_rel_qta_min_ord" value="<%= rs("rel_qta_min_ord")%>" size="4"></td>
			</tr>
			<tr>
				<td class="label">lotto di riordino</td>
				<td class="content" colspan="5"><input type="text" class="number" name="tfn_rel_lotto_riordino" value="<%= rs("rel_lotto_riordino")%>" size="4"></td>
			</tr>
 			<tr>
				<td class="label" colspan="2">foto:</td>
				<td class="content" colspan="4">
				<% 	dropDown conn, "SELECT * FROM gtb_art_foto WHERE fo_articolo_id=" & rs("rel_art_id") & " ORDER BY fo_ordine", "fo_id", "fo_thumb", "nfn_rel_foto_id", rs("rel_foto_id"), false, "", LINGUA_ITALIANO %>
				</td>
			</tr>
			<% if cBoolean(cString(Session("ATTIVA_GESTIONE_COLLI-PESI-DIMENSIONI")), false) then %>
				<tr><th colspan="6">COLLI, PESO E VOLUME</th></tr>
				<tr>
					<td class="label" rowspan="2">colli:</td>
					<td class="label">numero colli</td>
					<td class="content" colspan="4"><input type="text" class="number" name="tfn_rel_colli_num" value="<%= rs("rel_colli_num") %>" size="4"></td>
				</tr>
				<tr>
					<td class="label">numero pezzi per collo</td>
					<td class="content" colspan="4"><input type="text" class="number" name="tfn_rel_collo_pezzi_per" value="<%= rs("rel_collo_pezzi_per") %>" size="4"></td>
				</tr>
				<tr>
					<td class="label" rowspan="2">peso:</td>
					<td class="label">peso netto</td>
					<td class="content" colspan="4"><input type="text" class="number" name="tfn_rel_peso_netto" value="<%= rs("rel_peso_netto") %>" size="4"> Kg</td>
				</tr>
				<tr>
					<td class="label">peso lordo</td>
					<td class="content" colspan="4"><input type="text" class="number" name="tfn_rel_peso_lordo" value="<%= rs("rel_peso_lordo") %>" size="4"> Kg</td>
				</tr>
				<tr>
					<td class="label" rowspan="4">volume:</td>
					<td class="label">larghezza collo</td>
					<td class="content" colspan="4"><input type="text" class="number" name="tfn_rel_collo_width" value="<%= rs("rel_collo_width") %>" size="4"> m</td>
				</tr>
				<tr>
					<td class="label">altezza collo</td>
					<td class="content" colspan="4"><input type="text" class="number" name="tfn_rel_collo_height" value="<%= rs("rel_collo_height") %>" size="4"> m</td>
				</tr>
				<tr>
					<td class="label">lunghezza collo</td>
					<td class="content" colspan="4"><input type="text" class="number" name="tfn_rel_collo_lenght" value="<%= rs("rel_collo_lenght") %>" size="4"> m</td>
				</tr>
				<tr>
					<td class="label">volume collo</td>
					<td class="content" colspan="4"><input type="text" class="number" name="tfn_rel_collo_volume" value="<%= rs("rel_collo_volume") %>" size="4"> m</td>
				</tr>
			<% else %>
				<input type="hidden" name="tfn_rel_peso_netto" value="<%= rs("rel_peso_netto") %>">
				<input type="hidden" name="tfn_rel_peso_lordo" value="<%= rs("rel_peso_lordo") %>">
				<input type="hidden" name="tfn_rel_colli_num" value="<%= rs("rel_colli_num") %>">
				<input type="hidden" name="tfn_rel_collo_pezzi_per" value="<%= rs("rel_collo_pezzi_per") %>">
				<input type="hidden" name="tfn_rel_collo_width" value="<%= rs("rel_collo_width") %>">
				<input type="hidden" name="tfn_rel_collo_height" value="<%= rs("rel_collo_height") %>">
				<input type="hidden" name="tfn_rel_collo_lenght" value="<%= rs("rel_collo_lenght") %>">
				<input type="hidden" name="tfn_rel_collo_volume" value="<%= rs("rel_collo_volume") %>">
			<% end if %>
			<tr><th colspan="6">DESCRIZIONE</th></tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
					<td class="content" colspan="6">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>
								<td width="6%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<td><textarea style="width:100%;" rows="4" name="tft_rel_descr_<%= Application("LINGUE")(i) %>"><%= rs("rel_descr_" & Application("LINGUE")(i)) %></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			<%next
            
            CALL Form_DatiModifica(conn, rs, "rel_") %>
			<tr>
				<td class="footer" colspan="6">
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
<% 
rs.close
conn.close
set rs = nothing
set rsv = nothing
set conn = nothing

%> 
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>