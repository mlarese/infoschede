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
sezione_testata = "inserimento nuova variante" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>
<SCRIPT LANGUAGE="javascript"  src="Tools_B2B.js" type="text/javascript"></SCRIPT>

<% 
dim conn, rs, rsa, rsv, sql, AllVarianti, value,i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsv = Server.CreateObject("ADODB.Recordset")
set rsa = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM gtb_articoli WHERE art_id=" & cIntero(request("ART_ID"))
rsa.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>
<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="tfn_rel_art_id" value="<%= request("ART_ID") %>">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>Inserimento nuova variante per l'articolo</caption>
			<tr><th colspan="7">DATI VARIANTE</th></tr>
			<% sql = " SELECT * FROM gtb_varianti WHERE var_id IN (SELECT val_var_id FROM gtb_valori " + _
					 " INNER JOIN grel_art_vv ON gtb_valori.val_id=grel_art_vv.rvv_val_id " + _
					 " INNER JOIN grel_art_valori ON grel_art_vv.rvv_art_var_id = grel_art_valori.rel_id " + _
					 " WHERE grel_art_valori.rel_Art_id=" & cIntero(request("ART_ID")) & ") ORDER BY var_ordine"
			rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
			AllVarianti = false
			if rs.eof then
				rs.close
				'nessuna variante gi&agrave; associata all'articolo: le mostra tutte.
				sql = "SELECT * FROM gtb_varianti ORDER BY var_ordine "
				rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
				AllVarianti = true
			end if
			while not rs.eof
				sql = "SELECT * FROM gtb_valori WHERE val_var_id=" & rs("var_id") & " ORDER BY val_ordine"
				rsv.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
				if not rsv.eof then%>
					<input type="hidden" name="cod_int_<%= rs.AbsolutePosition %>" value="">
					<input type="hidden" name="cod_pro_<%= rs.AbsolutePosition %>" value="">
					<input type="hidden" name="cod_alt_<%= rs.AbsolutePosition %>" value="">
					<tr>
						<td class="label" colspan="2"><%= rs("var_nome_it") %></td>
						<td class="content">
							<%rsv.moveFirst
							CALL DropDownRecordSet(rsv, "val_id", "val_nome_it", "valore_" & rs.AbsolutePosition, request("valore_" & rs.AbsolutePosition), _
												  not(AllVarianti), " onchange=""SelezioneValore_" & rs.AbsolutePosition & "(this);""", LINGUA_ITALIANO)%>
						</td>
					</tr>
					<script language="JavaScript" type="text/javascript">
						function SelezioneValore_<%= rs.AbsolutePosition %>(obj_valori){
							switch(obj_valori.selectedIndex){
								<% if AllVarianti then %>
									case 0:{
										form1.cod_int_<%= rs.AbsolutePosition %>.value = '';
										form1.cod_alt_<%= rs.AbsolutePosition %>.value = '';
										form1.cod_pro_<%= rs.AbsolutePosition %>.value = '';
										break;
									}
								<% end if
								rsv.moveFirst
								while not rsv.eof %>
									case <%= IIF(AllVarianti, rsv.AbsolutePosition, rsv.AbsolutePosition-1) %>:{
										form1.cod_int_<%= rs.AbsolutePosition %>.value = '<%= rsv("val_cod_int") %>';
										form1.cod_alt_<%= rs.AbsolutePosition %>.value = '<%= rsv("val_cod_alt") %>';
										form1.cod_pro_<%= rs.AbsolutePosition %>.value = '<%= rsv("val_cod_pro") %>';
										break;
									}
									<% rsv.movenext
								wend %>
							}
							RebuildCodes();
						}
					</script>
				<%end if
				rsv.close
				rs.movenext
			wend%>
			<script language="JavaScript" type="text/javascript">
				//funzione che costruisce i codici a partire dagli input che mantengono le parti di codice delle varianti selezionate
				function RebuildCodes(){
					var cod_int='<%= rsa("art_cod_int") %>';
					var cod_alt='<%= rsa("art_cod_alt") %>';
					var cod_pro='<%= rsa("art_cod_pro") %>';
					<% rs.MoveFirst
					while not rs.eof %>
						if(form1.cod_int_<%= rs.AbsolutePosition %>.value!='' && cod_int!=''){
							cod_int += '<%= CODE_SEPARATOR %>';
							cod_int += form1.cod_int_<%= rs.AbsolutePosition %>.value;
						}
						if (form1.cod_alt_<%= rs.AbsolutePosition %>.value!='' && cod_alt!=''){
							cod_alt += '<%= CODE_SEPARATOR %>';
							cod_alt += form1.cod_alt_<%= rs.AbsolutePosition %>.value;
						}
						if (form1.cod_pro_<%= rs.AbsolutePosition %>.value!='' && cod_pro!=''){
							cod_pro += '<%= CODE_SEPARATOR %>';
							cod_pro += form1.cod_pro_<%= rs.AbsolutePosition %>.value;
						}
						<% rs.moveNext
					wend %>
					if (cod_int != '')
						form1.tft_rel_cod_int.value = cod_int;
					if (cod_alt != '')
						form1.tft_rel_cod_alt.value = cod_alt;
					if (cod_pro != '')
						form1.tft_rel_cod_pro.value = cod_pro;
				}
				
				function SetImputState(){
					var stato = document.getElementById("tipo_prezzo_indipendente")
					DisableControl(form1.tfn_rel_var_euro, stato.checked);
					DisableControl(form1.tfn_rel_var_sconto, stato.checked);
				}
				
				function ApplicaVariazioniPrezzo(tag){
					var altra_variazione;
					var prezzo_base = form1.prezzo_articolo;
					var prezzo_attuale = form1.tfn_rel_prezzo;
					if (tag.name == ('tfn_rel_var_euro')){
						//applica variazione in euro
						altra_variazione = form1.tfn_rel_var_sconto;
						
						//applica variazione in euro al prezzo
						CalcolaPrezzoEuro(prezzo_base, prezzo_attuale, tag)
					}
					else{
						//applica variazione in percentuale
						altra_variazione = form1.tfn_rel_var_euro;
						
						//applica variazione percentuale al prezzo
						CalcolaPrezzo(prezzo_base, prezzo_attuale, tag)
					}
					altra_variazione.value = '0,00';
					tag.value = FormatNumber(tag.value, 2);
					
				}
				
				function RicalcolaVariazioniPrezzo(prezzo_attuale){
					var value_euro, value_perc;
					var prezzo_base = form1.prezzo_articolo;
					var prezzo_var_euro = form1.tfn_rel_var_euro;
					var prezzo_var_perc = form1.tfn_rel_var_sconto;
					
					value_euro = toNumber(prezzo_var_euro.value);
					value_perc = toNumber(prezzo_var_perc.value);
					
					if (value_euro){
						//presente una variazione espressa in euro
						//azzera variazione in percentuale
						prezzo_var_perc.value = '0,00';
						//ricalcola variazione in euro
						CalcolaDifferenza(prezzo_base, prezzo_attuale, prezzo_var_euro);
					}
					else{
						//presente variazione in percentuale o nessuna variazione preesistente
						//azzera variazione in euro
						prezzo_var_euro.value='0,00';
						//ricalcola variazione in percentuale
						CalcolaVariazione(prezzo_base, prezzo_attuale, prezzo_var_perc);
					}
				}
	
			</script>
			<input type="hidden" name="prezzo_articolo" value="<%= rsa("art_prezzo_base") %>">
			<tr>
				<td class="label" rowspan="3" style="width:11%;">codici</td>
				<td class="label" style="width:31%;">interno:</td>
				<td class="content">
					<input type="text" class="text" name="tft_rel_cod_int" value="<%= LoadValue("tft_rel_cod_int", rsa("art_cod_int"))%>" maxlength="50" size="20">
					(*)
				</td>
			</tr>
			<tr>
				<td class="label">alternativo:</td>
				<td class="content"><input type="text" class="text" name="tft_rel_cod_alt" value="<%= LoadValue("tft_rel_cod_alt", rsa("art_cod_alt"))%>" maxlength="50" size="20"></td>
			</tr>
			<tr>
				<td class="label">produttore:</td>
				<td class="content"><input type="text" class="text" name="tft_rel_cod_pro" value="<%= LoadValue("tft_rel_cod_pro", rsa("art_cod_pro"))%>" maxlength="50" size="20"></td>
			</tr>
			<tr><th colspan="7">PREZZO VARIANTE</th></tr>
			<tr>
				<td class="label" rowspan="6" style="width:11%;">prezzo:</td>
				<td class="label" rowspan="2">
					calcola prezzo variante:
				</td>
				<td class="content">
					<input type="radio" class="checkbox" name="tfn_rel_prezzo_indipendente" id="tipo_prezzo_dipendente" value="0" <%= chk(cInteger(request("tfn_rel_prezzo_indipendente"))=0) %> onclick="SetImputState()">
					da prezzo articolo sulla base delle variazioni applicate
				</td>
			</tr>
			<tr>
				<td class="content">
					<input type="radio" class="checkbox" name="tfn_rel_prezzo_indipendente" id="tipo_prezzo_indipendente" value="1" <%= chk(cInteger(request("tfn_rel_prezzo_indipendente"))=1) %> onclick="SetImputState()">
					indipendente dal prezzo articolo
				</td>
			</tr>
			<tr>
				<td class="label">
					prezzo della variante:
				</td>
				<td class="content"><input type="text" class="number" name="tfn_rel_prezzo" value="<%= FormatPrice(LoadValue("tfn_rel_prezzo", rsa("art_prezzo_base")) , 2, false) %>" size="7" onchange="RicalcolaVariazioniPrezzo(this)"> &euro;&nbsp;&nbsp;&nbsp;&nbsp;(*)</td>
			</tr>
			<tr>
				<td class="label" rowspan="2">
					variazioni da prezzo articolo:
				</td>
				<td class="content" style="padding-left:19px"><input type="text" class="number" name="tfn_rel_var_sconto" value="<%= FormatPrice(LoadValue("tfn_rel_var_sconto", 0) , 2, false) %>" size="4" onchange="ApplicaVariazioniPrezzo(this)"> %</td>
			</tr>
			<tr>
				<td class="content" style="padding-left:19px"><input type="text" class="number" name="tfn_rel_var_euro" value="<%= FormatPrice(LoadValue("tfn_rel_var_euro", 0) , 2, false) %>" size="4" onchange="ApplicaVariazioniPrezzo(this)"> &euro;</td>
			</tr>
			<tr>
				<td class="label">classe di sconto per quantit&agrave;:</td>
				<td class="content">
					<%CALL dropDown(conn, "SELECT scc_id, scc_nome FROM gtb_scontiQ_classi ORDER BY scc_nome", _
									"scc_id", "scc_nome", "nfn_rel_scontoQ_id", LoadValue("nfn_rel_scontoQ_id", rsa("art_scontoQ_id")), false, "", LINGUA_ITALIANO)%>
				</td>
			</tr>
			<tr>
				<th class="L2" colspan="6">specifiche aggiuntive sul prezzo:</th>
			</tr>
			<% for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
					<td class="content" colspan="7">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>
								<td width="6%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<td><textarea style="width:100%;" rows="2" name="tft_rel_descr_prezzo_<%= Application("LINGUE")(i) %>"><%= request("tft_rel_descr_prezzo_" & Application("LINGUE")(i)) %></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			<% next %>
			<tr><th colspan="7">DATI  DI GESTIONE</th></tr>
			<tr>
				<td class="label" colspan="2">non visibile a catalogo:</td>
				<td class="content"><input type="checkbox" class="checkbox" name="chk_rel_disabilitato" <%= chk(LoadValue("chk_rel_disabilitato", IIF(rsa("art_disabilitato"), "1", ""))<>"") %>></td>
			</tr>
			<tr>
				<td class="label" rowspan="3" style="width:11%;">gestione:</td>
				<td class="label">giacenza minima</td>
				<% value = LoadValue("tfn_rel_giacenza_min", rsa("art_giacenza_min"))
				value = IIF(value<>"", value, "1") %>
				<td class="content"><input type="text" class="number" name="tfn_rel_giacenza_min" value="<%= value %>" size="4"></td>
			</tr>
			<tr>
				<td class="label">quantit&agrave; minima ordinabile</td>
				<% value = LoadValue("tfn_rel_qta_min_ord", rsa("art_qta_min_ord"))
				value = IIF(value<>"", value, "1") %>
				<td class="content"><input type="text" class="number" name="tfn_rel_qta_min_ord" value="<%= value %>" size="4"></td>
			</tr>
			<tr>
				<td class="label">lotto di riordino</td>
				<% value = LoadValue("tfn_rel_lotto_riordino", rsa("art_lotto_riordino"))
				value = IIF(value<>"", value, "1") %>
				<td class="content"><input type="text" class="number" name="tfn_rel_lotto_riordino" value="<%= value %>" size="4"></td>
			</tr>
 			<tr>
				<td class="label" colspan="2">foto:</td>
				<td class="content">
				<% 	dropDown conn, "SELECT * FROM gtb_art_foto WHERE fo_articolo_id=" & cIntero(request("ART_ID")) & " ORDER BY fo_ordine", "fo_id", "fo_thumb", "nfn_rel_foto_id", request.form("nfn_rel_foto_id"), false, "", LINGUA_ITALIANO %>
				</td>
			</tr>
			<% if cBoolean(cString(Session("ATTIVA_GESTIONE_COLLI-PESI-DIMENSIONI")), false) then %>
				<tr><th colspan="6">COLLI, PESI E DIMENSIONI</th></tr>
				<tr>
					<td class="label" colspan="2">peso netto:</td>
					<td class="content" colspan="4"><input type="text" class="text" name="tfn_rel_peso_netto" value="<%= request("tfn_rel_peso_netto") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">peso lordo:</td>
					<td class="content" colspan="4"><input type="text" class="text" name="tfn_rel_peso_lordo" value="<%= request("tfn_rel_peso_lordo") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">numero colli:</td>
					<td class="content" colspan="4"><input type="text" class="text" name="tfn_rel_colli_num" value="<%= IIF(request("tfn_rel_colli_num")<>"",request("tfn_rel_colli_num"),"1") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">numero pezzi per collo:</td>
					<td class="content" colspan="4"><input type="text" class="text" name="tfn_rel_collo_pezzi_per" value="<%= IIF(request("tfn_rel_collo_pezzi_per")<>"",request("tfn_rel_collo_pezzi_per"),"1") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">larghezza collo:</td>
					<td class="content" colspan="4"><input type="text" class="text" name="tfn_rel_collo_width" value="<%= request("tfn_rel_collo_width") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">altezza collo:</td>
					<td class="content" colspan="4"><input type="text" class="text" name="tfn_rel_collo_height" value="<%= request("tfn_rel_collo_height") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">lunghezza collo:</td>
					<td class="content" colspan="4"><input type="text" class="text" name="tfn_rel_collo_lenght" value="<%= request("tfn_rel_collo_lenght") %>" size="7"></td>
				</tr>
				<tr>
					<td class="label" colspan="2">volume collo:</td>
					<td class="content" colspan="4"><input type="text" class="text" name="tfn_rel_collo_volume" value="<%= request("tfn_rel_collo_volume") %>" size="7"></td>
				</tr>
			<% else %>
				<input type="hidden" name="tfn_rel_peso_netto" value="0">
				<input type="hidden" name="tfn_rel_peso_lordo" value="0">
				<input type="hidden" name="tfn_rel_colli_num" value="1">
				<input type="hidden" name="tfn_rel_collo_pezzi_per" value="1">
				<input type="hidden" name="tfn_rel_collo_width" value="0">
				<input type="hidden" name="tfn_rel_collo_height" value="0">
				<input type="hidden" name="tfn_rel_collo_lenght" value="0">
				<input type="hidden" name="tfn_rel_collo_volume" value="0">
			<% end if %>
			<tr><th colspan="7">DESCRIZIONE</th></tr>
			<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
				<tr>
					<td class="content" colspan="7">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
							<tr>
								<td width="6%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
								<td><textarea style="width:100%;" rows="4" name="tft_rel_descr_<%= Application("LINGUE")(i) %>"><%= request("tft_rel_descr_" & Application("LINGUE")(i)) %></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			<%next %>
			<tr>
				<td class="footer" colspan="5">
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
<script language="JavaScript" type="text/javascript">
	
	<% rs.moveFirst
	while not rs.eof%>
		SelezioneValore_<%= rs.AbsolutePosition %>(form1.valore_<%= rs.AbsolutePosition %>);
		<%rs.movenext
	wend
	
	if not Request.ServerVariables("REQUEST_METHOD")="POST" then  %>
		RebuildCodes();
	<% end if %>
	
</script>
<% 
rs.close
rsa.close
conn.close
set rs = nothing
set rsa = nothing
set rsv = nothing
set conn = nothing 


function LoadValue(RequestField, RecordValue)
	if Request.ServerVariables("REQUEST_METHOD")="POST" then
		LoadValue = request(RequestField)
	else
		LoadValue = RecordValue
	end if
end function
%>

<script language="JavaScript" type="text/javascript">
	SetImputState();
	FitWindowSize(this);
</script>