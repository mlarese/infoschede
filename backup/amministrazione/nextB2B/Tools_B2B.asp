<!--#INCLUDE VIRTUAL="/nextB2B_Integration/NEXTb2b_Events.asp" -->
<% 
'.................................................................................................
'.................................................................................................
'.................................................................................................
'FUNZIONI E COSTANTI PER IL NEXT-B2B
'.................................................................................................
'.................................................................................................

'.................................................................................................
'COSTANTI
'.................................................................................................
'separatore per la generazione del codice della variante dell'articolo (da sostituire con un eventuale parametro)
const CODE_SEPARATOR = "."

'costanti per la movimentazione delle quantita'
const QTA_IMPEGNATA = "I"
const QTA_ORDINATA = "O"
const QTA_GIACENZA = "G"

'costanti per i permessi per utenti area amministrativa
const POS_PERMESSO_ADMIN = 1
const POS_PERMESSO_AGENTE = 1

'costanti per i permessi degli utenti dell'area riservata
const UTENTE_PERMESSO_CLIENTE = "B2B_I_CLIENTE"
const UTENTE_PERMESSO_SUBCLIENTE = "B2B_I_CLIENTE_INTERNO"
const UTENTE_PERMESSO_AGENTE = "B2B_I_AGENTE"

'costanti per la gestione dell'ordine
const ORDINE_NON_CONFERMATO = 0
const ORDINE_CONFERMATO 	= 1
const ORDINE_EVASO 			= 2
const ORDINE_ARCHIVIATO 	= 3

dim STATI_ORDINE, STILI_STATI_ORDINE, STATI_ORDINE_TIPO
STATI_ORDINE = Array("non confermato", "confermato", "evaso", "archiviato", "annullato")
STATI_ORDINE_TIPO = Array(0, 1, 2, 3, 4)
STILI_STATI_ORDINE = Array("_b", " OrdConfermato", " OrdEvaso", "", " OrdAnnullato")


'.................................................................................................
'GESTIONE DELLE CATEGORIE
'.................................................................................................
dim categorie
set categorie = New objCategorie
with categorie
	.tabella = "gtb_tipologie"
	.prefisso = "tip"
	
    'gestione delle relazioni con le tabelle categorizzate
    CALL .AddRelazione("gtb_articoli", "art_id", "art_tipologia_id", true)
	
	.tabellaRelCaratteristiche = "gtb_tip_ctech"
	.chiaveEsternaRelCaratteristiche = "rct_tipologia_id"
	.ordineRelCaratteristiche = "rct_ordine"
	.idCarRelCaratteristiche = "rct_ctech_id"
	
	.tabellaCaratteristiche = "gtb_carattech"
	.idCaratteristiche = "ct_id"
	.nomeCaratteristiche = "ct_nome_it"
	.tipoCaratteristiche = "ct_tipo"
	.tabellaRelCorCaratteristiche = "grel_art_ctech"
	.idArtRelCorCaratteristiche = "rel_art_id"
	.idCarRelCorCaratteristiche = "rel_ctech_id"
	
	if Session("B2B_ABILITA_DESCRIZIONE_HTML") then 
		.attivaCKEditorPerDescrizione = true
	end if 
	
	.isB2B = true
	.GestioneCategorieMiste = false
	
	'abilitazione indice e contenuti
	.Index = Index
end with


'*********************************************************************
'GESTIONE FOTOGRAFIE DEGLI ARTICOLI
'*********************************************************************
'restituisce la ClassPhotoGallery Impostata correttamente

dim oArticoliFoto
set oArticoliFoto = new ClassPhotoGallery
with oArticoliFoto
	.WebId 						= Application("AZ_ID")
	.Index 						= Index
	.TableName 					= "gtb_art_foto"
	.FieldPrefix 				= "fo"
	.FieldForeignKey 			= "fo_articolo_id"
	.FilePrefix 				= "Articoli"
	.FotoSingola 				= false
	.DeleteKey 					= "ARTICOLI_FOTO"
	
	.ElementTableName 			= "gtb_articoli"
	.ElementFieldPrefix 		= "art"
	.ElementUpdateParams 		= true
	
	.Abilita_ElencoAddOn 		= true
		
	.NoteFormatoThumbnail 		= ""
	.NoteFormatoZoom 			= ""
	
	.TipiFotoTableName = "gtb_foto_tipo"
end with

'variabile usata nei due addon sotto per passare i dati dalla testata alle righe, leggendoli una sola volta.
dim ElencoTestata_ADDON_art_varianti

'funzione eseguita nella testata della tabella delle foto
sub ElencoTestata_ADDON(conn)
	dim sql
	sql = "SELECT art_id FROM gtb_articoli WHERE IsNull(art_varianti,0)=1 AND art_id = " & cIntero(request("ID"))
	ElencoTestata_ADDON_art_varianti = cIntero(GetValueList(conn, NULL, sql))>0
	
	if ElencoTestata_ADDON_art_varianti then %>
		<th class="L2" style="width:15%;">Varianti collegate</th>
	<% end if
	
	
end sub

'funzinoe eseguita per ogni riga della tabella delle foto
sub ElencoRow_ADDON(conn, rs)
	if ElencoTestata_ADDON_art_varianti then 
		dim sql, rsv 
		set rsv = server.createobject("ADODB.Recordset") 
		sql = "SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & cIntero(request("ID")) & " AND rel_foto_id=" & rs("fo_id") 
		rsv.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText%>
		<td class="content">
			<% if rsv.eof then %>
				&nbsp;
			<% else
				while not rsv.eof 
					CALL TableValoriVarianti(conn, server.createobject("ADODB.Recordset"), rsv("rel_id"), "content")
					rsv.movenext
				wend
			end if %>
		</td>
		<% rsv.close
	end if
	
end sub

'.................................................................................................
'FUNZIONI
'.................................................................................................

'.................................................................................................
'..			elenca valori varianti
'..			conn					aperta sul database
'..			rs						oggetto recordset chiuso e creato
'..			grel_art_valori_id 		id del prodotto-variante di cui reperire i dati delle varianti
'.................................................................................................
function ListValoriVarianti(conn, rs, grel_art_valori_id)
	dim sql
	sql = " SELECT val_nome_it, var_nome_it, val_cod_int FROM gv_articoli_varianti WHERE rvv_art_var_id=" & cIntero(grel_art_valori_id)
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	if not rs.eof then%>
		<span class="note">
			(
			<% while not rs.eof %>
				<span title="<%= rs("var_nome_it") %><%= IIF(cString(rs("val_cod_int"))<>"", ", cod: " & rs("val_cod_int"), "") %>">
					<%= rs("val_nome_it") %>
				</span>
				<% rs.movenext
				if not rs.eof then %>
					- 
				<% end if %>
			<% wend %>
			)
		</span>
	<%end if
	rs.close
end function


'.................................................................................................
'..			elenca valori varianti in una tabella con relative varianti
'..			conn					aperta sul database
'..			rs						oggetto recordset chiuso e creato
'..			grel_art_valori_id 		id del prodotto-variante di cui reperire i dati delle varianti
'.................................................................................................
function TableValoriVarianti(conn, rs, grel_art_valori_id, CssClass)
	dim sql
	sql = " SELECT val_nome_it, var_nome_it, val_cod_int FROM gv_articoli_varianti WHERE rvv_art_var_id=" & cIntero(grel_art_valori_id)
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	if not rs.eof then%>
		<table cellpadding="0" cellspacing="0">
			<tr>
				<% while not rs.eof %>
					<td class="<%= CssClass %>">
						<span class="note"><%= rs("var_nome_it") %>:</span>
						<span title="<%= IIF(cString(rs("val_cod_int"))<>"", "cod: " & rs("val_cod_int"), "") %>">
							<%= rs("val_nome_it") %>
						</span>
					</td>
					<% rs.movenext
					if not rs.eof then%>
						<td style="font-size:1px; width:10px;">&nbsp;</td>
					<%end if
				wend %>
			</tr>
		</table>
	<%end if
	rs.close
end function


'.................................................................................................
'procedura che visualizza il riepilogo dei dati dell'articolo
'		conn:	connessione aperta sul database
'		rs:		recordset aperto contenente il record dell'articolo
'		rsc:	recordset creato e chiuso per letture interne
'.................................................................................................
sub ArticoloScheda (conn, rs, rsc)
	dim txt, SchedaVariante
	SchedaVariante = FieldExists(rs, "rel_id")%>
	<tr>
		<td class="label" style="width:13%;">codici:</td>
		<td class="label" style="width:12%;">interno:</td>
		<td class="content_b" style="width:17%;">
			<% IF SchedaVariante then %>
				<%= rs("rel_cod_int") %>
			<% else %>
				<%= rs("art_cod_int") %>
			<% end if %>
		</td>
		<td class="label" style="width:12%;">alternativo:</td>
		<td class="content" style="width:17%;">
			<% IF SchedaVariante then %>
				<%= rs("rel_cod_alt") %>
			<% else %>
				<%= rs("art_cod_alt") %>
			<% end if %>
		</td>
		<td class="label" style="width:12%;">produttore:</td>
		<td class="content" style="width:17%;">
			<% IF SchedaVariante then %>
				<%= rs("rel_cod_pro") %>
			<% else %>
				<%= rs("art_cod_pro") %>
			<% end if %>
		</td>
	</tr>
	<tr>
		<td class="label">articolo:</td>
		<td class="content" colspan="6">
			<% IF SchedaVariante then 
				ArticoloLink rs("art_id"), rs("art_nome_it"), rs("rel_cod_int")
			else
				ArticoloLink rs("art_id"), rs("art_nome_it"), rs("art_cod_int")
			end if %>
			<% if SchedaVariante then%>
				<%= ListValoriVarianti(conn, rsc, rs("rel_id")) %>
			<% end if %>
		</td>
	</tr>
	<tr>
		<td class="label" <%= IIF(cInteger(rs("art_raggruppamento_id"))>0, " rowspan=""2"" ", "") %>>categoria:</td>
		<td class="content" colspan="6"><%= categorie.NomeCompleto(rs("art_tipologia_id")) %></td>
	</tr>
	<% if cInteger(rs("art_raggruppamento_id"))>0 then 
		sql = "SELECT rag_nome_it FROM gtb_tipologie_raggruppamenti WHERE rag_id=" & rs("art_raggruppamento_id")%>
		<tr>
			<td class="label" colspan="2">raggruppamento pubbl.:</td>
			<td class="content" colspan="4">
				<%= GetValueList(conn, rsc, sql) %>
			</td>
		</tr>
	<% end if %>
	<tr>
		<td class="label">tipo:</td>
		<% if rs("art_se_bundle") then %>
			<td class="content bundle" colspan="2">bundle</td>
		<% elseif rs("art_se_confezione") then %>
			<td class="content confezione" colspan="2">confezione</td>
		<% elseif rs("art_varianti") then %>
			<td class="content varianti" colspan="2">articolo con varianti</td>
		<% else %>
			<td class="content" colspan="2">articolo singolo</td>
		<% end if %>
		<td class="label" colspan="2">marchio / produttore:</td>
		<td class="content" colspan="3"><%= rs("mar_nome_it") %></td>
	</tr>
	<tr>
		<td class="label">stato:</td>
		<td class="content" colspan="2">
			<% if rs("art_disabilitato") then %>
				non a catalogo
			<% else %>
				a catalogo
			<% end if %>
		</td>
		<td class="content" colspan="4">
			<table width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<% if rs("art_in_bundle") then %>
						<td class="content bundle">in bundle</td>
					<% end if
					if rs("art_in_confezione") then%>
						<td class="content confezione">in confezione</td>
					<% end if
					if rs("art_se_accessorio") then %>
						<td class="content">
							<% if rs("art_NoVenSingola") then %>
								articolo collegato non vendibile singolarmente
							<% else %>
								articolo colleagto
							<% end if %>
						</td>
					<% end if
					if rs("art_ha_accessori") then%>
						<td class="content">con articoli collegati</td>
					<% end if %>
				</tr>
			</table>
		</td>
	</tr>
<% end sub 


'.................................................................................................
'funzione che stampa il valore passato con il link per aprire la scheda
'	id:		id dell'articolo di cui aprire la scheda
'	label:	valore a cui riferire il link (tipicamente nome o codice articolo)
'.................................................................................................
sub ArticoloLink(id, label, codice)%>
	<a href="javascript:void(0);" title="apri scheda dell'articolo in una nuova finestra" <%= ACTIVE_STATUS %>
	   onclick="OpenAutoPositionedScrollWindow('<%= GetAmministrazionePath() %>nextB2b/ArticoliMod.asp?ID=<%= id %>#<%= Server.HTMLEncode(codice) %>', 'articolo', 760, 400, true);">
		<%= label %>
	</a>
<%end sub


'.................................................................................................
'funzione che stampa il valore passato con il link per aprire la scheda del cliente
'	id:		id del cliente di cui aprire la scheda
'	label:	valore a cui riferire il link 
'.................................................................................................
sub ClienteLink(id, label)%>
	<a href="javascript:void(0);" title="apri scheda del cliente in una nuova finestra" <%= ACTIVE_STATUS %>
	   onclick="OpenAutoPositionedScrollWindow('<%= GetAmministrazionePath() %>nextB2b/ClientiGestione.asp?ID=<%= id %>', 'cliente', 760, 400, true);">
		<%= label %>
	</a>
<%end sub


'.................................................................................................
'scrivo i prezzi scontati
'		conn:		connessione aperta a database
'		rs_art:		recorset aperto con dati dell'articolo: id variante, listino, classe di sconto
'		rsc:		recordset chiuso e creato
'		colspan		colspan della colonna che contiene i prezzi
'.................................................................................................
Function ScontiQ(conn, rs_art, rsc, colspan, cambio, valuta)
	dim sql, prezzo
	if cInteger(rs_art("prz_scontoQ_id"))=0 then
		'articolo senza classe di sconto per quantita'
		%>
		<td class="content_right" nowrap title="<%= DettagliPrezzo(rs_art("prz_prezzo"), rs_art("iva_valore"), cambio, valuta) %>" <%= IIF(colspan>1, "colspan=""" & colspan & """", "")%>>
			<%= FormatPrice(rs_art("prz_prezzo"), 2, true) %> &euro;
		</td>
	<%else
		'articolo con classe di sconto per quantita'
		sql = " SELECT * FROM gtb_scontiQ_classi INNER JOIN gtb_scontiQ ON gtb_scontiQ_classi.scc_id=gtb_scontiQ.sco_classe_id " + _
		  	  " WHERE scc_id=" & rs_art("prz_scontoQ_id")
		rsc.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		if not rsc.eof then%>
			<td <%= IIF(colspan>1, "colspan=""" & colspan & """", "")%>>
				<table style="height:100%;" cellpadding="0" cellspacing="1" width="100%" class="tabella_madre" style="border:0px;">
					<tr>
						<td class="content_disabled" nowrap title="quantit&agrave; di partenza dell'applicazione dello sconto">
							<% if cInteger(rsc("sco_qta_da"))>2 then %>
								1 - <%= cInteger(rsc("sco_qta_da"))-1 %>
							<% else %>
								1
							<% end if %>
						</td>
						<td class="content_right" title="<%= DettagliPrezzo(rs_art("prz_prezzo"), rs_art("iva_valore"), cambio, valuta) %>" nowrap><%= FormatPrice(rs_art("prz_prezzo"), 2, true) %> &euro;</td>
					</tr>
					<% while not rsc.eof %>
						<tr>
							<td class="content_disabled" nowrap title="quantit&agrave; di partenza dell'applicazione dello sconto">
								<%= rsc("sco_qta_da") %>
								<% rsc.movenext
								if not rsc.eof then %>
									- <%= rsc("sco_qta_da") %>
								<% else %>
									+
								<% end if 
								rsc.movePrevious%>
							</td>
							<% prezzo = GetPricePercent(rs_art("prz_prezzo"), rsc("sco_sconto")) %>
							<td class="content_right" nowrap title="<%= DettagliPrezzo(prezzo, rs_art("iva_valore"), cambio, valuta) &vbCrLF %>sconto applicato:<%= FormatPrice(rsc("sco_sconto"), 2, true) %>%">
								<%= FormatPrice(prezzo, 2, true) %> &euro;
							</td>
						</tr>
						<% rsc.movenext
					wend %>
				</table>
			</td>
		<%else%>
			<td class="content_right" nowrap <%= IIF(colspan>1, "colspan=""" & colspan & """", "")%>>
				<%= FormatPrice(rs_art("prz_prezzo"), 2, true) %> &euro;
			</td>
		<%end if
		rsc.close
	end if
End Function


'.................................................................................................
'ritorna un testo riassuntivo di descrizione del prezzo secondo le caratteristiche impostate:
'	prezzo:			prezzo dell'articolo
'	aliquota:		aliquota iva a cui e' soggetto l'articolo
'	cambio			rapporto di cambio della valuta del cliente
'	valuta:			valuta del cliente
'.................................................................................................
function DettagliPrezzo(prezzo, aliquota, cambio, valuta)
	dim iva
	DettagliPrezzo = "prezzo netto:" & FormatPrice((prezzo * cambio), 2, true) & " " & valuta & vbCrLf
	if aliquota > 0 then
		iva = GetIva(prezzo, aliquota)
		DettagliPrezzo = DettagliPrezzo & "i.v.a.:" & FormatPrice((iva * cambio), 2, true) & " " & valuta & _
						 " (" & aliquota & "%) " & vbCrLF
	else
		iva = 0
		DettagliPrezzo = DettagliPrezzo & "i.v.a.: esente " & vbcRLf
	end if
	DettagliPrezzo = DettagliPrezzo & "prezzo iva inclusa:" & FormatPrice(((prezzo + iva) * cambio), 2, true) & " " & valuta
end function


'-------------------------------------------------------------------------------------------------GESTIONE GIACENZA

'.................................................................................................
'	setta la giacenza dato l'ordine o il dettaglio
'	conn:					connessione aperta a database
'	ID:						id dell'ordine o del dettaglio
'	tipoID:					se "O" l'ID è dell'ordine else "D" ID del dettaglio
'	op:						"+" o "-"
'	tipo:					secondo costanti per la movimentazione delle quantita' definite inizio pagina
'	magazzinoID:			gia_magazzino_id
'.................................................................................................
Function SetGiacenza_ord(conn, ID, tipoID, op, tipo, magazzinoID)
dim sql, rs
	set rs = server.createobject("adodb.recordset")

	if tipoID = "D" then
		sql = "SELECT det_art_var_id, det_qta FROM gtb_dettagli_ord WHERE ISNULL(det_art_var_id,0)>0 AND det_id="& cIntero(ID)
		rs.open sql, conn, adOpenStatic, adLockOptimistic
		if not rs.eof then
			CALL SetGiacenza(conn, rs("det_art_var_id"), op, tipo, magazzinoID, rs("det_qta"))
		end if
		rs.close
	else
		sql = "SELECT det_art_var_id, det_qta FROM gtb_dettagli_ord WHERE ISNULL(det_art_var_id,0)>0 AND det_ord_id="& cIntero(ID)
		rs.open sql, conn, adOpenStatic, adLockOptimistic
		while not rs.eof
			CALL SetGiacenza(conn, rs("det_art_var_id"), op, tipo, magazzinoID, rs("det_qta"))
			rs.movenext
		wend
		rs.close
	end if

	set rs = nothing
End Function

'.................................................................................................
'	setta la giacenza dato gia_art_var_id
'	conn:					connessione aperta a database
'	ID:						gia_art_var_id
'	op:						"+" o "-"
'	tipo:					secondo costanti per la movimentazione delle quantita' definite inizio pagina
'							se "I" impegna la merce, "G" giacenza, "O" ordinato
'	magazzinoID:			gia_magazzino_id
'	qta:					variazione assoluta di quantita'
'.................................................................................................
Function SetGiacenza(conn, ID, op, tipo, magazzinoID, qta)
dim sql, rs
	set rs = server.createobject("adodb.recordset")
	
	'caso usato solo in modifica dei dettagli di ordini
	if request("old_qta")<>"" AND request("tfn_det_ord_id")<>"" AND request("tfn_det_art_var_id")<>"" then
		qta = (cInteger(request("old_qta")) - qta)
	end if
	
	sql = "SELECT b.* FROM gtb_bundle b "& _
		  "LEFT JOIN gv_articoli a ON b.bun_bundle_id = a.rel_id "& _
		  "WHERE "& SQL_IsTrue(conn, "art_se_bundle") &" AND bun_bundle_id="& cIntero(ID)
	rs.open sql, conn, adOpenStatic, adLockOptimistic
	if rs.eof then								'sono articolo singolo, componente o confezione
		CALL SetGiacenza_nb(conn, ID, op, tipo, magazzinoID, qta)
	else										'sono un bundle
		'update componenti
		while not rs.eof
			CALL SetGiacenza_nb(conn, rs("bun_articolo_id"), op, tipo, magazzinoID, qta * rs("bun_quantita"))
			rs.movenext
		wend
	end if
	rs.close

	set rs = nothing
End Function

'.................................................................................................
'	setta la giacenza di un componente o articolo singolo dato gia_art_var_id e del bundle
'	conn:					connessione aperta a database
'	ID:						gia_art_var_id
'	op:						"+" o "-"
'	tipo:					secondo costanti per la movimentazione delle quantita' definite inizio pagina
'							se "I" impegna la merce, "G" giacenza, "O" ordinato
'	magazzinoID:			gia_magazzino_id
'	qta:					variazione assoluta di quantita'
'.................................................................................................
Function SetGiacenza_nb(conn, ID, op, tipo, magazzinoID, qta)
dim sql, rs
	set rs = server.createobject("adodb.recordset")

	'update articolo
	sql = "UPDATE grel_giacenze SET "
	if tipo = QTA_IMPEGNATA then			'impegna
		sql = sql & "gia_impegnato=gia_impegnato"
	elseif tipo = QTA_ORDINATA then		'ordine fornitore
		sql = sql & "gia_ordinato=gia_ordinato"
	else						'movimenta
		sql = sql & "gia_qta=gia_qta"
	end if
	sql = sql & op & qta &" WHERE gia_art_var_id="& cIntero(ID) &" AND gia_magazzino_id="& cIntero(magazzinoID)
	CALL conn.execute(sql, , adExecuteNoRecords)
	
	'update bundle associati
	sql = "SELECT b.* FROM gtb_bundle b "& _
		  "LEFT JOIN gv_articoli a ON b.bun_bundle_id = a.rel_id "& _
		  "WHERE "& SQL_IsTrue(conn, "art_se_bundle") &" AND bun_articolo_id="& cIntero(ID)
	rs.open sql, conn, adOpenStatic, adLockOptimistic
	while not rs.eof
		CALL SetGiacenza_b(conn, rs("bun_bundle_id"), tipo, magazzinoID)
		rs.movenext
	wend
	rs.close

	set rs = nothing
End Function

'.................................................................................................
'	setta la giacenza di un bundle o confezione
'	conn:					connessione aperta a database
'	ID:						gia_art_var_id
'	tipo:					se "I" impegna la merce, "G" giacenza, "O" ordinato
'	magazzinoID:			gia_magazzino_id
'.................................................................................................
Function SetGiacenza_b(conn, ID, tipo, magazzinoID)
dim sql, sql_min, campo
	if tipo = QTA_IMPEGNATA then			'impegna
		campo = "gia_impegnato"
		'se la merce  impegnata deve essere bloccato il massimo possibile con arrotondamento comunque al primo intero superiore.
		sql_min = " MAX(CASE WHEN (" + campo + " % bun_quantita)>0 THEN (" + campo + " / bun_quantita) + 1 " + _
				  " ELSE " + campo + " / bun_quantita END) "
	else
		if tipo = QTA_ORDINATA then		'ordine fornitore
			campo = "gia_ordinato"
		else						'movimenta
			campo = "gia_qta"
		end if
		sql_min = " MIN(" + campo + " / bun_quantita) "
	end if
	sql_min = "SELECT " + sql_min + " FROM grel_giacenze g " + _
			  "INNER JOIN gtb_bundle b ON g.gia_art_var_id=b.bun_articolo_id " + _
			  "WHERE gia_magazzino_id = " & cIntero(magazzinoID) & " AND bun_bundle_id=" & cIntero(ID)
	
	sql = "UPDATE grel_giacenze SET " + _
		  campo + "=(" + sql_min + ") " + _
		  " WHERE gia_art_var_id=" & cIntero(ID) & " AND gia_magazzino_id=" & cIntero(magazzinoID)
	CALL conn.execute(sql, , adExecuteNoRecords)

End Function


'-------------------------------------------------------------------------------------------------FINE GESTIONE GIACENZA

'.................................................................................................
'	ritorna la variazione percentuale di prezzo rispetto a prezzo_base
'	prezzo_base		prezzo di riferimento per il conteggio della percentuale
'	prezzo			prezzo di cui calcolare la variazione
'.................................................................................................
function GetVarPercent(prezzo_base, prezzo)
	if CIntero(prezzo_base)>0 then
		GetVarPercent = ArrotondaEuro(((prezzo - prezzo_base) / prezzo_base) * 100)
	else
		GetVarPercent = "-"
	end if
end function


'.................................................................................................
'	ritorna il prezzo risultante dopo aver applicato la variazione
'	prezzo_base		prezzo di riferimento su cui applicare la variazione
'	variazione		variazione percentuale da applicare
'.................................................................................................
function GetPricePercent(prezzo_base, variazione)
	GetPricePercent = ArrotondaEuro(prezzo_base + ((variazione / 100) * prezzo_base))
end function


'.................................................................................................
'	ritorna la parte di prezzo corrispondente alla percentuale
'	prezzo:			prezzo di riferimento di cui applicare l'iva
'	percentuale:	percentuale di iva da calcolare
'.................................................................................................
function GetIVA(prezzo, aliquota)
	GetIVA = ArrotondaEuro(ArrotondaEuro(prezzo) * (aliquota / 100))
end function


'.................................................................................................
'	ritorna il prezzo iva compresa
'	prezzo_netto:	prezzo netto su cui calcolare e sommare l'iva
'	aliquota:		aliquota iva applicata
'.................................................................................................
function GetPrezzoIvato(prezzo_netto, aliquota)
	GetPrezzoIvato = ArrotondaEuro(prezzo_netto) + GetIVA(prezzo_netto, aliquota)
end function


'..................................................................................................
'..		ritorna il prezzo formattato e convertito in valuta
'..		prezzo_euro		prezzo in euro da convertire
'..		cambio_valuta	rapporto di cambio euro/valuta al quale effettuare la conversione
'..................................................................................................
function Cambio(prezzo_euro, cambio_valuta)
	Cambio = ArrotondaEuro(prezzo_euro) * ArrotondaEuro(cambio_valuta)
	Cambio = FormatPrice(Cambio, 2, true)
end function


'.................................................................................................
'	ritorna il prezzo totale della riga d'ordine
'	prezzo_unitario:	prezzo unitario dell'articolo
'	qta:				quantita' richiesta
'.................................................................................................
function GetImporto(prezzo_unitario, qta)
	GetImporto = ArrotondaEuro(prezzo_unitario) * qta
end function


'.................................................................................................
'	procedura che esegue l'aggiornamento dei prezzi delle varianti dal prezzo base dell'articolo
'	solo per le varianti senza prezzo indipendente
'	conn:			connessione su cui eseguire l'aggiornamento
'	art_id		id dell'articolo
'.................................................................................................
sub AggiornaPrezziVarianti(conn, rs, art_id)
	dim sql
	sql = " UPDATE grel_art_valori SET rel_prezzo = (((SELECT art_prezzo_base FROM gtb_articoli WHERE art_id = grel_Art_valori.rel_art_id) + ISNULL(rel_var_euro,0)) * ((100 + ISNULL(rel_var_sconto, 0))/100) )" + _
		  " WHERE rel_art_id=" & cIntero(art_id) & " AND ISNULL(rel_prezzo_indipendente, 0)=0 "
	CALL conn.execute(sql, , adExecuteNoRecords)
	
	sql = "SELECT rel_id FROM grel_art_valori WHERE ISNULL(rel_prezzo_indipendente, 0)=0 AND rel_art_id=" & cIntero(art_id)
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	while not rs.eof 
		CALL AggiornaPrezzoListini(conn, rs("rel_id"))
		rs.movenext
	wend
	rs.close
end sub


'.................................................................................................
'	procedura che esegue l'aggiornamento dei prezzi dei listini base
'	conn:			connessione su cui eseguire l'aggiornamento
'	variante_id		id della variante/articolo di cui aggiornare i prezzi
'.................................................................................................
sub AggiornaPrezzoListiniBase(conn, variante_id)
	CALL AggiornaPrezzoListiniBaseForzabile(conn, variante_id, GetModuleParam(conn, "LISTINI_PREZZI_INDIPENDENTI"))
end sub


'.................................................................................................
'	procedura che esegue l'aggiornamento dei prezzi dei listini base
'	conn:			connessione su cui eseguire l'aggiornamento
'	variante_id		id della variante/articolo di cui aggiornare i prezzi
'	prezziIndipendenti se true permette di saltare l'aggiornamento dei prezzi di listino dal prezzo base
'.................................................................................................
sub AggiornaPrezzoListiniBaseForzabile(conn, variante_id, prezziIndipendenti)
	if not prezziIndipendenti then
		dim sql
		sql = " UPDATE gtb_prezzi SET prz_prezzo = (((SELECT rel_prezzo FROM grel_art_valori WHERE rel_id = gtb_prezzi.prz_variante_id) + ISNULL(prz_var_euro,0)) * ((100 + ISNULL(prz_var_sconto, 0))/100) ) " + _
			" WHERE prz_variante_id=" & cIntero(variante_id) & " AND prz_listino_id IN (SELECT listino_id FROM gtb_listini WHERE ISNULL(listino_base, 0)=1) "
		CALL conn.execute(sql, , adExecuteNoRecords)
	end if
	
end sub


'.................................................................................................
'	procedura che esegue l'aggionamento dei prezzi dei listini non-base per la variante indicata.
'	conn:			connessione su cui eseguire l'aggiornamento
'	variante_id		id della variante/articolo di cui aggiornare i prezzi nei listini
'.................................................................................................
sub AggiornaPrezzoListiniDaListinoBase(conn, variante_id)
	
	if not GetModuleParam(conn, "LISTINI_PREZZI_INDIPENDENTI") then
		dim sql
		sql = " UPDATE gtb_prezzi SET prz_prezzo = (((SELECT prz_prezzo FROM gtb_prezzi INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id=gtb_listini.listino_id " + _
			  " WHERE prz_variante_id=" & cIntero(variante_id) & " AND ISNULL(listino_base_attuale, 0)=1) + ISNULL(prz_var_euro,0)) * ((100 + ISNULL(prz_var_sconto, 0))/100) ) " + _
			  " WHERE prz_variante_id=" & cIntero(variante_id) & " AND prz_listino_id IN (SELECT listino_id FROM gtb_listini WHERE ISNULL(listino_base, 0)=0) "
		CALL conn.execute(sql, , adExecuteNoRecords)
	end if
	
end sub


'.................................................................................................
'	procedura che esegue l'aggiornamento dei prezzi dei listini base per la variante indicata
'	e successivamente i prezzi di tutti gli altri listini
'	conn:			connessione su cui eseguire l'aggiornamento
'	variante_id		id della variante/articolo di cui aggiornare i prezzi
'.................................................................................................
sub AggiornaPrezzoListini(conn, variante_id)
	'aggiorna prezzi dei listini base
	CALL AggiornaPrezzoListiniBase(conn, variante_id)
	
	'aggiorna prezzi dei listini non-base
	CALL AggiornaPrezzoListiniDaListinoBase(conn, variante_id)
end sub



'.................................................................................................
'	funzione che calcola lo sconto per quantit&agrave; sul prezzo, se appilcato
'	conn: 			connessione al database aperta
'	rs:				recordset creato e chiuso
'	prezzo_Base		prezzo base dell'articolo
'	qta:			quantita' in ordine
'	classe_sconto	id della classe di sconto applicata all'articolo
'.................................................................................................
function GetPrezzoUnitario(conn, rs, prezzo_base, qta, classe_sconto)
	dim sql
	if cInteger(classe_sconto)>0 then
		sql = "SELECT sco_sconto FROM gtb_scontiQ WHERE sco_classe_id = " & classe_sconto & _
			  " AND sco_qta_da <= " & qta & " ORDER BY sco_qta_da DESC "
		rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		if not rs.eof then
			GetPrezzoUnitario = GetPricePercent(prezzo_base, rs("sco_sconto"))
		else
			GetPrezzoUnitario = prezzo_base
		end if
		rs.close
	else
		GetPrezzoUnitario = prezzo_base
	end if
	GetPrezzoUnitario = ArrotondaEuro(GetPrezzoUnitario)
end function


'.................................................................................................
'	funzione che ritorna la condizione che filtra i prezzi della vista "gv_listino_vendita"
'	sulla base del listino del rivenditore
'.................................................................................................
function RivenditoreListinoCondition(listino_id, listino_base_attuale)
	if listino_base_attuale then
		RivenditoreListinoCondition = "( " + _
									  "	ISNULL(listino_offerte, 0)=1 " + _
									  " OR " + _
									  " ( " + _
									  " 	prz_variante_id NOT IN (SELECT prz_variante_id FROM gv_listino_offerte) " + _
									  "		AND " + _
									  "		prz_listino_id = " & cIntero(listino_id) & _
									  " )" + _
									  ")"
	else
		RivenditoreListinoCondition = "( " + _
									  "	ISNULL(listino_offerte, 0)=1 " + _
									  "	OR " + _
									  "	( " +_
									  " 	prz_variante_id NOT IN (SELECT prz_variante_id FROM gv_listino_offerte) " + _
									  "		AND " + _
									  "		( " + _
									  "			prz_listino_id = " & cIntero(listino_id) & _
									  "			OR " +_
									  "			( " + _
									  "				ISNULL(listino_base_attuale, 0)=1 " + _
									  "				AND " + _
									  "				prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_prezzi WHERE prz_listino_id=" & cIntero(listino_id) & ") " + _
									  "			) " + _
									  "		) " + _
									  " ) " + _
									  ") "
	end if
end function

'.................................................................................................
'	funzione che ritorna la condizione che filtra i prezzi della vista "gv_listino_vendita"
'	sulla base del listino del rivenditore
'.................................................................................................
function RivenditoreListinoConditionOptimized(listino_id, listino_base_attuale ,condition)
	if listino_base_attuale then
		RivenditoreListinoConditionOptimized = "( " + _
									  "	ISNULL(listino_offerte, 0)=1 " + _
									  " OR " + _
									  " ( " + _
									  " 	prz_variante_id NOT IN (SELECT prz_variante_id FROM gv_listino_offerte " + IIF(condition<>"", " WHERE " + condition, "") + ") " + _
									  "		AND " + _
									  "		prz_listino_id = " & cIntero(listino_id) & _
									  " )" + _
									  ")"
	else
		RivenditoreListinoConditionOptimized = "( " + _
									  "	ISNULL(listino_offerte, 0)=1 " + _
									  "	OR " + _
									  "	( " +_
									  " 	prz_variante_id NOT IN (SELECT prz_variante_id FROM gv_listino_offerte " + IIF(condition<>"", " WHERE " + condition, "") + ") " + _
									  "		AND " + _
									  "		( " + _
									  "			prz_listino_id = " & cIntero(listino_id) & _
									  "			OR " +_
									  "			( " + _
									  "				ISNULL(listino_base_attuale, 0)=1 " + _
									  "				AND " + _
									  "				prz_variante_id NOT IN (SELECT rel_id FROM gv_listini WHERE prz_listino_id=" & cIntero(listino_id) & " " + IIF(condition<>"", " AND " + condition, "") + ") " + _
									  "			) " + _
									  "		) " + _
									  " ) " + _
									  ") "
	end if
end function


'..................................................................................................
'recupera prezzo di listino attuale
'..................................................................................................
function GetPrezzoListinoAttuale(conn, rs, listino_id, rel_id)
	dim sql
	sql = "SELECT COUNT(*) FROM gtb_listini WHERE ISNULL(listino_base_attuale,0)=1 AND listino_id=" & cIntero(listino_id)
	GetPrezzoListinoAttuale = GetPrezzoListinoAttualeOptimized(conn, rs, listino_id, rel_id, (cInteger(GetValueList(conn, rs, sql))>0))
end function

function GetPrezzoListinoAttualeOptimized(conn, rs, listino_id, rel_id, listino_base_attuale)
	dim sql
	sql = "SELECT COUNT(*) FROM gtb_listini WHERE ISNULL(listino_base_attuale,0)=1 AND listino_id=" & cIntero(listino_id)
	sql = " SELECT prz_prezzo FROM gv_listino_vendita " + _
		  " WHERE rel_id =" & cIntero(rel_id) & " AND " + RivenditoreListinoCondition(listino_id, listino_base_attuale)
	GetPrezzoListinoAttualeOptimized = cReal(GetValueList(conn, rs, sql))
end function



'..................................................................................................
'recupera prezzo di listino base attuale
'..................................................................................................
function GetPrezzoListinoBase(conn, rs, rel_id)
	dim sql
	sql = " SELECT listino_id FROM gtb_listini WHERE listino_base_attuale=1"
	
	sql = " SELECT prz_prezzo FROM gv_listini " + _
		  " WHERE rel_id=" & cIntero(rel_id) & " AND prz_listino_id=" & GetValueList(conn, rs, sql)
	GetPrezzoListinoBase = cReal(GetValueList(conn, rs, sql))
end function


'.................................................................................................
'..			aggiorna il ranking degli articoli ordinati rispetto al rivenditore e agli eventuali sotto utenti
'..			conn			aperta sul database
'..			ordID			ID dell'ordine
'.................................................................................................
Function RankingOrdine(conn, ordID)
	dim rs, sql, rnk, rivID, detID, aux
	sql = " SELECT * FROM gtb_dettagli_ord d LEFT JOIN gtb_dettagli_ord_utenti u ON d.det_id=u.du_det_id "& _
		  " WHERE det_ord_id="& cIntero(ordID) & _
		  " ORDER BY du_ut_id"
	set rs = server.createobject("ADODB.recordset")
	set rnk = server.createobject("ADODB.recordset")
	set aux = server.createobject("ADODB.recordset")
	rivID = getValueList(conn, aux, "SELECT ord_riv_id FROM gtb_ordini WHERE ord_id="& cIntero(ordID))
	rs.open sql, conn, adOpenStatic, adLockOptimistic
	sql = "SELECT * FROM gtb_articoli_ordinati WHERE ao_ut_id="
	
	detID = 0
	while not rs.eof
		if detID <> rs("det_id") then
			'gestione rank rivenditore
			rnk.open sql & rivID &" AND ao_variante_id="& rs("det_art_var_id"), conn, adOpenDynamic, adLockOptimistic
			if rnk.eof then
				rnk.addNew
			end if
				rnk("ao_ut_id") = rivID
				rnk("ao_variante_id") = rs("det_art_var_id")
				rnk("ao_ranking") = getValueList(conn, aux, " SELECT SUM(det_qta) FROM gtb_dettagli_ord d "& _
														    " INNER JOIN gtb_ordini o ON d.det_ord_id=o.ord_id "& _
															" WHERE det_art_var_id="& rs("det_art_var_id") &" AND ord_riv_id="& cIntero(rivID))
			rnk.update
			rnk.close
			detID = rs("det_id")
		end if
		
		if not isNull(rs("du_ut_id")) then
			'gestione rank utente
			rnk.open sql & rs("du_ut_id") &" AND ao_variante_id="& rs("det_art_var_id"), conn, adOpenDynamic, adLockOptimistic
			if rnk.eof then
				rnk.addNew
			end if
				rnk("ao_ut_id") = rs("du_ut_id")
				rnk("ao_variante_id") = rs("det_art_var_id")
				rnk("ao_ranking") = getValueList(conn, aux, " SELECT SUM(du_qta) FROM gtb_dettagli_ord d "& _
														    " INNER JOIN gtb_dettagli_ord_utenti u ON d.det_id=u.du_det_id "& _
															" WHERE det_art_var_id="& rs("det_art_var_id") &" AND du_ut_id="& rs("du_ut_id"))
			rnk.update
			rnk.close
		end if
		
		rs.moveNext
	wend
	
	rs.close
	set rs = nothing
	set rnk = nothing
	set aux = nothing
End Function



'.................................................................................................
'..			visualizza la lista collegamenti ed articoli collegati
'..			conn			aperta sul database
'..			art_id 			id dell0'articolo&ugrave;
'..			Show_Collegati	indica se deve mostrare gli articoli a lui collegati
'..			Show_CollegatoA indica se deve mostrare gli articoli a cui e' collegato
'.................................................................................................
Sub ListaCollegamentiArticolo(conn, rs, art_id, Show_Collegati, Show_CollegatoA)
	if Show_Collegati then 
		sql = " SELECT at_nome_it, at_vincolo_vendita, aa_ordine, art_cod_int, art_nome_it, art_noVenSingola, art_varianti, art_id FROM gtb_articoli " + _ 
			  " INNER JOIN grel_art_acc ON gtb_articoli.art_id = grel_Art_acc.aa_acc_id " + _
			  " INNER JOIN gtb_accessori_tipo ON grel_art_acc.aa_tipo_id = gtb_accessori_tipo.at_id " + _
			  " WHERE grel_art_acc.aa_art_id=" & cIntero(art_id) & " ORDER BY gtb_accessori_tipo.at_ordine, grel_Art_acc.aa_ordine"
		rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		if not rs.eof then%>
			<tr>
				<td class="label">articoli collegati:</td>
				<td colspan="6">
					<% if rs.recordcount>2 then %> 
						<span class="overflow">
					<% end if  %>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr>
							<th class="L2" style="width:20%;">tipo</th>
							<th class="L2" style="width:15%;">codice</th>
							<th class="L2" style="width:37%;">nome</th>
							<th class="l2_center" style="width:20%;">non vend. sing.</th>
							<th class="l2_center" style="width:8%;">ordine</th>
						</tr>
						<% while not rs.eof %>
							<tr>
								<td class="content">
									<%= rs("at_nome_it") %>
									<% if rs("at_vincolo_vendita") then %>(vincolo vendita)<% end if %>
								</td>
								<td class="content"><%= rs("art_cod_int")%></td>
								<td class="content"><%CALL ArticoloLink(rs("art_id"), rs("art_nome_it"), rs("art_cod_int")) %></td>
								<td class="content_center">
									<% if rs("at_vincolo_vendita") then %>
										<input type="checkbox" class="checkbox" disabled <%= chk(rs("art_noVenSingola")) %>>
									<% else %>
										&nbsp;
									<% end if %>
								</td>
								<td class="content_center"><%= rs("aa_ordine")%></td>
							</tr>
							<%rs.movenext
						wend %>
					</table>
					<% if rs.recordcount>2 then %>
						</span>
					<% end if %>
				</td>
			</tr>
		<% end if
		rs.close
	end if
	if Show_CollegatoA then 
		sql = " SELECT at_nome_it, aa_ordine, art_cod_int, art_nome_it, art_noVenSingola, art_varianti, art_id FROM gtb_articoli " + _ 
			  " INNER JOIN grel_art_acc ON gtb_articoli.art_id = grel_Art_acc.aa_art_id " + _
			  " INNER JOIN gtb_accessori_tipo ON grel_art_acc.aa_tipo_id = gtb_accessori_tipo.at_id " + _
			  " WHERE grel_art_acc.aa_acc_id=" & cIntero(art_id) & " ORDER BY gtb_accessori_tipo.at_ordine, grel_Art_acc.aa_ordine"
		rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		if not rs.eof then%>
			<tr>
				<td class="label">collegato a:</td>
				<td colspan="6">
					<% if rs.recordcount>2 then %> 
						<span class="overflow">
					<% end if  %>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
						<tr>
							<th class="L2" width="20%">tipo</th>
							<th class="L2" width="15%">codice</th>
							<th class="L2">nome</th>
							<th class="l2_center" width="8%">ordine</th>
						</tr>
						<% while not rs.eof %>
							<tr>
								<td class="content"><%= rs("at_nome_it") %></td>
								<td class="content"><%= rs("art_cod_int")%></td>
								<td class="content"><%CALL ArticoloLink(rs("art_id"), rs("art_nome_it"), rs("art_cod_int")) %></td>
								<td class="content_center"><%= rs("aa_ordine")%></td>
							</tr>
							<%rs.movenext
						wend %>
					</table>
					<% if rs.recordcount>2 then %>
						</span>
					<% end if %>
				</td>
			</tr>
		<% end if
		rs.close
	end if
end sub



'.................................................................................................
'..			genera input per selezione articoli tramite elenco con ricerca su nuova finestra
'..			conn					aperta sul database
'..			rs						oggetto ausiliario
'..			form					nome del form che contiene gli input
'..			FieldName				nome del campo dove salvare i dati dell'articolo
'..			FieldValue				eventuale articolo selezionato
'..			Size					dimensione dell'input che contiene il valore selezionato
'..			Obbligatorio			visualizza o meno il pulsante di reset.
'..			AlternativePath 		percorso del file da aprire (per l'elenco degli aricoli) - NON OBBLIGATORIO
'..			--- SubmitAfterSelection	esegui il submit del form dopo la selezione dell'articolo
'.................................................................................................
sub WritePicker_ArticoloVariante(conn, rs, form, FieldName, FieldValue, Size, Obbligatorio, AlternativePath)
	dim DisplayValue, DisplayName, sql
	DisplayName = "view_" & FieldName
	if cInteger(FieldValue)>0 then
		sql = "Select art_nome_it, art_varianti FROM gv_Articoli WHERE rel_id=" & cIntero(FieldValue)
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		DisplayValue = rs("art_nome_it")
		if rs("art_Varianti") then
			rs.close
			sql = "SELECT var_nome_it + ': ' + val_nome_it FROM gv_Articoli_Varianti WHERE rvv_Art_var_id=" & cIntero(FieldValue)
			DisplayValue = DisplayValue & GetValueList(conn, rs, sql)
		else
			rs.close
		end if
	end if 
	
	if Trim(cString(AlternativePath)) = "" then
		AlternativePath = "nextb2b/ArticoliSelezionaVariante.asp?"
	end if
	
	%>
	<script language="JavaScript" type="text/javascript">
		function onclick_scegli_<%= DisplayName %>(){
			if (!<%= Form %>.<%= DisplayName %>.disabled){
				OpenAutoPositionedScrollWindow('<%= GetAmministrazionePath() %><%=AlternativePath%>formname=<%= form %>&inputname=<%= FieldName %>&selected=' + <%= Form %>.<%= FieldName %>.value, 'selezione_variante_articolo', 500, 400, true)
			}
		}
		function onclick_reset_<%= DisplayName %>(){
			if (!<%= Form %>.<%= DisplayName %>.disabled){
				<%= Form %>.<%= FieldName %>.value=''
				<%= Form %>.<%= DisplayName %>.value=''
			}
		}
	</script>
	<table cellpadding="0" cellspacing="0">
		<input type="hidden" name="<%= FieldName %>" value="<%= FieldValue %>">
		<tr>
			<td style="padding-top:2px;">
				<input READONLY type="text" name="<%= DisplayName %>" value="<%= DisplayValue %>" style="padding-left:3px;" size="<%= Size %>" title="Apre l'elenco degli articoli per la selezione" onclick="onclick_scegli_<%= DisplayName %>()">
			</td>
			<td style="padding-top:1px;" nowrap><a class="button_input" id="link_scegli_<%= FieldName %>" href="javascript:void(0)" onclick="<%= Form %>.<%= DisplayName %>.onclick();" title="Apre l'elenco degli articoli per la selezione." <%= ACTIVE_STATUS %>>SCEGLI</a></td>
			<% if not Obbligatorio then %>
				<td style="padding-top:1px;"><a class="button_input" id="link_reset_<%= FieldName %>" style="border-left:0px;" href="javascript:void(0)" onclick="onclick_reset_<%= DisplayName %>();" title="cancella la selezione eseguita" <%= ACTIVE_STATUS %>>RESET</a></td>
			<% else %>
				<td style="padding-top:1px;">&nbsp;(*)</td>
			<% end if %>
		</tr>
	</table>
	<%
end sub


'sovrascrive prezzi correnti del listino con le variazioni di default, impostando anche quelli mancanti dal listino base attuale
sub UpdateListinoFromVariazioniDefault(conn, listino_id, impostaErrori)
	dim rs
	
	'recupera dati di appoggio
	sql = " SELECT listino_Default_var_euro, listino_Default_var_sconto, " + _
		  " (SELECT listino_id FROM gtb_listini WHERE IsNull(listino_base_attuale, 0)=1) AS listino_base_attuale_id " + _
		  " FROM gtb_listini WHERE listino_id="& listino_id
	set rs = conn.Execute(sql)
	
	if not IsNull(rs("listino_default_var_euro")) OR not IsNull(rs("listino_default_var_sconto")) then
		
		'imposta prezzi dal listino base per quelli esistenti
		sql = " UPDATE gtb_prezzi SET " + _
					 " prz_var_euro = " & ParseSQL(rs("listino_default_var_euro"), adNumeric) & ", " + _
					 " prz_var_sconto = " & ParseSQL(rs("listino_default_var_sconto"), adNumeric) & ", " + _
					 " prz_prezzo = (SELECT prz_prezzo FROM gtb_prezzi base WHERE base.prz_variante_id = gtb_prezzi.prz_variante_id AND base.prz_listino_id=" & rs("listino_base_attuale_id") & ") " + _
			  " WHERE prz_listino_id = " & listino_id
		CALL conn.execute(sql, , adexecuteNoRecords)
		
		'imposta prezzi dal listino base per quelli non presenti
		sql = " INSERT INTO gtb_prezzi (prz_iva_id, prz_prezzo, prz_var_sconto, prz_var_euro, prz_visibile, " + _
			  " 	prz_promozione, prz_scontoQ_id, prz_listino_id, prz_variante_id ) " + _
			  "		SELECT prz_iva_id, prz_prezzo, " & ParseSQL(rs("listino_default_var_sconto"), adNumeric) & ", " & ParseSQL(rs("listino_default_var_euro"), adNumeric) & ", " & _
			  "			   1, 0, prz_scontoQ_id, " & listino_id & ", prz_variante_id " + _
			  "		FROM gtb_prezzi WHERE prz_listino_id = " & rs("listino_base_attuale_id") & " AND prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_prezzi WHERE prz_listino_id = " & listino_id & ") "
		CALL conn.execute(sql, , adexecuteNoRecords)
		
		'aggiorna prezzi base appena impostati applicando le variazioni
		sql = " UPDATE gtb_prezzi SET prz_prezzo = (prz_prezzo + " & ParseSQL(rs("listino_default_var_euro"), adNumeric) & " + (" & ParseSQL(rs("listino_default_var_sconto") / 100, adNumeric) & " * prz_prezzo) ) " + _
			  " WHERE prz_listino_id = " & listino_id
		CALL conn.execute(sql, , adexecuteNoRecords)
	
	else
		if impostaErrori then
			Session("ERRORE") = "Variazioni di default non impostate."
		end if
	end if
	set rs = nothing
	
end sub

'Calcola il codice rivenditore a partire dalla maschera e dall'idUtente e lo inserisce in db
sub GeneraCodiceRivenditore(conn,idUtente)
	dim mask
	mask = Session("maschera_codice_cl")
	dim cod
	cod = Left(mask,Len(mask)-Len(cString(idUtente)))
	cod=cod & cString(idUtente)
	'aggiorna prezzi base appena impostati applicando le variazioni
	dim sql 
	sql = " UPDATE gtb_rivenditori SET riv_codice = '" & cod & "'" + _
		" WHERE riv_id = " & idUtente
	CALL conn.execute(sql, , adexecuteNoRecords)
	
	sql = " UPDATE tb_Indirizzario SET PraticaPrefisso='" & cod & "' " + _
		  " WHERE IDElencoIndirizzi IN (SELECT ut_NextCom_ID FROM tb_Utenti where ut_ID=" & idUtente & ")"
	CALL conn.execute(sql, , adexecuteNoRecords)
	
end sub


'.................................................................................................
'	ritorna la classe content_disabled se il campo ha valore 0
'	value:			campo di cui controllare il valore
'.................................................................................................
function StiliCampoTestoAZero(value)
	if cReal(value) = 0 then
		StiliCampoTestoAZero = "content_disabled"
	end if
end function


'.................................................................................................
'funzione verifica se il codice dell'articolo &egrave; univoco
'		art_id 		id dell'articolo
'		cod			codice da verificare
'.................................................................................................
function CodeArtIsUnique(conn, rs, art_id, cod)
	dim sql
	sql = " SELECT COUNT(*) FROM gtb_articoli WHERE art_cod_int LIKE '" & ParseSQL(cod, adChar) & "' "
	if cIntero(art_id)>0 then
		sql = sql & " AND art_id<>" & art_id
	end if
	CodeArtIsUnique = (GetValueList(conn, rs, sql)=0)
end function
	
	
	

'.................................................................................................
'aggiorna log degli acquisti cliente con data di ultimo acquisto
'	conn	connessione a db aperta
'	riv_id	id rivenditore da aggiornare
'	sql_condition	eventuale condizione su articoli/ordini da aggiungere al filtro di caricamento dati.
'.................................................................................................
sub AggiornaAcquistiCliente(conn, riv_id, sql_condition)
	
	dim sql, baseQuery
	
	sql = "DELETE FROM glog_rivenditori_acquisti "  
	if cIntero(riv_id)>0 then
		sql = sql & " WHERE rac_rivenditore_id = " & riv_id
	end if
	
	baseQuery = " 	FROM gtb_ordini INNER JOIN gtb_dettagli_ord ON gtb_ordini.ord_id = gtb_dettagli_ord.det_ord_id " + _
		  "			 INNER JOIN grel_art_valori ON gtb_dettagli_ord.det_art_var_id = grel_art_valori.rel_id  " + _
		  "		WHERE IsNull(det_art_var_id ,0)<>0 "
	if cstring(sql_condition)<>"" then
		baseQuery = baseQuery & sql_condition
	end if
	if cIntero(riv_id)>0 then
		  baseQuery = baseQuery & " AND ord_riv_id = " & riv_id
	end if
	
	sql = sql & _
		  "INSERT INTO [dbo].[glog_rivenditori_acquisti] ([rac_data_ultimo_acquisto], [rac_rivenditore_id], [rac_riv_sede_id], [rac_art_var_id], rac_art_id) " + _
		  "SELECT MAX(ord_data), ord_riv_id, 0, det_art_var_id, rel_art_id "  + _
		  baseQuery + _
		  " GROUP BY ord_riv_id, det_art_var_id, rel_art_id " + _
		  "INSERT INTO [dbo].[glog_rivenditori_acquisti] ([rac_data_ultimo_acquisto], [rac_rivenditore_id], [rac_riv_sede_id], [rac_art_var_id], rac_art_id) " + _
	      "SELECT MAX(ord_data), ord_riv_id, det_ind_id, det_art_var_id, rel_art_id "  + _
		  baseQuery + _
		  " GROUP BY ord_riv_id, det_ind_id, det_art_var_id, rel_art_id "
		  
	CALL conn.execute(sql)
	
end sub
%>