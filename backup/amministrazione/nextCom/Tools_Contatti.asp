<%
'.................................................................................................
'.................................................................................................
'COSTANTI
'.................................................................................................
'.................................................................................................

'colori aree
const COLORE_CONTATTI_MACCHINE = "#ddf1ff"
const COLORE_CONTATTI_TRATTATIVE = "#b4dffd"
const COLORE_CONTATTI_ATTIVITA = "#def0de"
const COLORE_CONTATTI_CAMPAGNE = "#b9ddb9"


'tipi di access list diponibili
const AL_DEFAULT = "DEFAULT"
const AL_PRATICHE = "PRATICHE"
const AL_ATTIVITA = "ATTIVITA"
const AL_DOCUMENTI = "DOCUMENTI"


dim CatContatti
set CatContatti = New objCategorie
with CatContatti
	.tabella = "tb_indirizzario_categorie"
	.prefisso = "icat"
	.PrefissoPagine = "Contatti"
	
    'gestione delle relazioni con le tabelle categorizzate
    CALL .AddRelazione("tb_indirizzario", "IDElencoIndirizzi", "cnt_categoria_id", true)
	'gestione delle eventuali caratteristiche tecniche (descrittori)
	.tabellaCaratteristiche = "tb_indirizzario_carattech"
	.idCaratteristiche = "ict_id"
	.nomeCaratteristiche = "ict_nome_it"
	.tipoCaratteristiche = "ict_tipo"
	
	'gestione della eventuale relazione con caratteristiche tecniche
	.tabellaRelCaratteristiche = "rel_categ_ctech"
	.chiaveEsternaRelCaratteristiche = "rcc_categoria_id"
	.ordineRelCaratteristiche = "rcc_ordine"
	.idCarRelCaratteristiche = "rcc_ctech_id"
	'gestione della eventuale relazione tra tabella correlata e caratteristiche tecniche
	.tabellaRelCorCaratteristiche = "rel_cnt_ctech"
	.idArtRelCorCaratteristiche = "ric_cnt_id"
	.idCarRelCorCaratteristiche = "ric_ctech_id"
	
	.abilitaFoto = false
	.abilitaLogo = false
	.abilitaDescrittori = session("NEXTCOM_ATTIVA_GESTIONE_CATEGORIE")
	.GestioneCategorieMiste = false
	.categorieBloccate = false
	.attivaCKEditorPerDescrizione = false
	
	'abilitazione indice e contenuti
	.Index = Index
end with


'.................................................................................................
'.................................................................................................
'FUNZIONI E PROCEDURE
'.................................................................................................
'.................................................................................................


'.................................................................................................
'.. scrive l'icona della newsletter se il parametro è true
'.................................................................................................
function write_icona_newsletter(isNewsletter)
	response.write get_icona_newsletter(isNewsletter)
end function


'.................................................................................................
'.. restituisce l'icona della newsletter se il parametro è true
'.................................................................................................
function get_icona_newsletter(isNewsletter)
	if isNewsletter then 
		get_icona_newsletter = "<img src=""" + GetAmministrazionePath() + "grafica/i.p.new.gif"" title=""Indirizzo email utilizzato nella spedizione delle newsletters."" alt=""Indirizzo email utilizzato nella spedizione delle newsletters.""/>"
	else
		get_icona_newsletter = ""
	end if
	
end function


'.................................................................................................
'..			visualizza la parte di form per la gestione delle rubriche
'..			conn:			connessione al database aperta
'..			rs:				oggetto recordset chiuso
'..			CNT_ID			id del contatto
'.................................................................................................
function Write_Relazione_Rubriche(conn, rs, CntId)
	dim sql
	
	sql = "SELECT tb_rubriche.id_rubrica, tb_rubriche.nome_rubrica, rel_rub_ind.id_rub_ind " &_
		  " FROM tb_rubriche LEFT JOIN rel_rub_ind ON (tb_rubriche.id_Rubrica = rel_rub_ind.id_rubrica " &_
		  " AND rel_rub_ind.id_indirizzo=" & cIntero(CntId) & ")" & _
		  " WHERE tb_rubriche.id_rubrica IN (" & GetList_Rubriche(conn, rs) & ")" &_
		  " ORDER BY nome_rubrica"
	CALL Write_Relations_Checker(conn, rs, sql, 4, "id_rubrica", "nome_rubrica", "id_rub_ind", "rubriche")
end function


'.................................................................................................
'..			restituisce TRUE se ho i diritti sull'oggetto ID
'..			ID:				ID dell'oggetto
'..			conn:			connessione aperta
'..			tipo:			PRATICHE | DOCUMENTI | ATTIVITA è una stringa!!!
'..			Usa la variabile di sessione: Session("ID_ADMIN")
'.................................................................................................
Function AL(conn, ID, tipo)
	tipo = UCase(tipo)
	if Session("COM_ADMIN") <> "" OR ID = "" then
		AL = TRUE
	else
		dim rs, sql, nomeTab, prefisso
		'
		'controllo se e' pubblica 
		nomeTab = "tb_"& tipo
		prefisso = left(tipo, 3)
		sql = "SELECT "& prefisso &"_pubblica FROM "& nomeTab &" WHERE "& prefisso &"_id = "& cIntero(ID)
		set rs = conn.execute(sql)
		if rs(0) then
			AL = TRUE
		else
			if tipo=AL_ATTIVITA then
				sql = "SELECT "& prefisso &"_mittente_id FROM "& nomeTab &" WHERE "& prefisso &"_id = "& cIntero(ID)
			else
				sql = "SELECT "& prefisso &"_creatore_id FROM "& nomeTab &" WHERE "& prefisso &"_id = "& cIntero(ID)
			end if
			set rs = conn.execute(sql)
			if rs(0)=Session("ID_ADMIN") then
				AL = TRUE
			else
				'controllo se sono in AL utenti
				nomeTab = "al_"& tipo &"_utenti"
				sql = "SELECT COUNT(*) FROM "& nomeTab &" WHERE al_utente_id="& Session("ID_ADMIN") & _
					  " AND al_tipo_id="& cIntero(ID)
				set rs = conn.execute(sql)
				if rs(0) > 0 then
					AL = TRUE
				else
					'controllo se sono in AL gruppi
					nomeTab = "al_"& tipo &"_gruppi"
					sql = "SELECT COUNT(*) FROM "& nomeTab &" t INNER JOIN tb_rel_dipGruppi r ON "& _
						  "t.al_gruppo_id=r.id_gruppo "& _
						  "WHERE id_impiegato="& Session("ID_ADMIN") &" AND al_tipo_id="& cIntero(ID)
					set rs = conn.execute(sql)
					if rs(0) = 0 then
						AL = FALSE
					else
						AL = TRUE
					end if
				end if
			end if
		end if
		
		set rs = nothing
	end if
End Function

'.................................................................................................
'..			restituisce la query da concatenare per visualizzare solo gli oggetti visibili da AL (es.: pra_id IN (SELECT ...))
'..			conn:			connessione aperta
'..			tipo:			PRATICHE | DOCUMENTI | ATTIVITA
'.................................................................................................
Function AL_query(conn, tipo)
	tipo = UCase(tipo)
	if Session("COM_ADMIN") <> "" then
		AL_query = "(1=1)"
	else
		dim campoID, campoPubblica
		campoID = left(tipo, 3) &"_id"
		campoPubblica = left(tipo, 3) &"_pubblica"
		AL_query = " ("& campoID &" IN (SELECT al_tipo_id FROM al_"& tipo &"_utenti " & _
			  		   				   "WHERE al_utente_id="& Session("ID_ADMIN") &") OR " & _
				  		 campoID &" IN (SELECT al_tipo_id FROM al_"& tipo &"_gruppi t INNER JOIN " & _
			 						   "tb_rel_dipGruppi r ON t.al_gruppo_id=r.id_gruppo " & _
									   "WHERE id_impiegato="& Session("ID_ADMIN") &") OR " & _
						 SQL_IsTrue(conn, campoPubblica) &") "
	end if
End Function

'.................................................................................................
'..			scrive i checkbox dei gruppi e degli utenti con il javascript
'..			ID:				ID dell'oggetto
'..			conn:			connessione aperta
'..			tipo:			PRATICHE | DEFAULT | DOCUMENTI | ATTIVITA
'..			IL NOME DEL FORM DEVE ESSERE 'FORM1'
'.................................................................................................
Sub AL_disegna(conn, ID, tipo) 
	dim rs, rsu, sql, i, colspan 
	dim CanInherit, DefaultPratica, Prefisso, Eredita
	
	set rs = Server.CreateObject("ADODB.Recordset")
	set rsu = Server.CreateObject("ADODB.Recordset")
	ID = cInteger(ID)
	DefaultPratica = 0
	Prefisso = left(tipo, 3)
	if tipo <> AL_DEFAULT AND tipo <> AL_PRATICHE then
		if tipo = AL_ATTIVITA AND Session("ATT_PRA_ID")<>"" then	
			'sezione "attivita della pratica": l'attivita' puo' ereditare i permessi dalla pratica
			CanInherit = true
			DefaultPratica = Session("ATT_PRA_ID")
		elseif tipo = AL_DOCUMENTI AND Session("DOC_PRA_ID")<>"" then
			'sezione "documenti della pratica" il documento puo' ereditare i permessi dalla pratica
			CanInherit = true
			DefaultPratica = Session("DOC_PRA_ID")
		else
			if ID = 0 then
				'non puo' ereditare perche' non si conosce ancora al pratica a cui e' associato
				CanInherit = false
			else
				'recupera l'id della pratica: se presente puo' ereditare
				if tipo = AL_ATTIVITA then
					sql = "SELECT att_pratica_id FROM tb_attivita WHERE att_id=" & cIntero(ID)
				elseif tipo = AL_DOCUMENTI then
					sql = "SELECT doc_pratica_id FROM tb_documenti WHERE doc_id=" & cIntero(ID)
				end if
				rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
				DefaultPratica = cInteger(rs(0))
				CanInherit = (DefaultPratica>0)
				rs.close
			end if
		end if
	else	
		'access list di pratica e di default non possono ereditare
		CanInherit = false
	end if
					
	sql = "SELECT DISTINCT g.* FROM ((tb_gruppi g INNER JOIN tb_rel_dipGruppi r ON "& _
		  "g.id_gruppo=r.id_gruppo) INNER JOIN tb_admin a ON "& _
		  "r.id_impiegato=a.id_admin) INNER JOIN rel_admin_sito rel ON "& _
		  "a.id_admin=rel.admin_id "& _
		  "WHERE sito_id="& NEXTCOM & _
		  " ORDER BY nome_gruppo"
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	colspan = 5 %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
		<tr>
			<td class="label" style="width:30%;">
				Elenco dei gruppi e degli utenti
			</td>
			<td class="content_right" style="font-size: 1px; padding-right:1px;" colspan="<%= colspan-1 %>">
				<a id="tutti" class="button_L2" href="javascript:void(0);" onclick="Tutti()" title="seleziona tutti gli utenti ed i gruppi sotto elencati" <%= ACTIVE_STATUS %>>
					SELEZIONA TUTTI
				</a>
				&nbsp;
				<a id="nessuno" class="button_L2" href="javascript:void(0);" onclick="Reset()" title="toglie la selezione a tutti gli utenti ed i gruppi sotto elencati" <%= ACTIVE_STATUS %>>
					DESELEZIONA TUTTI
				</a>
				<% If CanInherit then %>
					&nbsp;
					<a id="eredita" class="button_L2" href="javascript:void(0);" onclick="Eredita(true)" title="Imposta i permessi ai permessi della pratica di appartenenza" <%= ACTIVE_STATUS %>>
						RESET A DEFAULT
					</a>
					<input type="hidden" name="ere" value="">
				<% Else  %>
					<input type="hidden" id="eredita">
					<input type="hidden" name="ere" value="">
				<% End If %>
			</td>
		</tr>
		<tr>
			<th class="L2">GRUPPI</th>
			<th class="L2" colspan="<%= colspan-1 %>">UTENTI</th>
		</tr>
		<tr>
			<% while not rs.eof %>
				<td class="content_b">
					<input onclick="ClickGruppo(this.checked, <%= rs("id_gruppo") %>)" type="Checkbox" id="grp_<%= rs("id_gruppo") %>" name="grp_<%= rs("id_gruppo") %>" <%= Chk(request("grp_"& rs("id_gruppo"))<>"") %> class="checkbox" value="<%= rs("id_gruppo") %>">
					<%= rs("nome_gruppo") %>
				</td>
				<%sql = "SELECT DISTINCT admin_cognome, admin_nome, id_admin " & _
						" FROM (tb_admin a INNER JOIN tb_rel_dipGruppi r ON a.id_admin=r.id_impiegato) " & _
						" INNER JOIN rel_admin_sito rel ON a.id_admin=rel.admin_id "& _
					    " WHERE id_gruppo="& rs("id_gruppo") &" AND sito_id="& NEXTCOM & _
						" ORDER BY admin_cognome"
				rsu.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
				i = 1
				while i < colspan
					if not rsu.eof AND i > 0 then%>
						<td class="content">
							<input onclick="ClickUtente(this.checked, 'adm_<%= rsu("id_admin") %>')" type="Checkbox" id="adm_<%= rsu("id_admin") %>" name="adm_<%= rsu("id_admin") %>" <%= Chk(request("adm_"& rsu("id_admin"))<>"") %> class="checkbox" value="<%= rs("id_gruppo") %>">
							<%= rsu("admin_cognome") &" "& Left(rsu("admin_nome"), 1) &"." %>
						</td>
						<%rsu.movenext
					else%>
						<td class="content">&nbsp;</td>
					<%end if
					i = i+1
					if i = colspan then
						response.write "</tr><tr>"
						i = IIF(rsu.eof, i, 0)
					end if
				wend
				rsu.close
				rs.movenext
			wend%>
		</tr>
	</table>
	<% rs.close %>
					
	<script language="javascript">

		function Tutti() {
			Eredita(false)
			for(var i=0; i < form1.elements.length; i++)
				if (form1.elements(i).id.substring(0, 4) == "grp_")
					ClickGruppo(true, form1.elements(i).id.substring(4, form1.elements(i).id.length))
		}

		function Reset() {
			Eredita(false)
			for(var i=0; i < form1.elements.length; i++)
				if (form1.elements(i).id.substring(0, 4) == "grp_")
					ClickGruppo(false, form1.elements(i).id.substring(4, form1.elements(i).id.length))
				else if (form1.elements(i).id.substring(0, 4) == "adm_")
					ClickUtente(false, form1.elements(i).id)
		}
				
		function ClickUtente(chk, id) {
			var campo
			Eredita(false)
			eval("campo = form1."+ id)
			if (!campo.length) {
				campo.checked = chk
				if (!campo.checked)
						eval("form1.grp_"+ campo.value +".checked = false")
			} else
				for(var i=0; i < campo.length; i++) {
					campo(i).checked = chk
					if (!campo(i).checked)
						eval("form1.grp_"+ campo(i).value +".checked = false")
				}
		}
				
		function ClickGruppo(chk, id) {
			var campo
			Eredita(false)
			eval("form1.grp_"+ id +".checked = "+ chk)
			for(var i=0; i < form1.length; i++) {
				campo = form1(i)
				if (campo.id.substring(0, 4) == "adm_")
					if (campo.value == id)
						ClickUtente(true, campo.id)
			}
		}
		
		<%'eventuale caricamento permessi di default
		if CanInherit AND DefaultPratica<>0 then
			'genera funzione associata al pulsante per impostare i permessi ereditati%>
			function Eredita(si) {
				if (si) {
					Reset()
					<%'imposta utenti abilitati
					sql = "SELECT * FROM al_default_utenti WHERE al_tipo_id="& DefaultPratica
					rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
					while not rs.eof %>
						ClickUtente(true, 'adm_<%= rs("al_utente_id") %>')
						<%rs.movenext
					wend
					rs.close
					
					'imposta gruppi abilitati
					sql = "SELECT * FROM al_default_gruppi WHERE al_tipo_id="& DefaultPratica
					rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
					while not rs.eof %>
						ClickGruppo(true, <%= rs("al_gruppo_id") %>)
						<%rs.movenext
					wend
					rs.close%>
					form1.ere.value = "true"
					document.all.eredita.disabled = true;
					document.all.eredita.className = "button_L2_disabled"
				} else{
					form1.ere.value = ""
					document.all.eredita.disabled = false;
					document.all.eredita.className = "button_L2"
				}
			}
			
			<%if ID > 0 then
				'caricamento permessi ereditati in modifica se l'elemento eredita
				sql = "SELECT " & prefisso &"_eredita FROM tb_"& tipo &" WHERE "& prefisso &"_id="& cIntero(ID)
				if conn.execute(sql)(prefisso &"_eredita").value OR request.form("ere")<>"" then
					Eredita = true%>
					Eredita(true);
				<%end if
			else
				Eredita = false
			end if
			
		else 
			'l'elemento non pu&ograve; ereditare:pratica o documento/attivita' non associato%>
			function Eredita(si) {
				return void(0)
				}
			<%
			Eredita = false
		end if
		
		if ID > 0 AND not Eredita then
			'carica permessi salvati se l'elemento non eredita
			
			sql = "SELECT * FROM al_"& tipo &"_utenti WHERE al_tipo_id="& cIntero(ID)
			rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
			while not rs.eof %>
				ClickUtente(true, 'adm_<%= rs("al_utente_id") %>')
				<%rs.movenext
			wend
			rs.close
			
			sql = "SELECT * FROM al_"& tipo &"_gruppi WHERE al_tipo_id="& cIntero(ID)
			rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
			while not rs.eof %>
				ClickGruppo(true, <%= rs("al_gruppo_id") %>)
				<%rs.movenext
			wend
			rs.close
			
		end if
		%>
	</script>
	<%
	set rs = nothing
	set rsu = nothing
End Sub


'.................................................................................................
'..			inserisce e modifica le AL nel DB
'..			conn:			connessione aperta
'..			tipo:			DOCUMENTI | ATTIVITA | DEFAULT | PRATICHE
'..			ID:				ID del record dell'oggetto
'..			eredita:		flag se l'oggetto eredita o meno da default
'..							puo' essere TRUE solo per i tipi <> da DEFAULT e PRATICHE
'.................................................................................................
Sub AL_ins(conn, tipo, ID, eredita)
	dim prefisso, rs, pratica, count
	set rs = server.createobject("adodb.recordset")
	tipo = UCase(tipo)
	prefisso = left(tipo, 3)
	
	if tipo <> AL_DEFAULT AND tipo <> AL_PRATICHE then
		'prendo l'id della pratica associata
		pratica = GetValueList(conn, rs, "SELECT "& prefisso &"_pratica_id FROM tb_"& tipo & _
										 " WHERE "& prefisso &"_id="& cIntero(ID))
	else
		pratica = ID
		prefisso = "PRA"
		eredita = false			'controllo sui tipi
	end if
	
	'cancello AL
	conn.execute("DELETE FROM al_"& tipo &"_gruppi WHERE al_tipo_id="& cIntero(ID))
	conn.execute("DELETE FROM al_"& tipo &"_utenti WHERE al_tipo_id="& cIntero(ID))

	if eredita then
		'inserimento ereditato
		conn.execute("INSERT INTO al_"& tipo &"_gruppi(al_tipo_id, al_gruppo_id) "& _
					 "SELECT "& cIntero(ID) &", al_gruppo_id FROM al_default_gruppi "& _
					 "WHERE al_tipo_id="& cIntero(pratica))
		conn.execute("INSERT INTO al_"& tipo &"_utenti(al_tipo_id, al_utente_id) "& _
					 "SELECT "& cIntero(ID) &", al_utente_id FROM al_default_utenti "& _
					 "WHERE al_tipo_id="& cIntero(pratica))
		'setto il flag
		conn.execute("UPDATE tb_"& tipo &" SET "& prefisso &"_eredita=1 "& _
					 "WHERE "& prefisso &"_id="& cIntero(ID))
	elseif tipo = AL_PRATICHE then
		'se c'è un doc o un'att pubblica la pratica resta pubblica else inserisco AL di doc e att
		if CInt(GetValueList(conn, rs, "SELECT COUNT(*) FROM tb_attivita WHERE " & _
			    SQL_IsTrue(conn, "att_pubblica") &" AND att_pratica_id="& cIntero(ID))) = 0 AND _
		   CInt(GetValueList(conn, rs, "SELECT COUNT(*) FROM tb_documenti WHERE " & _
			    SQL_IsTrue(conn, "doc_pubblica") &" AND doc_pratica_id="& cIntero(ID))) = 0 then
		
			'inserimento da attivita' e documenti attuali
			dim tempSQL
			if DB_Type(conn) = DB_SQL then
				tempSQL = "INSERT INTO al_pratiche_gruppi(al_tipo_id, al_gruppo_id) "& _
						  "(SELECT "& cIntero(ID) &", al_gruppo_id FROM al_documenti_gruppi r INNER JOIN "& _
						  "tb_documenti d ON r.al_tipo_id=d.doc_id WHERE doc_pratica_id="& cIntero(ID) &" UNION "& _
						  "SELECT "& cIntero(ID) &", al_gruppo_id FROM al_attivita_gruppi r INNER JOIN "& _
						  "tb_attivita a ON r.al_tipo_id=a.att_id WHERE att_pratica_id="& cIntero(ID) &")"
				conn.execute(tempSQL)
				'inserisce anche contatti gia presenti nei gruppi per semplicita e velocita d'exe
				tempSQL = "INSERT INTO al_pratiche_utenti(al_tipo_id, al_utente_id) "& _
						  "(SELECT "& cIntero(ID) &", al_utente_id FROM al_documenti_utenti r INNER JOIN "& _
						  "tb_documenti d ON r.al_tipo_id=d.doc_id WHERE doc_pratica_id="& cIntero(ID) &") UNION "& _
						  "(SELECT "& cIntero(ID) &", al_utente_id FROM al_attivita_utenti r INNER JOIN "& _
						  "tb_attivita a ON r.al_tipo_id=a.att_id WHERE att_pratica_id="& cIntero(ID) &")"
				conn.execute(tempSQL)
			else
				tempSQL = "(SELECT "& cIntero(ID) &", al_gruppo_id FROM al_documenti_gruppi r INNER JOIN "& _
						  "tb_documenti d ON r.al_tipo_id=d.doc_id WHERE doc_pratica_id="& cIntero(ID) &" UNION "& _
						  "SELECT "& cIntero(ID) &", al_gruppo_id FROM al_attivita_gruppi r INNER JOIN "& _
						  "tb_attivita a ON r.al_tipo_id=a.att_id WHERE att_pratica_id="& cIntero(ID) &")"
				rs.open tempSQL, conn
				while not rs.eof
					conn.execute("INSERT INTO al_pratiche_gruppi(al_tipo_id, al_gruppo_id) VALUES ("& _
								 rs(0) &", "& rs(1) &")")
					rs.movenext
				wend
				rs.close
				'inserisce anche contatti gia presenti nei gruppi per semplicita e velocita d'exe
				tempSQL = "(SELECT "& cIntero(ID) &", al_utente_id FROM al_documenti_utenti r INNER JOIN "& _
						  "tb_documenti d ON r.al_tipo_id=d.doc_id WHERE doc_pratica_id="& cIntero(ID) &") UNION "& _
						  "(SELECT "& cIntero(ID) &", al_utente_id FROM al_attivita_utenti r INNER JOIN "& _
						  "tb_attivita a ON r.al_tipo_id=a.att_id WHERE att_pratica_id="& cIntero(ID) &")"
				rs.open tempSQL, conn
				while not rs.eof
					conn.execute("INSERT INTO al_pratiche_utenti(al_tipo_id, al_utente_id) VALUES ("& _
								 rs(0) &", "& rs(1) &")")
					rs.movenext
				wend
				rs.close
			end if
			
		end if
	else
		'inserimento dal form
		dim campo, ins, grp, inGruppo
		for each campo in request.form
			if left(campo, 4) = "grp_" then
				ins = right(campo, len(campo)-4)
				conn.execute("INSERT INTO al_"& tipo &"_gruppi(al_tipo_id, al_gruppo_id) VALUES ("& _
							 ID &", "& ins &")")
			elseif left(campo, 4) = "adm_" then
				ins = right(campo, len(campo)-4)
				'controllo di non aver inserito un gruppo di appartenenza
				inGruppo = false
				for each grp in split(request.form(campo), ",")
					grp = Trim(grp)
					if request.form("grp_"& grp) <> "" then
						inGruppo = true
					end if
				next
				if NOT inGruppo then
					conn.execute("INSERT INTO al_"& tipo &"_utenti(al_tipo_id, al_utente_id) VALUES ("& _
								 ID &", "& ins &")")
				end if
			end if
		next
		
		'se non eredita setto il flag
		if tipo <> AL_DEFAULT then
			conn.execute("UPDATE tb_"& tipo &" SET "& prefisso &"_eredita=0 "& _
						 "WHERE "& prefisso &"_id="& ID)
		end if
	end if
	
	if tipo = AL_DEFAULT then
		'update delle AL degli oggetti che ereditano ottimizzata
		conn.execute("DELETE FROM al_documenti_gruppi WHERE al_tipo_id IN "& _
					 "(SELECT doc_id FROM tb_documenti WHERE "& SQL_IsTrue(conn, "doc_eredita")& _
					 " AND doc_pratica_id="& cIntero(ID) &")")
		conn.execute("DELETE FROM al_documenti_utenti WHERE al_tipo_id IN "& _
					 "(SELECT doc_id FROM tb_documenti WHERE "& SQL_IsTrue(conn, "doc_eredita")& _
					 " AND doc_pratica_id="& cIntero(ID) &")")
		conn.execute("DELETE FROM al_attivita_gruppi WHERE al_tipo_id IN "& _
					 "(SELECT att_id FROM tb_attivita WHERE "& SQL_IsTrue(conn, "att_eredita")& _
					 " AND att_pratica_id="& cIntero(ID) &")")
		conn.execute("DELETE FROM al_attivita_utenti WHERE al_tipo_id IN "& _
					 "(SELECT att_id FROM tb_attivita WHERE "& SQL_IsTrue(conn, "att_eredita")& _
					 " AND att_pratica_id="& cIntero(ID) &")")
		
		conn.execute("INSERT INTO al_documenti_gruppi(al_tipo_id, al_gruppo_id) "& _
					 "SELECT doc_id, al_gruppo_id FROM (tb_documenti d INNER JOIN tb_pratiche p "& _
					 "ON d.doc_pratica_id=p.pra_id) INNER JOIN al_default_gruppi a "& _
					 "ON p.pra_id=a.al_tipo_id "& _
					 "WHERE pra_id="& cIntero(ID) &" AND "& SQL_IsTrue(conn, "doc_eredita"))
		conn.execute("INSERT INTO al_documenti_utenti(al_tipo_id, al_utente_id) "& _
					 "SELECT doc_id, al_utente_id FROM (tb_documenti d INNER JOIN tb_pratiche p "& _
					 "ON d.doc_pratica_id=p.pra_id) INNER JOIN al_default_utenti a "& _
					 "ON p.pra_id=a.al_tipo_id "& _
					 "WHERE pra_id="& cIntero(ID) &" AND "& SQL_IsTrue(conn, "doc_eredita"))
		conn.execute("INSERT INTO al_attivita_gruppi(al_tipo_id, al_gruppo_id) "& _
					 "SELECT att_id, al_gruppo_id FROM (tb_attivita d INNER JOIN tb_pratiche p "& _
					 "ON d.att_pratica_id=p.pra_id) INNER JOIN al_default_gruppi a "& _
					 "ON p.pra_id=a.al_tipo_id "& _
					 "WHERE pra_id="& cIntero(ID) &" AND "& SQL_IsTrue(conn, "att_eredita"))
		conn.execute("INSERT INTO al_attivita_utenti(al_tipo_id, al_utente_id) "& _
					 "SELECT att_id, al_utente_id FROM (tb_attivita d INNER JOIN tb_pratiche p "& _
					 "ON d.att_pratica_id=p.pra_id) INNER JOIN al_default_utenti a "& _
					 "ON p.pra_id=a.al_tipo_id "& _
					 "WHERE pra_id="& cIntero(ID) &" AND "& SQL_IsTrue(conn, "att_eredita"))
					 
		'setto il flag di pubblico xogni oggetto che eredita
		if AL_pubblica(conn, rs, ID, tipo) then
			conn.execute("UPDATE tb_documenti SET doc_pubblica = 1 "& _
						 "WHERE doc_pratica_id="& cIntero(pratica))
			conn.execute("UPDATE tb_attivita SET att_pubblica = 1 "& _
						 "WHERE att_pratica_id="& cIntero(pratica))
		else
			conn.execute("UPDATE tb_documenti SET doc_pubblica = 0 "& _
						 "WHERE doc_pratica_id="& cIntero(pratica))
			conn.execute("UPDATE tb_attivita SET att_pubblica = 0 "& _
						 "WHERE att_pratica_id="& cIntero(pratica))
		end if
	else
		'setto l'eventuale flag di pubblico
		if AL_pubblica(conn, rs, ID, tipo) then
			conn.execute("UPDATE tb_"& tipo &" SET "& prefisso &"_pubblica=1 "& _
						 "WHERE "& prefisso &"_id="& cIntero(ID))
		else
			conn.execute("UPDATE tb_"& tipo &" SET "& prefisso &"_pubblica=0 "& _
						 "WHERE "& prefisso &"_id="& cIntero(ID))
		end if
	end if
	
	if tipo <> AL_PRATICHE AND CInteger(pratica) <> 0 then
		'update dei permessi della pratica
		CALL AL_ins(conn, AL_PRATICHE, pratica, false)
	end if
	'x acl di default e documenti scrivo la data di modifica negli appositi campi
	if tipo = AL_DOCUMENTI OR tipo = AL_DEFAULT then
		conn.execute("UPDATE tb_"& IIF(tipo = AL_DEFAULT, AL_PRATICHE, tipo) &" SET "& _
			         prefisso &"_mod_data="& SQL_Now(conn) &", "& _
					 prefisso &"_mod_utente="& Session("ID_ADMIN") & _
					 " WHERE "& prefisso &"_id="& ID)
	end if
End Sub

'.................................................................................................
'..			restituisce TRUE se l'oggetto e' pubblico
'..			conn:			connessione aperta
'..			rs:				recordset chiuso
'..			tipo:			PRATICHE | DOCUMENTI | ATTIVITA
'..			ID: 			ID del record dell'oggetto
'.................................................................................................
Function AL_pubblica(conn, rs, ID, tipo)
	dim count
	count = CInt(GetValueList(conn, rs, "SELECT COUNT(*) FROM al_"& tipo &"_gruppi "& _
								   		"WHERE al_tipo_id="& cIntero(ID)))
	if count = 0 then
		count = CInt(GetValueList(conn, rs, "SELECT COUNT(*) FROM al_"& tipo &"_utenti "& _
									   		"WHERE al_tipo_id="& cIntero(ID)))
	end if
	AL_pubblica = (count = 0)
End Function


'.................................................................................................
'..		scrive il nome del contatto linkabile se l'utente lo puo' vedere
'..		rs:		recordset aperto su un record dell'indirizzario contenente il contatto interessato
'.................................................................................................
function ContactLinkedName(byref rs)
	CALL ContactLinkedNameExtra(rs, true, "")
end function

'.................................................................................................
'..		scrive il nome del contatto linkabile se l'utente lo puo' vedere
'..		rs:		recordset aperto su un record dell'indirizzario contenente il contatto interessato
'..		openInANewWindow: se true apre in una nuova finestra
'.................................................................................................
function ContactLinkedNameExtra(byref rs, openInANewWindow, queryString)
	dim sql, visible
	
	if session("COM_ADMIN")<>"" then
		visible = true
	else	
		sql = " SELECT (COUNT(*)) AS VAL FROM (tb_rubriche INNER JOIN rel_rub_ind " & _
			  " ON tb_rubriche.id_Rubrica = rel_rub_ind.id_rubrica) " & _
			  " INNER JOIN tb_rel_gruppirubriche ON tb_rubriche.id_Rubrica = tb_rel_gruppirubriche.id_dellaRubrica " & _
			  " WHERE rel_rub_ind.id_indirizzo=" & rs("IDElencoIndirizzi") & _
			  " AND id_Gruppo_assegnato IN (" & Session("DIP_GROUP") & ")"
		if cInteger(rs.ActiveConnection.execute(sql)("VAL"))> 0 then
			visible = true
		else
			visible = false
		end if
	end if
	
	if visible then
		if openInANewWindow then
			%>
			<a href="javascript:void(0);" title="apri scheda del contatto n&ordm;<%= rs("IDElencoIndirizzi") %>" <%= ACTIVE_STATUS %>
				onclick="OpenAutoPositionedScrollWindow('ContattiMod.asp?ID=<%= rs("IDElencoIndirizzi") & queryString %>', 'contatto', 1200, 800, true);">
				<%= ContactFullName(rs) %>
			</a>
			<%
		else 
			%>
			<a href="ContattiMod.asp?ID=<%= rs("IDElencoIndirizzi") & queryString %>" title="apri scheda del contatto n&ordm;<%= rs("IDElencoIndirizzi") %>" <%= ACTIVE_STATUS %>>
				<%= ContactFullName(rs) %>
			</a>
			<%
		end if
		%>
	<% else %>
		<a><%= ContactFullName(rs) %></a>
	<% end if
end function


'.................................................................................................
'..		scrive il nome della pratica linkabile
'..		rs:		recordset aperto su un record della pratica
'.................................................................................................
function PraticaLinkedName(byref rs)
	if not IsNull(rs("pra_id")) then %>
	<a href="javascript:void(0);" title="apri scheda della pratica" <%= ACTIVE_STATUS %>
		onclick="OpenAutoPositionedScrollWindow('PraticaMod.asp?ID=<%= rs("pra_id") %>', 'pratica', 760, 400, true);">
		<%= rs("pra_nome") %>
	</a>
<%	end if
end function


'.................................................................................................
'..		scrive il nome del documento linkabile
'..		rs:		recordset aperto su un record del documento
'.................................................................................................
function DocLinkedName(byref rs)
	if not IsNull(rs("doc_id")) then %>
	<a href="javascript:void(0);" title="apri scheda del documento" <%= ACTIVE_STATUS %>
		onclick="OpenAutoPositionedScrollWindow('DocumentoMod.asp?ID=<%= rs("doc_id") %>', 'documento', 760, 500, true);">
		<%= rs("doc_nome") %>
	</a>
<%	end if
end function


'.................................................................................................
'..		gestione dell'input per la scelta della pratica
'..		tipo:		ATT | DOC
'..		id:			valore dell'id della pratica da impostare solo in modifica
'..		bozza:		true se sto lavorando con attività come bozze
'..		colspan:	# totale di colonne nella tabella madre
'.................................................................................................
Sub SelezionaPratica(conn, rs, tipo, id, modificabile)
dim sql, nome, cnt
	'recupera id della pratica (se gia' impostato)
	id = cInteger(id)
 	tipo = lcase(tipo)
	If modificabile then
		'selezione / modifica della pratica
		if id > 0 then
			'recupera dati della pratica e del contatto da database
			sql = "SELECT pra_id, pra_nome, isSocieta, NomeElencoIndirizzi, CognomeElencoIndirizzi, " & _
			  	  " NomeOrganizzazioneElencoIndirizzi, IDElencoIndirizzi " & _
				  " FROM tb_pratiche INNER JOIN tb_indirizzario ON tb_pratiche.pra_cliente_id = tb_Indirizzario.IDElencoIndirizzi " & _
				  " WHERE pra_id="& cIntero(id)
			rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
			if not rs.eof then
				nome = rs("pra_nome")
				cnt = ContactFullName(rs)
			end if
			rs.close
		end if%>
		<tr>
			<td class="label">contatto:</td>
			<td class="content" name="contatto" colspan="3">
				<input READONLY type="text" name="contatto" value="<%= cnt %>" style="width:70%" 
					   onclick="form1.pratica.onclick();" title="Click per aprire la finestra per la selezione della pratica e del contatto">
			</td>
		</tr>
		<tr>
			<td class="label">pratica:</td>
			<td class="content" colspan="3">
				<table cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td>
						<input type="hidden" name="tfn_<%= tipo %>_pratica_id" value="<%= id %>">
						<input READONLY type="text" name="pratica" style="width:100%" value="<%= nome %>" 
							   onclick="OpenAutoPositionedScrollWindow('PraticheSelezione.asp?field=tfn_<%= tipo %>_pratica_id&selected=' + tfn_<%= tipo %>_pratica_id.value, 'SelezionePratica', 640, 480, true)" title="Click per arpire la filnestra per la selezione della pratica e del contatto">
					</td>
					<td width="30%">
						<a class="button_input" href="javascript:void(0)" onclick="form1.pratica.onclick();" 
							 title="Apre la filnestra per la selezione della pratica e del contatto" <%= ACTIVE_STATUS %>>
							SELEZIONA PRATICA
						</a>
						<a class="button_input" href="javascript:void(0)" 
							 onclick="form1.pratica.value=''; form1.contatto.value=''; form1.tfn_<%= tipo %>_pratica_id.value=''" 
							 title="Cancella la selezione della pratica effettuata" <%= ACTIVE_STATUS %>>
							RESET
						</a>
					</td>
				</tr>
				</table>
			</td>
		</tr>
	<% Elseif id <> 0 then
		'pratica non modificabile
		sql = "SELECT pra_id, pra_nome, isSocieta, NomeElencoIndirizzi, CognomeElencoIndirizzi, " + _
			  " NomeOrganizzazioneElencoIndirizzi, IDElencoIndirizzi " + _
			  " FROM tb_pratiche INNER JOIN tb_indirizzario ON tb_pratiche.pra_cliente_id = tb_Indirizzario.IDElencoIndirizzi " + _
			  " WHERE pra_id=" & cIntero(id)
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		if not rs.eof then%>
			<input type="hidden" name="tfn_<%= tipo %>_pratica_id" value="<%= id %>">
			<tr>
				<td class="label">contatto:</td>
				<td class="content" colspan="3">
					<% ContactLinkedName(rs) %>
				</td>
			</tr>
			<tr>
				<td class="label">pratica:</td>
				<td class="content" colspan="3">
					<% PraticaLinkedName(rs) %>
				</td>
			</tr>
		<% end if
		rs.close
	end if
	
End Sub



'.................................................................................................
'..		controlla se il recapito ha duplicati ed eventualmente mostra il numero di recapiti ed il
'..		link per aprire l'elenco, dal controllo esclude recapiti del contatto corrente
'..		conn:				connessione aperta a database
'..		rs:					recordset aperto su un record del documento
'..		IdIndirizzo:		Id del contatto corrente
'..		Valore:				Valore da controllare
'..		rubriche_visibili	(opzionale) indica il filtro delle rubriche abilitate all'utente
'.................................................................................................
function Check_DuplicatiRecapito(conn, rs, IdIndirizzo, Valore, Rubriche_Visibili)
	dim sql, value
	
	'esegue controllo solo su contatti visibili dall'utente
	if cString(Rubriche_Visibili)="" then
		Rubriche_Visibili = GetList_Rubriche(conn, rs)
		if rubriche_visibili="" then
			rubriche_visibili="0"
		end if
	end if
	
	sql = " SELECT COUNT(*) FROM tb_valorinumeri " + _
		  " WHERE id_Indirizzario<>" & cIntero(IdIndirizzo) & " AND " + _
		  		" LTRIM(RTRIM(ValoreNumero)) LIKE '" & ParseSql(Trim(valore), adChar) & "' " + _
				" AND id_Indirizzario IN (SELECT id_indirizzo FROM rel_rub_ind WHERE id_rubrica IN (" & rubriche_visibili & ") ) "
	value = cInteger(GetValueList(conn, rs, sql))
	if value > 0 then
		'valore duplicato %>
		<table align="right" cellpadding="0" cellspacing="0" style="margin-top:1px;">
			<tr>
				<td class="content warning">
					n&ordm; <%= value %> duplicati
					<a class="button_L2" title="Apri l'elenco dei contatti nei quali &egrave; presente il recapito." <%= ACTIVE_STATUS %>
					   href="javascript:void(0);" onclick="OpenAutoPositionedScrollWindow('ContattiRecapitiDuplicati.asp?CONTATTO=<%= IdIndirizzo %>&recapito=<%= Server.URLEncode(valore) %>', 'CntDuplicato', 500, 405, true)">
						elenco
					</a>
				</td>
			</tr>
		</table>
	<% end if
end function

%>