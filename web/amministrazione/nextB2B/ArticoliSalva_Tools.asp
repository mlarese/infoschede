
<%
dim redirect
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	if cInteger(request("tfn_art_varianti"))=0 OR request("valori")<>"" then
		Classe.Requested_Fields_List	= "tft_art_cod_int;"
		if cIntero(request("ID")) = 0 then
			Classe.Requested_Fields_List	= Classe.Requested_Fields_List & "tfn_art_prezzo_base;"
		end if
	end if
	Classe.Requested_Fields_List	= Classe.Requested_Fields_List + "tft_art_nome_it"
	Classe.Checkbox_Fields_List 	= "chk_art_disabilitato;chk_art_NoVenSingola;chk_art_unico"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_articoli"
	Classe.id_Field					= "art_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
    Classe.SetUpdateParams("art_")
	
	
'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, objVariante
	dim rel_id, var, listaid
	
    set objVariante = new GestioneVariante
    set objVariante.conn = conn
    
	if cBoolean(cString(Session("INIBISCI_PREZZO_A_ZERO")), false) AND cReal(Request("tfn_art_prezzo_base")) = 0 then
		Session("ERRORE") = "Impossibile inserire un articolo con il prezzo uguale a zero"
	end if
	
	if cInteger(request("tfn_art_varianti"))=0 then
		
        'inserimento variante unica
        sql = "SELECT rel_id FROM grel_art_valori WHERE rel_art_id=" & ID
        CALL objVariante.InsertUpdateComplete(ID, GetValueList(conn, rs, sql), "", _
                                      request("tft_art_cod_int"), request("tft_art_cod_pro"), request("tft_art_cod_alt"), _
                                      cReal(request("tfn_art_prezzo_base")), NULL, NULL, NULL, request("nfn_art_scontoq_id"), _
                                      (request("chk_art_disabilitato") <> ""), request("tfn_art_giacenza_min"), request("tfn_art_lotto_riordino"), request("tfn_art_qta_min_ord"), _
									  request("extN_rel_peso_netto"), request("extN_rel_peso_lordo"), request("extN_rel_colli_num"), request("extN_rel_collo_pezzi_per"), _
									  request("extN_rel_collo_width"), request("extN_rel_collo_height"), request("extN_rel_collo_lenght"), request("extN_rel_collo_volume") _
									  )
        rel_id = cIntero(GetValueList(conn, rs, sql))
		
		if rel_id > 0 then
			'salva codici alternativi
			for each var in request.form
				if instr(1, var, "codice_articolo_", vbTextCompare)>0 AND _
				  request.form(var)<>"" then
					'campo con codice alternativo
					ListaId = cIntero(RemoveInvalidChar(var, NUMERIC_CHARSET))
					if listaId > 0 then
						sql = "SELECT * FROM gtb_codici WHERE cod_variante_id = " & rel_id & " AND cod_lista_id = " & ListaId
						rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
						if rs.eof then
							rs.addnew
							rs("cod_variante_id") = rel_id
							rs("cod_lista_id") = ListaId
						end if
						rs("cod_codice") = request.form(var)
						rs.update
						rs.close
					end if
				end if
			next
		end if
	else
		'articolo con varianti
		if request("ID")="" then
			'inserimento nuovo articolo con varianti
			if request("valori")<>"" then
				'inserimento varianti scelte
				dim rsval, rsvar, Fsql, Wsql, Osql, ListaIdValori
				set rsval = Server.CreateObject("ADODB.Recordset")
				set rsvar = Server.CreateObject("ADODB.Recordset")
				
				'recupera elenco varianti
				sql = " SELECT * FROM gtb_varianti WHERE var_id IN (SELECT val_var_id " + _
					  " FROM gtb_valori WHERE val_id IN (" & ParseSQL(request("valori"), adChar) & ")) ORDER BY var_ordine"
				rsvar.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				
				'costruzione della query di prodotto cartesiano per "esploso valori"
				sql = ""		'elenco dei campi
				Fsql = ""		'clausola from
				Wsql = ""		'clausola where
				Osql = ""		'clausola order by
				while not rsvar.eof
					sql = sql + IIF(rsvar.absoluteposition = 1, " SELECT ", ", ") + _
								"(v" & rsvar("var_id") & ".val_id) AS val_id_" & rsvar("var_id") & _
								", (v" & rsvar("var_id") & ".val_nome_it) AS val_nome_it_" & rsvar("var_id") & _
								", (v" & rsvar("var_id") & ".val_cod_int) AS val_cod_int_" & rsvar("var_id") & _
								", (v" & rsvar("var_id") & ".val_cod_pro) AS val_cod_pro_" & rsvar("var_id") & _
								", (v" & rsvar("var_id") & ".val_cod_alt) AS val_cod_alt_" & rsvar("var_id") & _
								", (v" & rsvar("var_id") & ".val_ordine) AS val_ordine_" & rsvar("var_id")
					Fsql = Fsql + IIF(rsvar.absoluteposition = 1, " FROM ", ", ") + "gtb_valori v" & rsvar("var_id")
					Wsql = Wsql + IIF(rsvar.absoluteposition = 1, " WHERE ", " AND ") + "(v" & rsvar("var_id") & ".val_var_id=" & rsvar("var_id") & " AND v" & rsvar("var_id") & ".val_id IN (" & ParseSQL(request("valori"), adChar) & "))"
					Osql = Osql + IIF(rsvar.absoluteposition = 1, " ORDER BY ", ", ") + "v" & rsvar("var_id") & ".val_ordine "
					rsvar.movenext
				wend
				
				'elenco valori risultanti dal prodotto cartesiano
				sql = sql + Fsql + Wsql + Osql
				rsval.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				
                while not rsval.eof
                    ListaIdValori = ""
                    rsvar.movefirst
					while not rsvar.eof
                        ListaIdValori = ListaIdValori & rsval("val_id_" & rsvar("var_id"))
						rsvar.movenext
                        if not rsvar.eof then
                            ListaIdValori = ListaIdValori & ", "
                        end if
					wend
                    
                    CALL objVariante.InsertUpdateComplete(ID, NULL, ListaIdValori, _
                                                  GetCode(ID, rsvar, rsval, "int", true), GetCode(ID, rsvar, rsval, "pro", false), GetCode(ID, rsvar, rsval, "alt", false), _
                                                  cReal(request("tfn_art_prezzo_base")), false, 0, 0, request("nfn_art_scontoq_id"), _
                                                  (request("chk_art_disabilitato") <> ""), request("tfn_art_giacenza_min"), request("tfn_art_lotto_riordino"), request("tfn_art_qta_min_ord"), _
												request("extN_rel_peso_netto"), request("extN_rel_peso_lordo"), request("extN_rel_colli_num"), request("extN_rel_collo_pezzi_per"), _
												request("extN_rel_collo_width"), request("extN_rel_collo_height"), request("extN_rel_collo_lenght"), request("extN_rel_collo_volume") _
												)
                    
                    rsval.movenext
                wend

				rsval.close
				rsvar.close
				set rsval = nothing
				set rsvar = nothing
			end if
		end if
	end if
	
	'gestione caratteristiche tecniche
	CALL DesSalva(conn, ID, "grel_art_ctech", "rel_ctech_", "rel_art_id", "rel_ctech_id")
	
	
	'gestione pagina successiva
	if Session("ERRORE") = "" then
	
		'..............................................................................
		'sincronizzazione con i contenuti e l'indice
		CALL Index_UpdateItem(conn, Classe.Table_Name, ID, false)
		'..............................................................................
	
		if request("SALVA")<>"" then
			if cString(redirect)<>"" then
				'per Infoschede (nell'inserimento di un ricambio)
				Classe.Next_Page = redirect & "&RELID=" & rel_id
			else
				Classe.Next_Page = "ArticoliMod.asp?ID=" & ID
			end if
		else
			Classe.Next_Page = "Articoli.asp"
		end if
		
		if request("ID")="" then
			Classe.isReport = false
			
			if request("external")<>"" then
				sql = "SELECT art_nome_it, art_cod_int FROM gtb_articoli WHERE art_id=" & ID
				rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText%>
				<script type="text/javascript" language="JavaScript">
					//parte che imposta i dati dell'articoli immesso nel form di selezione (ArticoliCollegamento_WZD_2.asp in next-B2B import
					if (opener){
						if (opener.document.form1.articolo_id){
							opener.document.form1.articolo_id.value='<%= ID %>';
						}
						if (opener.document.form1.articolo_codice){
							opener.document.form1.articolo_codice.value='<%= JSEncode(rs("art_cod_int"), "'") %>';
						}
						if (opener.document.form1.articolo_name){
							opener.document.form1.articolo_name.value='<%= JSEncode(rs("art_nome_it"), "'") %>';
						}
					}
					
					//esegue il redirect alla pagina desiderata
					document.location = "<%= Classe.Next_Page %>";
				</script>
				<%rs.close
				response.flush
			end if
		else
			Classe.isReport = false
		end if
	end if
end Sub

'salvataggio/modifica dati
Classe.Salva()


function GetCode(art_ID, rsvar, rsval, Radix, forced)
	if request("tft_art_cod_" + radix)="" AND not(forced) then
		GetCode = ""
	else
		dim cod_part
		cod_part = GetCodePart(request("tft_art_cod_" + radix), art_ID)
		rsvar.moveFirst
		while not rsvar.eof
			cod_part = cod_part & CODE_SEPARATOR & GetCodePart(rsval("val_cod_" & radix & "_" & rsvar("var_id")), rsval("val_id_" & rsvar("var_id")))
			rsvar.movenext
		wend
		GetCode = cod_part
	end if
end function

function GetCodePart(Code, ID)
	Code = cString(Code)
	if Code<>"" then
		GetCodePart = Code
	else
		GetCodePart = cString(ID)
	end if
end function
%> 