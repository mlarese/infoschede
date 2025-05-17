<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE file="Tools_b2b.asp" -->
<%
dim Classe
Set Classe = New OBJ_Salva

'Impostazione parametri
Classe.ConnectionString 		= Application("DATA_ConnectionString")
Classe.Requested_Fields_List	= "tft_listino_codice"
Classe.Checkbox_Fields_List 	= "chk_listino_B2C;chk_listino_offerte;chk_listino_base"
Classe.Page_Ins_Form			= ""
Classe.Page_Mod_Form			= ""
Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
Classe.Next_Page_ID				= FALSE
Classe.Table_Name				= "gtb_listini"
Classe.id_Field					= "listino_id"
Classe.Read_New_ID				= TRUE
Classe.isReport 				= TRUE
Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, dcrea, dscad
	
	if request("ID")="" then
		'inserimento nuovo listino
		sql = ""
		'copia dei prezzi
		if request("chk_listino_base_attuale")<>"" OR _
		   (request("chk_listino_base")<>"" AND cInteger(request("copia_da_listino"))=0) then
			'copia i prezzi dal prezzo base dell'articolo nella scheda.
			sql = " SELECT art_iva_id, rel_prezzo, 0, 0, 1, 0, rel_scontoQ_id, " & ID & ", rel_id " + _
				  " FROM gv_articoli "
			
		elseif cInteger(request("copia_da_listino"))>0 or cInteger(request("tfn_listino_ancestor_id"))>0 then
			'copia i prezzi da altro listino (listino semplice e non in offerta speciale)
			sql = " SELECT prz_iva_id, prz_prezzo, prz_var_sconto, prz_var_euro, prz_visibile, " + _
				  " prz_promozione, prz_scontoQ_id, " & ID & ", prz_variante_id " + _
				  " FROM gv_listini WHERE prz_listino_id=" & _
				  IIF(cInteger(request("copia_da_listino"))>0, cInteger(request("copia_da_listino")), cInteger(request("tfn_listino_ancestor_id")))
		elseif cReal(request("tfn_listino_default_var_sconto"))<>0 OR cReal(request("tfn_listino_default_var_euro"))<>0 then
			'applica variazione di default dal listino base attualmente in vigore
			sql = " SELECT prz_iva_id, "
			
			'applica eventuali variazioni di defaut
			dim default_var_sconto, default_var_euro
			default_var_sconto = ConvertForSave_Number(request("tfn_listino_default_var_sconto"), 0)
			default_var_euro = ConvertForSave_Number(request("tfn_listino_default_var_euro"), 0)
			if cReal(request("tfn_listino_default_var_sconto"))<>0 then
				'sconto di default in percentuale
				sql = sql + " prz_prezzo + (" & ParseSQL(default_var_sconto / 100, adNumeric) & " * prz_prezzo), " & ParseSQL(default_var_sconto, adNumeric) & ", 0,"
			elseif cReal(request("tfn_listino_default_var_euro"))<>0 then
				'sconto di default in euro
				sql = sql + " prz_prezzo + " & ParseSQL(default_var_euro, adNumeric) & ", 0, " & ParseSQL(default_var_euro, adNumeric) & ", "
			else
				'nessuno sconto di default
				sql = sql + " prz_prezzo, 0, 0, "
			end if
			sql = sql + " 1, 0, prz_scontoQ_id, " & ID & ", prz_variante_id " + _
						" FROM gv_listini WHERE prz_listino_id IN (SELECT listino_id FROM gtb_listini WHERE ISNULL(listino_base_attuale,0)=1) "
		end if
		if sql <> "" then
			sql = " INSERT INTO gtb_prezzi (prz_iva_id, prz_prezzo, prz_var_sconto, prz_var_euro, prz_visibile, " + _
			  	  " prz_promozione, prz_scontoQ_id, prz_listino_id, prz_variante_id ) " + sql			  
			CALL conn.execute(sql, , adExecuteNoRecords)
		end if
	else
		'gestione del listino derivato
		if cInteger(request("listino_ancestor_id"))>0 then
			if cInteger(request("listino_ancestor_id"))<>cInteger(request("tfn_listino_ancestor_id")) then
				'cambiato il listino da cui deriva: deve sosituire tutti i prezzi:
				'eliminando quelli uguali al vecchi listino ANCESTOR
				'importando quelli del nuovo listino ancestor
				'sempre escludendo quelli variati.
				CALL Listino_ChangeAncestor(conn, ID, cInteger(request("tfn_listino_ancestor_id")), cInteger(request("listino_ancestor_id")))
			end if
		end if
	end if
	
	if cInteger(request("listino_ancestor_id")) <> cInteger(request("tfn_listino_ancestor_id")) then
		'imposta flag del listino principale
		sql = "UPDATE gtb_listini SET listino_with_child=1 WHERE listino_id=" & cIntero(request("tfn_listino_ancestor_id"))
		CALL conn.execute(sql, , adExecuteNoRecords)
		
		if cInteger(request("listino_ancestor_id"))>0 then
			sql = " UPDATE gtb_listini SET listino_with_child=CASE WHEN (SELECT COUNT(*) FROM gtb_listini L_child " + _
				  " WHERE L_child.listino_ancestor_id = gtb_listini.listino_id)>1 THEN 1 ELSE 0 END " + _
				  " WHERE gtb_listini.listino_id=" & cIntero(request("listino_ancestor_id"))
			CALL conn.execute(sql, , adCmdText)
		end if
	end if
	
	
	'controllo per listino "offerta speciale"
	if request.form("chk_listino_offerte") <> "" OR _
	   request.form("chk_listino_b2c")<>"" then
		
		dcrea = request.form("tfd_listino_dataCreazione")
		dscad = request.form("tfd_listino_dataScadenza")
		'controllo le scadenze
		if dcrea<>"" AND dscad<>"" AND isDate(dcrea) AND isDate(dscad) then
			if CDate(dcrea) > CDate(dscad) then
				session("ERRORE") = "Data creazione maggiore della data scadenza"
				Exit Sub
			end if
		end if
		
		sql = " SELECT COUNT(*) FROM gtb_listini "& _
			  " WHERE listino_id <> "& ID 
		if dcrea<>"" AND dscad<>"" AND isDate(dcrea) AND isDate(dscad) then
			sql = sql + " AND (" + _
			  			" (NOT (" + SQL_CompareDateTime(conn, "listino_dataScadenza", adCompareLessThan, dcrea) + " OR " + SQL_CompareDateTime(conn, "listino_dataCreazione", adCompareGreaterThan, dateAdd("d", 1, dscad)) + ")) " + _
			  			" OR " +_
			  			" (" + SQL_CompareDateTime(conn, "listino_dataCreazione", adCompareLessThan, dscad) + " AND " + SQL_IsNull(conn, "listino_dataScadenza") + "))"
		else
			sql = sql + " AND (" + SQL_CompareDateTime(conn, "listino_dataScadenza", adCompareGreaterThan, dcrea) + " OR " + SQL_IsNull(conn, "listino_dataScadenza") + ") "
		end if
		if request.form("chk_listino_offerte")<>"" then
			sql = sql + " AND listino_offerte=1"
		else
			sql = sql + " AND listino_b2c=1"
		end if

		if cInteger(GetValueList(conn, rs, sql)) > 0 then
			if request.form("chk_listino_offerte")<>"" then
				session("ERRORE") = "Esiste gi&agrave; un'offerta speciale che interseca il periodo scelto"
			else
				session("ERRORE") = "Esiste già un listino per il pubblico per il periodo scelto."
			end if
			Exit Sub
		end if
		
	end if

	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	if request("salva")<>"" then
		Classe.Next_Page = "ListiniMod.asp?ID=" & ID
	else
		Classe.Next_Page = "Listini.asp"
	end if
end Sub


'salvataggio/modifica dati
Classe.Salva()

%>