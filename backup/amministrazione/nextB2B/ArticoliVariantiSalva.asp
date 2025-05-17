<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<!--#INCLUDE FILE="Tools4save_B2B.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	if request("ID")<>"" then
		if request("codice_modificabile")<>"" then
			Classe.Requested_Fields_List	= "tft_rel_cod_int"
		else
			Classe.Requested_Fields_List	= ""
		end if
	else
		Classe.Requested_Fields_List	= "tft_rel_cod_int;tfn_rel_prezzo"
	end if
	Classe.Checkbox_Fields_List 	= "chk_rel_disabilitato"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "grel_art_valori"
	Classe.id_Field					= "rel_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
    Classe.SetUpdateParams("rel_")
	
'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, field, ValoriList, ValoriCount, rsv
    dim objVariante 
	set rsv = Server.CreateObject("ADODB.RecordSet")
    
    set objVariante = new GestioneVariante
    set objVariante.conn = conn
	
	'verifica se il codice &egrave; univoco
	if request("ID")="" OR request("codice_modificabile")<>"" then
		if not objVariante.CodeIsUnique(ID, request("tft_rel_cod_int")) then	
			Session("ERRORE") = "Codice dell'articolo non univoco: Esiste gi&agrave; un articolo con codice &quot;" & request("tft_rel_cod_int") & "&quot;."
			exit sub
		end if
	end if
	
	'inserimento dell'esploso delle varianti nei listini, nei codici
	if request("ID")="" then
		sql = "SELECT * FROM grel_art_valori WHERE rel_id=" & ID
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		
		'inserisce i valori della variante nella relazione
		ValoriCount = 0
		for each field in request.form
			if instr(1, field, "valore_", vbTextCompare)>0 AND request(field)<>"" then
				ValoriList = ValoriList & request(field) & ", "
				ValoriCount = ValoriCount + 1
				sql = "INSERT INTO grel_art_vv (rvv_art_var_id, rvv_val_id) VALUES(" & ID & ", " & ParseSQL(request(field), adChar) & ")"
				CALL conn.execute(sql, , adExecuteNoRecords)
			end if
		next
		if ValoriCount=0 then
			'nessun valore variante selezionato
			Session("ERRORE") = "Selezionare almeno un valore di variate."
			exit sub
		else
			ValoriList = left(ValoriList, len(ValoriList)-2)
		end if
				
		'verifica che la variante non sia gi&agrave; stata immessa
		sql = " SELECT rel_id FROM grel_art_valori INNER JOIN grel_art_vv ON grel_art_valori.rel_id=grel_art_vv.rvv_art_var_id " + _
			  " WHERE rel_art_id = " & rs("rel_art_id") & " AND rvv_val_id IN (" & ValoriList & ") " & _
			  " GROUP BY rel_id HAVING COUNT(rel_id)=" & ValoriCount
		if instr(1, GetValueList(conn, NULL, sql), ",", vbTextCompare) > 0 then			'se sono presenti pi&ugrave; record nell'elenco risultante e' presente almeno una virgola
			Session("ERRORE") = "La variante dell'articolo inserita &egrave; gi&agrave; presente. Verificare l'elenco delle varianti."
			exit sub
		end if
		
        'iniserisce righe di default nei magazzini e nei listini
        CALL objVariante.InsertDefaultRows(rs("rel_id"))
		
		'calcola l'ordiamento della variante
		rs("rel_ordine") = objVariante.GetOrdineVariante(rsv, rs("rel_id"))
		rs.update
		rs.close
	end if
	
    'imposta date aggiornamento articolo "padre"
    CALL objVariante.UpdateParamsArticolo(ID)
	
	'imposta parametri per salvare e chiudere la finestra corrente
	Classe.isReport = TRUE
    
	if request("collegamento_articoli")="" AND request("collegamento_variante")="" then 
		'non esegue il reload della pagina principale se sono in inserimento nuova variante da gestione articoli importati (ArticoliCollegamento_WZD_NewVar.asp)%>
		<script language="JavaScript" type="text/javascript">
			opener.location.reload(true);
			window.close();
		</script>
	<% else
		Session("VARIANTE_INSERITA") = ID
	end if
	
	set rsv = nothing
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>