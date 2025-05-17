<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tfn_det_ord_id;tfn_det_ind_id;tfn_det_art_var_id;tfn_det_qta"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_dettagli_ord"
	Classe.id_Field					= "det_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
dim sql, prezzo, listino_id, impegna, movimenta, magazzino_id
	
	'controllo che qta > 0
	if CInteger(request("tfn_det_qta")) = 0 then
		session("ERRORE") = "Quantit&agrave; non valida"
		Exit Sub
	end if
	
	'verifica esistenza stesso articolo ordinato nella stessa destinazione.
	sql = "SELECT COUNT(*) FROM gtb_dettagli_ord "& _
		  "WHERE det_art_var_id="& cIntero(request("tfn_det_art_var_id")) & _
		  " AND det_ord_id="& cIntero(request("tfn_det_ord_id")) &" AND det_ind_id="& cIntero(request("tfn_det_ind_id")) & _
		  " AND det_cod_promozione='" & cIntero(request("tft_det_cod_promozione")) & "'"
	if CInt(GetValueList(conn, rs, sql)) > 1 then
		session("ERRORE") = "Dettaglio gi&agrave; immesso! Cambiare indirizzo o articolo."
		exit sub
	end if
	
	'controllo vendibile singolarmente e quantita articolo collegato con vincolo in vendita 
	'non maggiore della somma degli articoli contenenti ordinati
	sql = "SELECT art_noVenSingola, "& _
		  "		(SELECT SUM(det_qta) FROM (gtb_dettagli_ord d "& _
		  "		 INNER JOIN gv_articoli a ON d.det_art_var_id = a.rel_id) "& _
		  "		 INNER JOIN grel_art_acc acc ON a.art_id = acc.aa_art_id " & _
		  "		 INNER JOIN gtb_accessori_tipo ON acc.aa_tipo_id=gtb_accessori_tipo.at_id " & _
		  "		 WHERE aa_acc_id = art.art_id AND gtb_accessori_tipo.at_vincolo_vendita=1) AS num "& _
		  "FROM gv_articoli art "& _
		  "WHERE rel_id = "& cIntero(request("tfn_det_art_var_id"))
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	if rs("art_noVenSingola") AND CInteger(rs("num")) < CInteger(request("tfn_det_qta")) then
		session("ERRORE") = "Articolo non vendibile singolarmente o quantit&agrave; superiore agli articoli collegati ordinati."
		Exit Sub
	end if
	rs.close
	
	if session("ERRORE") = "" then
		sql = "SELECT * FROM gtb_ordini o "& _
			  "INNER JOIN gtb_rivenditori r ON o.ord_riv_id=r.riv_id "& _
			  "WHERE ord_id="& cIntero(request.form("tfn_det_ord_id"))
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			listino_id = rs("riv_listino_id")
			impegna = rs("ord_impegna")
			movimenta = rs("ord_movimenta")
			magazzino_id = rs("ord_magazzino_id")
		rs.close
			
		'imposto giacenza
		if cInteger(request("qta_old")) <> cInteger(request("tfn_det_qta")) then
			if impegna then
				CALL SetGiacenza_ord(conn, ID, "D", "+", "I", magazzino_id)
			elseif movimenta then
				CALL SetGiacenza_ord(conn, ID, "D", "-", "M", magazzino_id)
			end if
		end if
		
		'imposta parametri per salvare e chiudere la finestra corrente
		Classe.isReport = TRUE%>
		<script language="JavaScript" type="text/javascript">
			opener.location.reload(true);
			window.close();
		</script>
	
<%	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>