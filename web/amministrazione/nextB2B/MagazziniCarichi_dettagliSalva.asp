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
	Classe.Requested_Fields_List	= "tfn_rcv_car_id;tfn_rcv_art_var_id;tfn_rcv_qta"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "grel_carichi_var"
	Classe.id_Field					= "rcv_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
dim sql,qta,qta_old,magazzinoID,caricoID,articolo_variante_id

	'controllo che la quantità da caricare sia positiva
	if CInteger(request("tfn_rcv_qta")) <= 0 then
		session("ERRORE") = "Quantit&agrave; errata"
		conn.rollbacktrans
		Exit Sub
	else
		qta = request("tfn_rcv_qta")
		qta_old = request("old_qta")
	end if
	
	'se ho gia inserito questo dettaglio restituisco errore
	sql = "SELECT COUNT(*) FROM grel_carichi_var "& _
		  "WHERE rcv_art_var_id="& cIntero(request("tfn_rcv_art_var_id")) & _
		  " AND rcv_car_id="& cIntero(request("tfn_rcv_car_id"))
	if CInt(GetValueList(conn, rs, sql)) > 1 then
		session("ERRORE") = "Variante gi&agrave; caricata!"
		conn.rollbacktrans
	end if
	
	if session("ERRORE") = "" then
		'Recupero i dati del Magazzino e del Carico
		caricoID = request("tfn_rcv_car_id")
		magazzinoID = request("MAG_ID")
		articolo_variante_id = request("tfn_rcv_art_var_id")
		'imposto campi
		sql = "SELECT * FROM grel_carichi_var WHERE rcv_id="& ID
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		rs("rcv_qta") = qta
		rs.update
		rs.close
	
		' Esegui l'aggiornamento delle giacenze (quantità e ordinato) 	
		'CALL SetGiacenza(conn, articolo_variante_id, "+", "G", magazzinoID, qta-qta_old)
		'CALL SetGiacenza(conn, articolo_variante_id, "-", "O", magazzinoID, qta-qta_old)
		
		'imposta parametri per salvare e chiudere la finestra corrente
		Classe.isReport = TRUE
%>
		<script language="JavaScript" type="text/javascript">
			opener.location.reload(true);
			window.location = "ArticoliSelPrz.asp?CAR_ID=<%=caricoID  %>"
			// window.close();
		</script>
	
<%	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>
