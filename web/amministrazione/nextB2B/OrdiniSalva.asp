<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Class_Mailer.asp" -->
<!--#INCLUDE FILE="../library/ClassConfiguration.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tfn_ord_riv_id"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_ordini"
	Classe.id_Field					= "ord_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, stato_ordine
	
	if request("ID") = "" then
		'......................................................................................................
		'EVENTO DICHIARATO NEL FILE DI CONFIGURAZIONE EVENTI DEL CLIENTE
		CALL On_Order_NEW(conn, ID)
		'......................................................................................................
	end if
	
	'gestione tipo e stato dell'ordine
	sql = "SELECT * FROM gtb_ordini WHERE ord_id="& ID
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	
	stato_ordine = request("stato_ordine")
	if cInteger(request("old_stato_ordine")) <> cInt(stato_ordine) then
		SELECT CASE cInt(stato_ordine)
		CASE ORDINE_NON_CONFERMATO
			if rs("ord_movimenta") then
				Session("ERRORE") = "Errore nello stato dell'ordine: non &egrave; possibile passare allo stato &quot;non confermato&quot;"
				Exit Sub
			end if
			'immissione ordine non confermato
			if rs("ord_impegna") then
				'se lo stato dell'ordine e' impegnato: toglie quantit&agrave; da impegnato
				CALL SetGiacenza_ord(conn, ID, "O", "-", "I", rs("ord_magazzino_id"))
			end if
			
			rs("ord_impegna") = false
			rs("ord_movimenta") = false
			rs("ord_archiviato") = false
			rs.update
			
		CASE ORDINE_CONFERMATO
			if rs("ord_movimenta") OR rs("ord_archiviato") then
				Session("ERRORE") = "Errore nello stato dell'ordine: non &egrave; possibile passare allo stato &quot;confermato&quot;"
				Exit Sub
			end if
			'conferma dell'ordine
			if not rs("ord_impegna") then
				'impegna la merce dell'ordine
				CALL SetGiacenza_ord(conn, ID, "O", "+", "I", rs("ord_magazzino_id"))
				
				rs("ord_impegna") = true
				rs("ord_movimenta") = false
				rs.update
			end if
			
		CASE ORDINE_EVASO
			'evasione dell'ordine
			if rs("ord_impegna") then
				'se la merce &egrave; gi&agrave; stata impegnata la toglie dall'impegnato
				CALL SetGiacenza_ord(conn, ID, "O", "-", "I", rs("ord_magazzino_id"))
			end if
			if not rs("ord_movimenta") then
				'movimenta la merce per evasione
				CALL SetGiacenza_ord(conn, ID, "O", "-", "M", rs("ord_magazzino_id"))
			end if
			
			rs("ord_impegna") = false
			rs("ord_movimenta") = true
			rs("ord_archiviato") = false
			rs.update
		
		CASE ORDINE_ARCHIVIATO
			rs("ord_archiviato") = true
			rs.update
		
		END SELECT
		
		if cInteger(request("old_stato_ordine")) = ORDINE_NON_CONFERMATO AND _
		   (cInt(stato_ordine) = ORDINE_CONFERMATO OR cInt(stato_ordine) = ORDINE_EVASO) then
			'conferma dell'ordine
			'agggiorna rank
			RankingOrdine conn, rs("ord_id")
			
			'......................................................................................................
			'EVENTO DICHIARATO NEL FILE DI CONFIGURAZIONE EVENTI DEL CLIENTE
			'ATTENZIONE NON METTERE ALCUNA OPERAZIONE IN QUESTO PUNTO: All'INTERNO DELLA FUNZIONE 
			'VIENE GIA' CHIUSA LA TRANSAZIONE!!
			if Session("ERRORE") = "" then
				CALL On_Order_CONFIRM(conn, ID)
			end if
			'......................................................................................................
		else
		end if
	end if
	rs.close
	
	
	'......................................................................................................
	'FUNZIONE DICHIARATA NEL FILE DI CONFIGURAZIONE DEL CLIENTE
	CALL ADDON__ORDINI__form_salva(conn, rs)
	'......................................................................................................
	
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	if request.form("salva") <> "" then
		Classe.Next_Page = "OrdiniMod.asp?ID="& ID
	else
		Classe.Next_Page = "Ordini.asp"
	end if
	
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
%>