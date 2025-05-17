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
	Classe.Requested_Fields_List	= "tft_car_fornitore_cod;tfn_car_magazzino_id;tfd_car_data"
	Classe.Checkbox_Fields_List 	= "chk_car_movimentato"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_carichi"
	Classe.id_Field					= "car_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
dim sql, cond,rslocal
dim articolo_variante_id,qta,qta_old
	' TO DO: Gestione del flag movimentato.
	' Se il flag viene settato occorre eseguire su tutte le righe di dettaglio l'aggiornamento 
	' articolo_variante_id: codice dell'articolo in corso di aggiornamento

	If request("chk_car_movimentato")<>"" then
		sql = " SELECT gia_magazzino_id, gia_art_var_id, gia_qta, gia_impegnato, gia_ordinato, rel_id, rel_art_id, rel_qta_min_ord, rcv_car_id, rcv_qta " & _
		      " FROM grel_giacenze INNER JOIN grel_art_valori ON grel_giacenze.gia_art_var_id = grel_art_valori.rel_id " + _
			  "		 INNER JOIN grel_carichi_var ON grel_art_valori.rel_id = grel_carichi_var.rcv_art_var_id " + _
			  " WHERE rcv_car_id = " & ID
		set rslocal = Server.CreateObject("ADODB.Recordset")
		rslocal.open sql,conn,adOpenStatic, adLockOptimistic, adCmdText
		do while not rslocal.eof 
			articolo_variante_id = rslocal("rel_id")
			qta = rslocal("rcv_qta")
			qta_old = 0
			CALL SetGiacenza(conn, articolo_variante_id, "+", "G", rslocal("gia_magazzino_id"), qta-qta_old)
			CALL SetGiacenza(conn, articolo_variante_id, "-", "O", rslocal("gia_magazzino_id"), qta-qta_old)
			rslocal.movenext
		loop
	end if
	'imposta parametri per passare alla pagina successiva
		Classe.isReport = FALSE
		if request.form("salva") <> "" then
			Classe.Next_Page = "MagazziniCarichiMod.asp?ID="& ID
		else
			Classe.Next_Page = "MagazziniCarichi.asp?IDMAG=" & request("IDMAG")
		end if

end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>