<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Class_Mailer.asp" -->
<!--#INCLUDE FILE="../library/ClassConfiguration.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede.asp" -->

<%
dim conn, sql
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tfn_ddt_cliente_id;tfn_ddt_destinazione_id;"
	if request("tfn_ddt_trasportatore_id") <> "" then
		Classe.Requested_Fields_List = Classe.Requested_Fields_List & "tft_ddt_peso;tft_ddt_volume;tft_ddt_numero_colli;"
	end if
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "sgtb_ddt"
	Classe.id_Field					= "ddt_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, listaId, ids
	
	'inserissco il numero del ddt, se è un nuovo inserimento
	if request("nuovo_inserimento") = "true" then
		sql = " UPDATE sgtb_ddt SET " & _
			  "		ddt_numero = (SELECT ISNULL(MAX(ddt_numero), 0) + 1 FROM sgtb_ddt WHERE ddt_categoria_id = " & cIntero(request("tfn_ddt_categoria_id")) & _
			  " 				  AND " & SQL_BetweenDate(conn, "ddt_data", "01/01/"&Year(Now()), "31/12/"&Year(Now())) & ") WHERE ddt_id = " & ID
		conn.execute(sql)
	end if

	sql = "UPDATE sgtb_schede SET sc_rif_DDT_di_resa_id = 0 WHERE sc_rif_DDT_di_resa_id = " & ID
	CALL conn.execute(sql, , adExecuteNoRecords)
	
	sql = "UPDATE sgtb_schede SET sc_costo_riconsegna = 0 WHERE ISNULL(sc_rif_DDT_di_resa_id, 0)=0 AND sc_cliente_id = " & cIntero(request("tfn_ddt_cliente_id"))
	CALL conn.execute(sql, , adExecuteNoRecords)
	
	
	'inserimento id della spedizione nelle schede scelte
	if Trim(Replace(request("id_schede"),",",""))<>"" then
		listaId = cString(request("id_schede"))
		listaId = Split(listaId, ",")
		for each ids in listaId
			ids = cIntero(Trim(ids))
			sql = "UPDATE sgtb_schede SET sc_rif_DDT_di_resa_id = " & ID & " WHERE sc_id = " & ids
			conn.Execute(sql)
			if cString(request("costo_scheda_"&ids))<>"" then
				sql = "UPDATE sgtb_schede SET sc_costo_riconsegna = " & ParseSQL(request("costo_scheda_"&ids), adNumeric) & " WHERE sc_id = " & ids
				conn.Execute(sql)
			end if
		next
	'else
	'	Session("ERRORE") = "ATTENZIONE! Associare almeno una scheda."
	end if
	

	'cambiare lo stato della scheda?
	
	
	'imposta parametri per passare alla pagina successiva
	Classe.Next_Page = "Spedizioni.asp"
	if request("reload")="true" then
		%>
		<script language="JavaScript" type="text/javascript">
			opener.location.reload(true);
			window.close();
		</script>
		<%
	else
		Classe.isReport = FALSE
		if request.form("salva") <> "" then
			Classe.Next_Page = "SpedizioniMod.asp?ID="& ID
		end if
	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
%>