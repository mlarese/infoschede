<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->

<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_dot_nome_it"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_dettagli_ord_tipo"
	Classe.id_Field					= "dot_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
    
    'GESTIONE INFORMAZIONI PER RIGA
    if request("ID")<>"" then
        dim CaratList, Carat, sql
    
        'cancella relazioni precedenti
        sql = "DELETE FROM grel_dettagli_ord_tipo_Des WHERE rtd_tipo_id=" & ID
        CALL conn.execute(sql, 0, adExecuteNoRecords)
    
        'gestione categorie associate
        CaratList = split(replace(request("caratteristiche_associate"), " ", ""), ",")
    
        for each Carat in CaratList
            'recupera dati ordine
            sql = "INSERT INTO grel_dettagli_ord_tipo_des (rtd_descrittore_id, rtd_tipo_id, rtd_ordine) " + _
                  " VALUES (" & Carat & ", " & ID & ", " & cIntero(request("rel_ordine_" & Carat)) & ")"
            CALL conn.execute(sql, , adExecuteNoRecords)
        next
    end if

	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	if request.form("salva") <> "" then
		Classe.Next_Page = "OrdiniRigheTipologieMod.asp?ID=" & ID
	else
		Classe.Next_Page = "OrdiniRigheTipologie.asp"
	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>