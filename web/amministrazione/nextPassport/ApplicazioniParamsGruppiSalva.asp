<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
'controllo permessi
CALL CheckAutentication(session("PASS_ADMIN") <> "")

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_sdr_titolo_it"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= "ApplicazioniParamsGruppi.asp"
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_siti_descrittori_raggruppamenti"
	Classe.id_Field					= "sdr_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql
	
	'controllo univocita codice
	sql = " SELECT COUNT(*) FROM tb_siti_descrittori_raggruppamenti"& _
		  " WHERE sdr_titolo_it LIKE '"& request.form("tft_sdr_titolo_it") &"'"& _
		  " AND sdr_id <> "& ID
	if CIntero(GetValueList(conn, rs, sql)) > 0 then
		session("ERRORE") = "Il titolo inserito appartiene ad un altro gruppo."
	else
		dim c
		for each c in Split(request.form("car"), ",")
			conn.Execute("UPDATE tb_siti_descrittori SET sid_raggruppamento_id = "& ID &" WHERE sid_id = "& c)
		next
	end if
	
	Classe.isReport = FALSE
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
%>