<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
'controllo permessi
CALL CheckAutentication(session("PASS_ADMIN") <> "")

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_sid_nome_it;tft_sid_codice"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= "ApplicazioniParams.asp"
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_siti_descrittori"
	Classe.id_Field					= "sid_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql
	
	'controllo univocita codice
	sql = " SELECT COUNT(*) FROM tb_siti_descrittori"& _
		  " WHERE sid_codice LIKE '"& request.form("tft_sid_codice") &"'"& _
		  " AND sid_id <> "& ID
	if CIntero(GetValueList(conn, rs, sql)) > 0 then
		session("ERRORE") = "Il codice inserito appartiene ad un altro descrittore."
	else
		dim car
		car = request.form("car")
		
		'inserisco nuove relazioni
		if car <> "" then
			sql = " INSERT INTO rel_siti_descrittori(rsd_sito_id, rsd_descrittore_id)"& _
				  " SELECT id_sito, "& ID &" FROM tb_siti"& _
				  " WHERE id_sito IN ("& car &")"& _
				  " AND NOT EXISTS (SELECT 1 FROM rel_siti_descrittori WHERE rsd_sito_id = id_sito AND rsd_descrittore_id = "& ID &")"
			conn.Execute(sql)
		end if
		
		'cancello vecchie
		sql = " DELETE FROM rel_siti_descrittori"& _
			  " WHERE rsd_descrittore_id = "& ID
		if car <> "" then
			sql = sql &" AND NOT rsd_sito_id IN ("& car &")"
		end if
		conn.Execute(sql)
	end if
	
	Classe.isReport = FALSE
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
%>