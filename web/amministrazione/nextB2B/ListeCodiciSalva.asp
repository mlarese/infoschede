<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_LstCod_nome"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "gtb_lista_codici"
	Classe.id_Field					= "LstCod_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql
	'inserisce esploso dei codici
	if request("ID") = "" then
		sql = " INSERT INTO gtb_codici (Cod_lista_id, Cod_variante_id, Cod_Codice) " + _
			  " SELECT " & ID & ", rel_id, " 
		select case request("copia_da")
			case "i"		'genera da codice interno
				sql = sql + "rel_cod_int FROM grel_art_valori "
			case "a"		'genera codice alternativo
				sql = sql + "rel_cod_alt FROM grel_art_valori "
			case "p"		'genera codice produttore
				sql = sql + "rel_cod_pro FROM grel_art_valori "
			case "l"		'genera da altro listino
				sql = sql + " Cod_codice FROM grel_art_valori INNER JOIN gtb_codici " + _
							" ON grel_art_valori.rel_id = gtb_codici.cod_variante_id " + _
							" WHERE cod_lista_id=" & cIntero(request("copia_da_lista"))
			case else
				sql = ""
		end select
		if sql <> "" then
			CALL conn.execute(sql, , adExecuteNoRecords)
		end if
	end if
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	if request("salva_avanti")<>"" then
		Classe.Next_Page = "ListeCodiciCodici.asp?ID=" & ID
	else
		Classe.Next_Page = "ListeCodici.asp"
	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
	
	
%>