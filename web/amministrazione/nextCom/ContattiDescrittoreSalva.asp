<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="Tools_DocumentiFiles.asp" -->
<%

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_ict_nome_IT"
	Classe.Checkbox_Fields_List 	= "chk_ict_per_ricerca;chk_ict_per_confronto"
	Classe.Page_Ins_Form			= "ContattiDescrittoreNew.asp"
	Classe.Page_Mod_Form			= "ContattiDescrittoreMod.asp?ID="& request("ID")
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_indirizzario_carattech"
	Classe.id_Field					= "ict_id"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql, i, tipi

	if request("ID")<>"" then
		sql = "DELETE FROM rel_categ_ctech WHERE rcc_ctech_id=" & ID
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	end if
	tipi = split(request("categorie_associate"), ",")
	for i = lbound(tipi) to ubound(tipi)
		sql = "INSERT INTO rel_categ_ctech(rcc_categoria_id, rcc_ctech_id, rcc_ordine) VALUES (" & tipi(i) & ", " & ID & ", "&cIntero(request("rel_ordine_"&tipi(i)))&")"
		CALL conn.execute(sql, 0, adExecuteNoRecords)
	next
	
	'imposta parametri per passare alla pagina successiva
	Classe.isReport = FALSE
	Classe.Next_Page = "ContattiDescrittori.asp"
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>