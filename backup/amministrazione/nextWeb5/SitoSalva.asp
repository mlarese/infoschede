<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 1000000 %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<!--#INCLUDE FILE="SitoAnalisiStat_TOOLS.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_siti_gestione, 0))

dim Classe, conn
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	'Classe.ConnectionString 		= Application("DATA_ConnectionString")
	set Classe.conn = conn
	Classe.Requested_Fields_List	= "tft_nome_webs;tft_URL_base"
	Classe.Checkbox_Fields_List 	= "chk_lingua_EN; chk_lingua_FR; chk_lingua_DE; chk_lingua_ES;" & _
									  IIF(DB_Type(Classe.conn) = DB_ACCESS, "", "chk_lingua_RU; chk_lingua_CN; chk_lingua_PT;")
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_webs"
	Classe.id_Field					= "id_webs"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
	set Classe.Conn = index.conn
	index.conn.BeginTrans
	
	if request("ID") = "" then
		'azzera contatori per nuovo sito
		CALL Classe.AddForcedValue("contatore", 0)
		CALL Classe.AddForcedValue("contUtenti", 0)
		CALL Classe.AddForcedValue("contCrawler", 0)
		CALL Classe.AddForcedValue("contAltro", 0)
	end if
	
	
	'update campi gestione modifiche
	classe.SetUpdateParams("webs_")

'definizione eventuali operazioni su relazioni
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim co_id
	
    'verifica se il parametro id e' stato inserito correttamente
    if cIntero(id)>0 then
    	'controllo selezione lingua iniziale
    	if request.form("chk_lingua_"& request.form("tft_lingua_iniziale")) = "" AND request.form("tft_lingua_iniziale") <> "it" AND request.form("tft_lingua_iniziale") <> "" then
    		'imposta parametri per passare alla pagina successiva
    		Classe.isReport = TRUE
    		session("ERRORE") = "Hai selezionato una lingua iniziale non attiva!"
    	else
			If request.Querystring("ID")="" Then 
				'se in inserimento del sito crea directory fisica a partire dal template
    			Dim FSO, UploadPath, cssO
    			Set FSO = CreateObject("Scripting.FileSystemObject")
    			
    			UploadPath =  Application("IMAGE_PATH") & ID
    			if not FSO.FolderExists(UploadPath) then
    				CALL FSO.CreateFolder(UploadPath)
    				
					'crea cartelle interne
    				CALL FSO.CreateFolder(UploadPath + "\" + FILE_TYPE_XML)
    				CALL FSO.CreateFolder(UploadPath + "\" + FILE_TYPE_CSS)
    				CALL FSO.CreateFolder(UploadPath + "\" + FILE_TYPE_FLASH)
    				CALL FSO.CreateFolder(UploadPath + "\" + FILE_TYPE_IMAGE)
    				CALL FSO.CreateFolder(UploadPath + "\" + FILE_TYPE_TEXT)
					
    				'copia file di esempio testi
    				CALL FSO.CopyFile(Server.MapPath("../library/FilesTemplate/testodiprova.txt"), UploadPath + "\testi\testodiprova.txt", true)
    				
    			end if
    			
    			'ricostruzione degli stili
    			set cssO = new CssManager
    			CALL cssO.ResetDbToDefault(conn, ID)
    			set cssO = nothing
    			
    			set FSO = nothing
    		
	    		'modifico il contenuto
	    		set index.dizionario = server.createobject("Scripting.Dictionary")
	    		set index.content.dizionario = server.createobject("Scripting.Dictionary")
	    		index.content.dizionario("co_titolo_it") = request("tft_nome_webs")
	    		index.content.dizionario("co_visibile") = true
	    		index.content.co_F_key_id = ID
	    		index.content.co_F_table_id = index.GetTable("tb_webs")
	    		co_id = index.content.Salva(0)
				
	    		'modifico la voce
	            index.dizionario("idx_ordine") = ID
	    		index.dizionario("idx_content_id") = co_id
				index.dizionario("idx_principale") = false
	    		index.Salva(GetValueList(conn, rs, "SELECT idx_id FROM tb_contents_index WHERE idx_content_id = "& co_id))
    			
			else						
				if request("aggiorna_indice")<>"" then
					'..............................................................................
					'sincronizzazione con i contenuti e l'indice
					CALL Index_UpdateItem(conn, Classe.Table_Name, ID, false)
					'..............................................................................
				end if
				
    		End If
			
			'verifica se è stato cambiato lo stato di registrazione contatori, da attivo a non attivo.
			if cIntero(request("old_statistiche_attive")) <> cIntero(request("tfn_statistiche_attive")) then
				'stato di attivazione statistiche cambiato: registra situazione
				CALL StatisticheArchiviaAzzera(conn, ID)
			end if
			
    		'imposta parametri per passare alla pagina successiva
    		Classe.isReport = FALSE
    		Classe.Next_Page = "Siti.asp"
    	end if
    end if
    
end Sub

'salvataggio/modifica dati
Classe.Salva()

%>