<% 
'*******************************************************************************
'classe per la configurazione degli oggetti
'*******************************************************************************
'DESCRIZIONE METODI:*********************
'AddDefault (key, value)				: aggiunge una poprieta' di nome "KEY" e valore "value" alla collezione delle proprieta'
'SetConfigurationString(ByVal vData)	: imposta la stringa di configurazione e la scompone inserendone ogni termine 
'											nella collezione delle proprieta'
'EncodePage(number)						:converte l'id paginasito in id pagina definitivo
'DecodePage(number)						:converte l'id di pagina definitivo in id pagina sito
'toString()								:restituisce la stringa di configurazione
										

'DESCRIZIONE PROPRIETA'*******************
'Count									: restituisce il numero di proprieta' della collezione
'Item(Key)								: restituisce il valore della proprieta' nella collezione con chiave "KEY"
'											se non trova il valore restituisce ""
'PageName								: restitiuisce il nome della pagina corrente
'Lingua									:restituisce / imposta la lingua corrente
'ImageURL								:restituisce / imposta il percorso di base delle immagini
'BaseURL								:restituisce / imposta il percorso di base dell'applicazione
'SecureURL								:restituisce / imposta il percorso di base dell'applicazione su server sicuro
'SecurePostUrl							:restitusce l'indirizzo per passare in modalita' sicura: con indirizzo sicuro attivo, altrimenti restituisce indirizzo corrente.

class Configuration
	'variabili interne
	private Configuration_String
	private Properties
	Private VettorePagine
	Private Configuration_Loaded
	
	'variabili di cache utilizzate per velocizzare richieste successive dello stesso valore
	Private CACHE_PageName				
	Private CACHE_DecodePage_IN
	Private CACHE_DecodePage_OUT
	Private CACHE_GetItem_IN
	Private CACHE_GetItem_OUT
	
	Public Lingua
	Public UploadURL
	Public ImageURL
	Public OriginalImageURL
	Public BaseURL
	Public SecureURL
		
	Private Sub Class_Initialize()
		Configuration_String = ""
		Configuration_Loaded = FALSE
		
		'crea oggetto per collezione proprieta'
		set Properties = Server.CreateObject("Scripting.Dictionary")
		Properties.CompareMode = vbTextCompare
		
		'imposta lingua
		Lingua = lcase(Session("LINGUA"))
		If Lingua = "" then
			Lingua = LINGUA_ITALIANO
		end if
			
		'imposta il percorso di base delle immagini
		'imposta http sicuro
		if instr(1,Request.ServerVariables("HTTPS"),"on",vbTextCompare) then
			if Application("SECURE_IMAGE_SERVER")<>"" then
				UploadURL = "https://" & Application("SECURE_IMAGE_SERVER")
			else
				UploadURL = "https://" & Application("IMAGE_SERVER")
			end if
		else
			UploadURL = "http://" & Application("IMAGE_SERVER")
		end if	
		ImageURL = UploadURL & "/" & Session("AZ_ID")  & "/images/"
		OriginalImageURL = ImageURL
		
		BaseURL = "http://" & Application("SERVER_NAME") & "/"
		
		if Application("SECURE_SERVER_NAME")<>"" then
			SecureURL = "https://" & Application("SECURE_SERVER_NAME") & "/"
		else
			if instr(1, request.ServerVariables("LOCAL_ADDR"), "192.168.", vbTExtCompare)> 0 then
				SecureURL = "http://" & Application("SERVER_NAME") & "/"
			else
				SecureURL = "https://" & Application("SERVER_NAME") & "/"
			end if
		end if
	End Sub
	
	Private Sub Class_Terminate()
		set Properties = nothing
	End Sub
	
'DEFINIZIONE METODI:*********************
	
	'procedura che imposta la stringa di configurazione
	Public sub SetConfigurationString(ByVal vData)
		dim prop_value_list, single_prop, equal_pos, key

		'esegue la pulizia della stringa
		Configuration_String = Clear4Property(vData)

		On error resume next
		'inserisce nell'oggetto dizionario tutti i termini delle proprieta'
		prop_value_list = replace(Configuration_String, ":=", "=")
		prop_value_list = split(prop_value_list, ";")
		for each single_prop in prop_value_list
			equal_pos = InStr(1, single_prop, "=", vbTextCompare)
			if equal_pos>0 then
				key = Trim(Left(single_prop, equal_pos - 1))
				if Properties.Exists(key) then
					Properties(Key) = Trim(Right(single_prop, Len(single_prop) - equal_pos))
				else
					Properties.Add Key, Trim(Right(single_prop, Len(single_prop) - equal_pos))
				end if
			end if
		next
		
		if err.number<>0 then
			response.write "ERRORE NELLA STRINGA DI CONFIGURAZIONE DELL'OGGETTO"
			response.end
		else
			Configuration_Loaded = TRUE
		end if
		
		'verifica se il plugin deve "lavorare" nel sito corrente o su un particolare sito
		if cInteger(Item("AZ_ID"))>0 then
			ImageUrl = UploadURL & "/" & Item("AZ_ID")  & "/images/"
		end if
		
		On error goto 0
		
	End sub
	
	'aggiunge una proprieta' alla collezione
	Public function AddDefault (key, value)
		if Properties.Exists(key) then
			'elemento gia' esistente : lo aggiorna
			AddDefault = false
			Properties(Key) = value
		else
			'elemento inesistente: lo aggiunge
			AddDefault = true
			Properties.Add Key, value
		end if
		if CACHE_GetItem_IN = key then
			CACHE_GetItem_OUT = value
		end if
	end function
	
	
	'converte l'id paginasito in id pagina definitivo
	Public function EncodePage(id_paginasito)
		if isNumeric(id_paginasito) then
			EncodePage = SESSION("PAGINE")(id_paginasito)
		else
			EncodePage = NULL
		end if
	end function
	
	
	'converte l'id pagina in id paginasito
	Public Function DecodePage(id_pagina)
		dim i
		if cInt(id_pagina) = cInt(CACHE_DecodePage_IN) then
			DecodePage = CACHE_DecodePage_OUT
		else
			for i = lbound(SESSION("PAGINE")) to ubound(SESSION("PAGINE"))
				if cInt(SESSION("PAGINE")(i)) = cInt(id_pagina) then
					CACHE_DecodePage_OUT = i
					Exit for
				end if
			next
			CACHE_DecodePage_IN = id_pagina
			DecodePage = CACHE_DecodePage_OUT
		end if
	end function
	
	
	'restituisce la stringa di configurazione ricavata dalle proprieta'
	Public Function toString()
		dim s, Key
		for each Key in Properties.Keys
			s = s & Key & "=" & Properties(key) & ";" & vbCrLf
		next
		toString = s
	end function
	
'DEFINIZIONE METODI ESTESI*******************
	Public function LoadParameters(conn, rs, ApplicationID)
		dim CreateRs, CreateConn
		'apre connessione se non inizializzata
		if not IsObject(conn) then
			Set conn = Server.CreateObject("ADODB.Connection")
			conn.open Application("DATA_ConnectionString"),"",""
			CreateConn = true
		else
			CreateConn = false
		end if
		'apre recordset se non inizializzato
		if not IsObject(rs) then
			set rs = server.createobject("adodb.recordset")
			CreateRs = true
		else
			CreateRs = false
		end if
		
		'carica parametri dell'applicazione
		sql = "SELECT par_key, par_value FROM tb_siti_parametri WHERE par_sito_ID=" & cIntero(ApplicationID)
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		while not rs.eof
			CALL AddDefault (rs("par_key").value, rs("par_value").value)
			rs.movenext
		wend
		rs.close
		
		'distrugge recordset se creato nella funzione
		if CreateRs then
			set rs = nothing
		end if
		
		'chiude e distrugge connessione se creato nella funzione
		if CreateConn then
			conn.close
			set conn = nothing
		end if
	end function

'DEFINIZIONE PROPRIETA'*******************

	'restituisce il numero di termini della configurazione
	Public Property Get Count()
    	Count = Properties.Count
	End Property
	
	
	'restituisce il valore dell'elemento richiesto
	Public Default Property Get Item(ByVal Key)
		'carica stringa di configurazione se non gia' eseguito
		if not Configuration_Loaded then
			SetConfigurationString(Session("CONFSTR"))
		end if
		if CACHE_GetItem_IN = Key then
			Item = CACHE_GetItem_OUT
		else
			CACHE_GetItem_IN = Key
			CACHE_GetItem_OUT = Properties(Key)
			Item = CACHE_GetItem_OUT
		end if
	end Property
	
	'verifica se esiste o meno il parametro richiesto
	Public Property Get Exists(ByVal Key)
		Exists = Properties.Exists(Key)
	end Property
	
	'restituisce la pagina corrente
	Public Property Get PageName()
    	if CACHE_PageName="" then
			CACHE_PageName = GetPageName()
		end if
		PageName = CACHE_PageName
	End Property
	
	'restituisce l'indirizzo a cui fare il post per "indirizzo sicuro" se attivo, altrimenti manda all'indirizzo normale
	Public Property Get SecurePostUrl()
		if Application("SECURE_SERVER_NAME")<>"" then
			SecurePostUrl = SecureUrl
		else
			SecurePostUrl = GetCurrentBaseUrl() & "/"
		end if
	end Property
	
'FUNZIONI PRIVATE************************
	
	'pulisce stringa di configurazione prima di parsing
	Private Function Clear4Property(ByVal Property_value)
	    Property_value = Replace(Property_value, vbCrLf, "")
	    Property_value = Replace(Property_value, vbTab, "")
	    Property_value = Replace(Property_value, "  ", "")
	    If Right(Property_value, 1) = ";" Then
	        'toglie ultimo ";" per evitare out of bound
	        Property_value = Left(Property_value, Len(Property_value) - 1)
	    End If
	    Clear4Property = Property_value
	End Function
	
	
end class
'*******************************************************************************
 %>