<%
'*******************************************************************************************************************
'*******************************************************************************************************************
'DEFINIZIONE CLASSE
'*******************************************************************************************************************
class OBJ_Salva

'variabili e proprieta' locali

'variabili utilizzate nelle proprieta'
Private Ins_Page					'pagina form di inserimento
Private Mod_Page					'pagina form di modifica
Public Requested_Fields_List		'lista campi obbligatori (separati da ";")
Public Checkbox_Fields_List			'lista campi di tipo checkbox (separati da ";")
Public Gestione_Relazioni			'(Default = FALSE) Indica se richiamare la funzione per la gestione delle relazioni
Public Next_Page					'Pagina successiva
Public Next_Page_ID					'(Default = FALSE) Indica se deve esseer passato l'ID alla pagina successiva
Public Table_Name					'Nome tabella
Public id_Field						'nome campo ID
Public ID_value						'opzionale: forza il valore dell'id in modifica
Public Read_New_ID					'(default = FALSE) Indica se deve essere letto ID dopo l'inserimento
Public IsReport						'(default = FALSE) Indica se la pagina produce un report.
Public ConnectionString				'Stringa per la connessione al database
Public Conn							'Connessione utlilizzata per classe salva
Public ForcedValues					'oggetto dictionary che contiene i campi con i rispettivi valori da forzare.

'variabili interne
Private rs, sql, prev_page, form_field, field

Private Sub Class_Initialize()
	Next_Page_ID = False
	Read_New_ID = False
	Gestione_Relazioni = False
	IsReport = FALSE
	ID_value = cLng("0" & request.Querystring("ID"))
end sub

private sub Class_Terminate()

end sub

'************************************************************************************************************
'FUNZIONI DI INTERFACCIA PUBBLICA
'******************************************************

Public Sub Salva()
	'converte il valore dell'ID in intero
	ID_value = cLng("0" & ID_value)
	
	'controllo campi obbligatori
	CALL Controllo()
	if Session("ERRORE")<>"" then
		if not isReport then
			response.redirect prev_page
		else
			Exit Sub
		end if
	end if
	
	'Apertura recordset e connessioni
	if isEmpty(conn) then
		set conn = Server.CreateObject("ADODB.Connection")
		conn.open ConnectionString,"",""
		conn.begintrans
	end if

	set rs = Server.CreateObject("ADODB.RecordSet")

	if ID_value > 0 then
		'imposta query per modifica
		sql = "SELECT * FROM " & table_name & " WHERE " & id_field & "=" & ID_value
	else
		'imposta query per inserimento
		sql = "SELECT TOP 1 * FROM " & table_name
	end if
	rs.open sql, conn, adOpenKeySet, adLockOptimistic

	if ID_value > 0 AND rs.Recordcount = 0 then 
		Session("ERRORE") = "MODIFICA NON RIUSCITA: RECORD NON TROVATO!"
		if not isReport then
			conn.rollbacktrans
			response.redirect prev_page
		else
			Exit Sub
		end if
	elseif ID_value = 0 then
		'aggiunge record se in inserimento
		rs.AddNew
	end if
	
	'Aggiornamento campi
	for each form_field in request.Form
		'response.write "form_field : " & form_field & "<br> value=""" & request(form_field) & """<br>"
		if len(form_field)>4 then
			field = right(form_field, len(form_field)-4)
			'response.write "field:" & field & "<br>"
			if instr(1, left(form_field, 4), "tft_", vbTextCompare)>0 then
				'campo testo
				'rs(field) = Replace(request(form_field), """", "'") 'modifica disattivata il 30/05/2013
				rs(field) = request(form_field) 
			elseif instr(1, left(form_field, 4), "tfd_", vbTextCompare)>0 then
				'campo di tipo data
				rs(field) = ConvertForSave_Date(request(form_field))
			elseif instr(1, left(form_field, 4), "fdt_", vbTextCompare)>0 then
				'campo di tipo orario
				rs(field) = ConvertForSave_Time(request(form_field))
			elseif instr(1, left(form_field, 4), "tfn_", vbTextCompare)>0 then
				'campo di tipo numerico con valore default 0
				rs(field) = ConvertForSave_Number(request(form_field), 0)
			elseif instr(1, left(form_field, 4), "nfn_", vbTextCompare)>0 then
				'campo di tipo numerico nullabile
				rs(field) = ConvertForSave_Number(request(form_field), NULL)
			elseif instr(1, left(form_field, 4), "tfh_", vbTextCompare)>0 then
				'campo HTML
				rs(field) = MakeRelativeLink(request(form_field))
			end if
		end if
	next
   
	'gestione campi checkbox
	if Checkbox_Fields_List<>"" then
		CALL checkBox_Manage()
	end if
		
	'gestione campi con valore impostato
	if not isEmpty(ForcedValues) then
		CALL UpdateForcedFields(rs)
	end if

	'registrazione modifiche record
	rs.Update
	'response.end

	'Lettura Nuovo ID Inserito se non in modifica
	If Read_New_ID AND ID_value=0 then
		ID_value = cLng("0" & rs(id_field))
		if ID_value = 0 then	'se non impostato automaticamente lo ricava
			rs.close
			sql = "SELECT MAX(" & id_field & ") AS ID FROM " & table_name
			rs.open sql, conn, adOpenStatic, adLockOptimistic
            if not rs.eof then
				ID_value = rs("ID")
			else
				Session("ERRORE") = "ERRORE NELL'INSERIMENTO DEL RECORD."
				if not isReport then
					conn.rollbacktrans
					
					response.redirect prev_page
				else
					Exit Sub
				end if
			end if
		end if
	end if
	rs.close

	'gestione relazioni
	if Gestione_Relazioni then
		CALL Gestione_Relazioni_Record(conn, rs, ID_value)
		if Session("ERRORE")<>"" then
			if not isReport AND prev_page<>"" then
				conn.rollbacktrans
				response.redirect prev_page
			else
				Exit Sub
			end if
		end if
	end if
	
	'Imposta parametro per pagina successiva
	if Next_Page_ID and ID_value>0 then
		if instr(1, next_page, "?", vbTextCompare)>0 then
			next_page = next_page + "&" & UCASE(id_Field) & "=" & ID_value
		else
			next_page = next_page + "?" & UCASE(id_Field) & "=" & ID_value
		end if
	end if
	
	'chiusura connessioni
	conn.committrans
	set rs = nothing
	conn.close
	set conn = nothing
	if not IsReport then
		response.redirect next_page
	end if
end sub


'************************************************************************************************************
'FUNZIONI PRIVATE'
'******************************************************
Public sub Controllo()
	dim fields,i, num_empty, what_empty
	num_empty = 0
	what_empty = ""
	fields = Split(Requested_Fields_List, ";")
	for i=0 to uBound(fields)
		if Trim(fields(i))<>"" then
			if request(Trim(fields(i)))="" then
				num_empty = num_empty + 1
				what_empty = what_empty & "<font style=""font-size:11px;"">""" & fields(i) & """&nbsp;&nbsp;&nbsp;</font>"
			end if
		end if
	next
	if num_empty>0 then
		Session("ERRORE") = ChooseValueByAllLanguages(Session("LINGUA"), "Numero ", "", "", "", "", "", "", "") & num_Empty & ChooseValueByAllLanguages(Session("LINGUA"), " campi obbligatori non riempiti correttamente!", " mandatory fields not filled properly!", "", "", "", "", "", "")
		
		if instr(1, request.ServerVariables("SERVER_NAME"), ".local", vbTextCompare)>0 then
			Session("ERRORE") = Session("ERRORE") & "<br>" & what_empty 
		end if
	end if
end sub

Private Sub checkBox_Manage()
	dim form_fields, i, field
	form_fields = Split(Checkbox_Fields_List, ";")
	for i=0 to uBound(form_fields)
'response.write """" & form_fields(i) & """=""" & request(Trim(form_fields(i))) & """<br>" 
		if instr(1, form_fields(i), "chk_", vbTextCompare)>0 then
			field = right(Trim(form_fields(i)), len(Trim(form_fields(i)))-4)
		else
			field = Trim(form_fields(i))
		end if
'response.write """" & field & """<br>"
		if field<>"" then
			if request(Trim(form_fields(i)))<>"" then
				rs(field) = True
			else
				rs(field) = False
			end if
		end if
	next
end Sub



'************************************************************************************************************
'GESTIONE PROPRIETA'
'******************************************************

'******************************************************
'Pagina Form Inserimento dati
Public Property Let Page_Ins_Form(ByVal vData)
    Ins_Page = vData
	if ID_value=0 then
		prev_page = Ins_Page
	end if
End Property

Public Property Get Page_Ins_Form()
    Page_Ins_Form = Ins_Page
End Property


'******************************************************
'Pagina Form modifica dati
Public Property Let Page_Mod_Form(ByVal vData)
    Mod_Page = vData
	if ID_value>0 then
		prev_page = Mod_Page
	end if
End Property

Public Property Get Page_Mod_Form()
    Page_Mod_Form = Mod_Page
End Property


'************************************************************************************************************
'GESTIONE VALORI FORZATI'
'******************************************************
'******************************************************
'metodo per l'aggiunta di valori forzati
Public Function AddForcedValue(field, value)
	if isEmpty(ForcedValues) then
		'crea oggetto dictionary per valori forzati
		set ForcedValues = Server.CreateObject("Scripting.Dictionary")
		ForcedValues.CompareMode = vbTextCompare
	end if
    if right(field, 1) = "_" then
        dim lingua
        for each lingua in Application("LINGUE")
            if not ForcedValues.Exists(field & lingua) then
	    	    CALL ForcedValues.Add(field & lingua, value)
        	end if
        next
    else
    	if not ForcedValues.Exists(field) then
	    	CALL ForcedValues.Add(field, value)
    	end if
    end if
end function

'aggiorna i valori del recordset
Private Function UpdateForcedFields(rs)
	if not isEmpty(ForcedValues) then
		dim field
		for each Field in ForcedValues.Keys
'			response.write "<br>"& field
			rs(Field) = ForcedValues(Field)
		next
	end if
end function


'************************************************************************************************************
'GESTIONE CAMPI PER TRACCIATURA INSERIMENTO E MODIFICA'
'******************************************************
'imposta i valori dei campi per l'impostazione dei dati di tracciatura e modifica
Public Function SetUpdateParams(prefix)
	dim tempo
	tempo = Now()
	if cInteger(ID_value) = 0 then
		CALL AddForcedValue(prefix & "insData", tempo)
		CALL AddForcedValue(prefix & "insAdmin_id", Session("ID_ADMIN"))
	end if
	CALL AddForcedValue(prefix & "ModData", tempo)
	CALL AddForcedValue(prefix & "ModAdmin_id", Session("ID_ADMIN"))
end function

'******************************************************
end class

%>