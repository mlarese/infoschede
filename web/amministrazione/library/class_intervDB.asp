<%
class interventoDB

'variabili e proprieta' locali 


Public idRif
Public field_idRif
Public connessione
Public ilForm
Public doveDopo
Public nomeID
Public Checkbox_Fields
Public id_rec
Public Campi_obbligatori

Private operazione 		'quale operazione sul DB valori "I" = insert; "U" = update; "D" = delete
Private tabella 			'la tabella interessata
Private conn
Private rs
Private sql_1
Private sql_2
Private lo_id

Public sub opera(tipo,tab)
	dim sql_recup, rs_recup
	operazione = tipo
	tabella = tab
	
	if Check_Field(Campi_obbligatori) then
		incipitQuery operazione, tabella
		'response.write "sql_1 (dopo incipit) = " & sql_1 & "<br>"
		Select Case operazione
			Case "I"
				costruisciIns()
			Case "U"
				costruisciUp()
		end select
		if Checkbox_Fields <> "" then
			checkBox_Manage(operazione)
		end if
		CALL chiudiQuery(operazione)
		'response.write "sql_1 (dopo chiudi) = " & sql_1 & "<br>"
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.open connessione,"",""
		'response.write "sql_1 = " & sql_1 & "<br>"
		'response.end
		set rs = conn.execute(sql_1)
		'response.write "doveDopo = " & doveDopo & "<br>"
		'response.end
		Select Case operazione
			Case "I"
				sql_recup = "SELECT MAX("+nomeID+") FROM " + tabella
				set rs_recup = conn.execute(sql_recup)
				lo_id = rs_recup(0)
				if doveDopo<>"" then
					if InStr(1,doveDopo,"?",1) > 0 then
						response.redirect doveDopo
					else
						response.redirect doveDopo & "?ID=" & lo_id
					end if
				end if
			Case "U"
				if doveDopo<>"" then
					'response.end
					if InStr(1,doveDopo,"?",1) > 0 then
						response.redirect doveDopo
					else
						response.redirect doveDopo & "?ID=" & lo_id
					end if				
				end if
			Case "D"
				if doveDopo<>"" then
					response.redirect doveDopo
				end if
		end select
		id_rec = lo_id
	else
		Session("ERRORE") = "Campi obbligatori non riempiti correttamente!"
	end if
end sub

Public sub aggiungi_IdInRelaz(tab,campi,valori)
	incipitQuery "I", tab
	i_campi = Split(campi, ";")
	i_valori = Split(valori, ";")
	for a = 0 to ubound(i_campi)
		agg_numerico i_campi(a), i_valori(a), "I"
	next
	chiudiQuery "I"
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open connessione,"",""
	'response.write "sql_1_agg = " & sql_1 & "<br>"
	'response.end
	set rs = conn.execute(sql_1)
	if doveDopo<>"" then
		if InStr(1,doveDopo,"?",1) > 0 then
			response.redirect doveDopo
		else
			response.redirect doveDopo & "?ID=" & lo_id
		end if
	end if
end sub

Public sub aggiungi()
	tabella = aggiunta_tab
	incipitQuery "I", tabella 
	costruisciIns()
	chiudiQuery "I"
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open connessione,"",""
	'response.write "sql_1 = " & sql_1 & "<br>"
	'response.end
	set rs = conn.execute(sql_1)
	'response.write "doveDopo = " & doveDopo & "<br>"
	'response.end
	if doveDopo<>"" then
		if InStr(1,doveDopo,"?",1) > 0 then
			response.redirect doveDopo
		else
			response.redirect doveDopo & "?ID=" & lo_id
		end if
	end if
end sub

Public sub inserisci(tab)
	
	if Check_Field(Campi_obbligatori) then
		dim sql_recup, rs_recup
		operazione = "I"
		tabella = tab
		'response.write "tipo = " & tipo & "<br>"
		incipitQuery "I", tabella 
		costruisciIns()
		'response.write "sql_1 = " & sql_1 & "<br>"
		if Checkbox_Fields <> "" then
			checkBox_Manage("I")
		end if
		chiudiQuery "I"
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.open connessione,"",""
		'response.write "sql_1 = " & sql_1 & "<br>"
		set rs = conn.execute(sql_1)
		sql_recup = "SELECT MAX("+nomeID+") FROM " + tabella
		set rs_recup = conn.execute(sql_recup)
		id_rec = rs_recup(0)
	else
		Session("ERRORE") = "Campi obbligatori non riempiti correttamente!"
	end if
end sub
'-------------------------------------------------------
'
'-------------------------------------------------------
Public sub opera_scelte_Rel(tab, valore1, nome, operaz)
	dim sql_recup, rs_recup, field, form_fields
	tabella = tab
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open connessione,"",""
	if operaz = "U" then
		'azzera le scelte precedenti
		sql_zero = "Delete From "+tabella+" Where "+nomeID+" = "+cstr(valore1)
		'response.write sql_zero
		set rs_zero = conn.execute(sql_zero)
	end if 
	'ora controlla i box
	Checkbox_Fields = request.Form(nome)
	'response.write "<br>" & Checkbox_Fields
	if InStr(1,Checkbox_Fields,";",1) > 0 then
		form_fields = Split(Checkbox_Fields, ";")
	else
		form_fields = Array(0)
		redim preserve form_fields(0)
		form_fields(0) = Checkbox_Fields
	end if
	'Response.write "<br>" & uBound(form_fields)
	for i=0 to uBound(form_fields)
		field = form_fields(i)
		'response.write "<br>"+field
		if request(Trim(form_fields(i)))<>"" then
			incipitQuery "I", tabella
			agg_numerico nomeID, valore1, "I"
		    agg_numerico field_idRif, field, "I"
			chiudiQuery "I"
			'response.write "<br>sql_1 = " & sql_1 & "<br>"
			set rs = conn.execute(sql_1)
		end if	
	next
	if doveDopo <> "" and Session("ERRORE")="" then
		response.redirect doveDopo
	end if
end sub



Private sub incipitQuery(tipo,tab)
	Select Case tipo
		Case "I"
			sql_1 = "INSERT INTO "+tab+" ("
			sql_2 = " VALUES ("	
		Case "U"
			sql_1 = "UPDATE "+tab+" SET " 
		Case "D"
			sql_1 = "DELETE FROM "+tab
	End Select
end sub

Private sub costruisciIns()
	dim var_field, prefisso, nome, contenuto
	for each var_field in request.form
		prefisso = left(var_field,3)
		nome = mid(var_field,5,len(var_field))
		contenuto = Request.Form(var_field)	
		'response.write var_field & " = " & contenuto & "<br>"	
		if contenuto <> "" then
			Select Case prefisso
				case "tft"
					agg_testo nome, contenuto, operazione		
				case "tfn"	
					agg_numerico nome, contenuto, operazione		
				case "rbu"	
					agg_radio nome, contenuto, operazione	
			end select
		end if
	next
end sub

Private sub costruisciUp()
	dim var_field, prefisso, nome, contenuto
	for each var_field in Request.Form
		prefisso = left(var_field,3)
		nome = mid(var_field,5,len(var_field))
		contenuto = Request.Form(var_field)
		'response.write "Request.Form(var_field) = " & Request.Form(var_field) & "<br>"
		'response.write var_field & "-contenuto = " & contenuto & "<br>"
		Select Case prefisso
			case "tft"
				agg_testo nome, contenuto, operazione	
			case "tfn"	
				agg_numerico nome, contenuto, operazione	
				'response.write "sono in TFN<br>"	
			case "rbu"	
				agg_radio nome, contenuto, operazione	
		end select
	next
end sub

Public sub chiudiQuery(tipo)
	Select Case tipo
		Case "I"
			sql_1 = left(sql_1,len(sql_1) - 2) & ")"
			sql_2 = left(sql_2,len(sql_2) - 2) & ")"
			sql_1 = sql_1 & sql_2
		Case "U"
			sql_1 = left(sql_1,len(sql_1) - 2)
			sql_1 = sql_1 & " Where " & field_idRif & " = " & cstr(idRif)
		Case "D"
			sql_1 = sql_1 & " Where " & field_idRif & " = " & cstr(idRif)
	End Select
end sub

Private sub agg_testo(nome_campo, valore, oper)
	Select Case oper
		Case "I"
			sql_1 = sql_1 & nome_campo + ", "
			sql_2 = sql_2 & "'" & replace(valore,"'","''",1) & "', "
			'response.write 	"nome_campo = " & nome_campo & "; valore = " & valore & "<br>"
		Case "U"
			sql_1 = sql_1 & nome_campo & " = '" & replace(valore,"'","''",1) & "', "
			'response.write 	"valore = " & valore & "<br>"
			'response.write 	"sql_1 = " & sql_1 & "<br>"
	End select
end sub	
	
Private sub agg_numerico(nome_campo, valore, oper)
	if valore = "" then
		valore = "NULL"
	end if
	valore = parsesql(valore,adNumeric)
	Select Case oper
		Case "I"
			sql_1 = sql_1 & nome_campo + ", "
			sql_2 = sql_2 & valore & ", "
		Case "U"
			sql_1 = sql_1 & nome_campo & " = " & valore & ", "
	End select
	'response.write 	"nome_campo = " & nome_campo & "; valore = " & cstr(valore) & "<br>"
end sub	

Private sub agg_box(nome_campo, oper)
	Select Case oper
		Case "I"
			sql_1 = sql_1 & nome_campo + ", "
			sql_2 = sql_2 & "1, "
		Case "U"
			sql_1 = sql_1 & nome_campo & " = 1, "
	End select
end sub	

Private sub agg_box_null(nome_campo, oper)
	Select Case oper
		Case "I"
			sql_1 = sql_1 & nome_campo + ", "
			sql_2 = sql_2 & "0, "
		Case "U"
			sql_1 = sql_1 & nome_campo & " =0, "
	End select
end sub			
	
Private sub agg_radio(nome_campo, valore, oper)
	Select Case oper
		Case "I"
			sql_1 = sql_1 & nome_campo + ", "
			sql_2 = sql_2 & cstr(valore) & ", "
		Case "U"
			sql_1 = sql_1 & nome_campo & " = " & cstr(valore) & ", "
	End select
end sub			

Private Sub checkBox_Manage(oper)
	dim form_fields, i, field
	if InStr(1,Checkbox_Fields,";",1) > 0 then
		form_fields = Split(Checkbox_Fields, ";")
	else
		form_fields = Array(0)
		redim preserve form_fields(0)
		form_fields(0) = Checkbox_Fields
	end if
	for i=0 to uBound(form_fields)
		field = form_fields(i)
		if request(Trim(form_fields(i)))<>"" then
		    agg_box field, oper
		else
			agg_box_null field, oper
		end if
	next
end Sub

Private function Check_Field(Requested_Fields)
	dim fields,i, num_empty
	num_empty = 0
	fields = Split(Requested_Fields, ";")
	for i=0 to uBound(fields)
		if request(Trim(fields(i)))="" then
			'response.write fields(i)
			num_empty = num_empty + 1
		end if
	next
	Check_Field = (num_empty = 0)
	'response.end
end function

'----------------------------------------------------------------
' opera_scelte_Rel_ConValore: aggiunge due valori
'----------------------------------------------------------------
Public sub opera_scelte_Rel_ConValore(tab, valore, valore1, nome, operaz )
	dim sql_recup, rs_recup, field, form_fields
	tabella = tab

	if Check_Field(Campi_obbligatori) then
		
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.open connessione,"",""
		
		if operaz = "U" then
			'azzera le scelte precedenti
			sql_zero = "Delete From "+tabella+" Where "+nomeID+" = "+cstr(valore1)
			'response.write sql_zero
			set rs_zero = conn.execute(sql_zero)
		end if 
		'ora controlla i box
		Checkbox_Fields = request.Form(nome)
		'response.write "<br>" & Checkbox_Fields
		if InStr(1,Checkbox_Fields,";",1) > 0 then
			form_fields = Split(Checkbox_Fields, ";")
		else
			form_fields = Array(0)
			redim preserve form_fields(0)
			form_fields(0) = Checkbox_Fields
		end if
		'Response.write "<br>" & uBound(form_fields)
		for i=0 to uBound(form_fields)
			field = form_fields(i)
			'response.write "<br>"+field
			if request(Trim(form_fields(i)))<>"" then
				incipitQuery "I", tabella
				agg_numerico nomeID, valore1, "I"
			    agg_numerico field_idRif, field, "I"
				agg_numerico valore, request(Trim(form_fields(i))), "I"
			    'agg_numerico 
				chiudiQuery "I"
				'response.write "<br>sql_1 = " & sql_1 & "<br>"
				set rs = conn.execute(sql_1)
			end if	
		next
		if doveDopo <> "" and Session("ERRORE")="" then
			response.redirect doveDopo
		end if
	else
		Session("ERRORE") = "Campi obbligatori non riempiti correttamente!"
	end if
end sub

end class
 %>