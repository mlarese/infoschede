<%
class ricerca

'variabili e proprieta' locali 
Public fields_ricerca 		'array di campi visibili nella ricerca il nome e' quello del field
Public fields_tipo 			'il tipo di campo
Public fields_dicit
Public fields_selquery 		'la query per una SELECT
Public fields_scelta
Public fields_selID 		'l'id della tabella da cui la SELECT
Public fields_iniziale
Public page_this
Public connessione
Public rec_mess

Private conn
Private rs

Public sub inizializza(ByVal qfel)
	fields_ricerca = Array(qfel)
	fields_tipo = Array(qfel)
	fields_dicit = Array(qfel)
	fields_selquery = Array(qfel)
	fields_selID = Array(qfel) 
	fields_scelta = Array(qfel)
	fields_iniziale = Array(qfel)
	redim preserve fields_ricerca(qfel)
	redim preserve fields_tipo(qfel)
	redim preserve fields_dicit(qfel)
	redim preserve fields_selquery(qfel)
	redim preserve fields_selID(qfel)
	redim preserve fields_scelta(qfel)
	redim preserve fields_iniziale(qfel)
end sub

Public Sub creaform_unariga()
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open connessione,"",""
	response.write "<FORM action='"+page_this+"' method='POST' name='formRicerca'>"
	response.write "<table width='735' cellspacing='1' cellpadding='0' border='0' class='ricerca' style=""margin-bottom:4px;"">"
	response.write "<tr><td bgcolor='#E6E6E6' colspan='" & ubound(fields_ricerca)+1 & "' style='border-bottom:1px solid Gray;'><font class='testo11b'>"+rec_mess+"</font>"
	response.write "<tr>"
	for a = 1 to ubound(fields_ricerca)
		response.write "<td><font class='testo11n'>&nbsp;"+fields_dicit(a)+"&nbsp;</font>"
		if  fields_tipo(a) = "SELECT" then
		    response.write "<select name='"+fields_ricerca(a)+"' class='sel_ricerca'>"
			set rs = conn.execute(fields_selquery(a))
			response.write "<option value=''>"+fields_iniziale(a)+"</option>"
			do while not rs.EOF
				stringa_valore = rs(fields_ricerca(a))+"//"+cstr(rs(fields_selID(a)))
				if fields_scelta(a) = stringa_valore then
					response.write "<option value='"+stringa_valore+"' selected>"+rs(fields_ricerca(a))+"</option>"
				else
					response.write "<option value='"+stringa_valore+"'>"+rs(fields_ricerca(a))+"</option>"
				end if
				rs.MoveNext
			loop
			response.write "</select></td>"
		end if
		if  fields_tipo(a) = "INPUT" then
			response.write "<input type='text' name='"+fields_ricerca(a)+"' value='"+fields_scelta(a)+"' size='30' class='ricerca'></td>"
		end if
	next
	response.write "<td align='right'><input type='submit' name='invia_ric' value='TROVA' class='puls_1'>&nbsp;</td></tr></form></table>"
	conn.close
	conn = null
end sub


end class
 %>