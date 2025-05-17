<% 
'funzione che ritorna l'url del menu selezionato
function getHREF(Config, rs, gruppo)
	If cInteger(rs("id_pagina")) > 0 Then
		getHREF= "http://" + Application("SERVER_NAME") & "/dynalay.asp" & _
				 "?PAGINA=" & Config.EncodePage(rs("id_pagina")) & "&" & gruppo & "=" & rs("id_MenuItem")
	else
		getHREF= rs("link_menuitem_" & Config.lingua)
	end if
end function


'funzione che ritorna l'indice del menu selezionato
function getMenuItemSelected(conn, rs, Config, Gruppo)
	dim sql
	getMenuItemSelected = 0
	if cInteger(request(gruppo))>0 then
		'voce di menu indicata direttamente dal request
		getMenuItemSelected = cInteger(request(gruppo))
	end if
	if getMenuItemSelected = 0 then
		'verifica nelle pagine, nella lingua corrente se la pagina &egrave; collegata a qualche menu.
		sql = " SELECT id_menuItem FROM tb_menuItem " +_
			  " WHERE id_pagina = " & cIntero(Session("CURRENT_PAGINASITO")) & _
			  " AND id_link=" & cInteger(Config("MenuID"))
		rs.open sql, conn, adOpenforwardOnly, adLockReadOnly, adCmdText
		if not rs.eof then
			getMenuItemSelected = cInteger(rs("id_menuItem"))
		end if
		rs.close
		if getMenuItemSelected = 0 then
			'non trovata tra le pagine collegate ai menu: verifica indicazione di sessione (se utilizzata)
			if Gruppo<>"" then
				if cInteger(Session("Current_MIid_" + Gruppo))<>"" then
					getMenuItemSelected = cInteger(Session("Current_MIid_" + Gruppo))
				end if
			end if
		end if
	end if
	'imposta la variabile di sessione corrispondente alla voce di menu
	if Gruppo<>"" then
		Session("Current_MIid_" + Gruppo) = getMenuItemSelected
	end if
end function

%>