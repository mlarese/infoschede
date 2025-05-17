<%

'............................................................................................................................................................
'..		funzione che copia i valori del record sorgente in quello destinazione, saltando i campi indicati
'............................................................................................................................................................
sub RecordsetCopyFields(rsSource, rsDest, FieldsToSkip)
	dim field
	for each field in rsSource.Fields
		if not instr(1, "," & replace(FieldsToSkip, " ", "") & ",", "," & field.name & ",", vbTextCompare)>0 then
			rsDest(field.name) = rsSource(field.name)
		end if
	next
end sub

%>