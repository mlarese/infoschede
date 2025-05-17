<%

public sub Save(conn, rs)
	
	dim sql
	'inserisce record messaggio
	sql = "SELECT TOP 1 * FROM tb_email"
	rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
	rs.AddNew
	rs("email_object") = Subject
	rs("email_mime") = MimeType
	if MimeType = MIME_HTML then
		rs("email_text") = HtmlBody
	else
		rs("email_text") = Body
	end if
	
	rs("email_data") = Now
	rs("email_dipgenera") = SenderID
	rs("email_tipi_messaggi_id") = MessageType
	rs("email_control_key") = GetRandomString(ALPHANUMERIC_CHARSET, 8)
	rs.Update
	
	MessageId = rs("email_id")

	rs.close
end sub


Public  Function SendToContact(conn, rs, cntID)
	dim sql, recapito, name
	
	'recupera email del contatto destinatario
	sql = " SELECT TOP 1 ValoreNumero FROM tb_valoriNumeri WHERE id_indirizzario=" & cIntero(cntID) & _
		  " AND id_tipoNumero=" & RecipientType & " AND " & SQL_IsTrue(conn, "email_default")
	recapito = GetValueList(Conn, rs, sql)
	
	sql = "SELECT * FROM tb_indirizzario WHERE idElencoIndirizzi = "& cIntero(cntID)
	rs.open sql, conn, adOpenStatic, adLockOptimistic
	name = ContactFullName(rs)
	rs.close
	
	SendToContact = SendSave(conn, cntID, recapito, name)
	
End Function


Public Function SendToAdmin(conn, rs, adminID)
	dim sql, email, name
	
	'recupera email del contatto destinatario
	sql = " SELECT * FROM tb_admin WHERE id_admin = " & adminID
	rs.open sql, conn, adOpenStatic, adLockOptimistic
	email = rs("admin_email")
	name = Trim(rs("admin_cognome") & " " & rs("admin_nome"))
	rs.close
	
	SendToAdmin = SendAdminSave(conn, adminID, email, name)
	
End Function


Public Function SendSave(conn,  cntID, cntRecapito, cntName)
	dim sql, errorNumber
	
	sql = "SELECT COUNT(*) FROM log_cnt_email WHERE log_email_id = "& cIntero(MessageID) &" AND log_email LIKE '"& ParseSql(cntRecapito, adChar) &"'"
	if cIntero(GetValueList(conn, NULL, sql)) = 0 then
		errorNumber = SendOnError(cntRecapito)
	else
		errorNumber = 0
	end if
	
	'inserisce riga su log
	sql = " INSERT INTO log_cnt_email (log_cnt_id, log_email_id, log_email, log_cnt_nominativo, log_inviato_ok) " & _
		  " VALUES (" & cIntero(cntID) & ", " & cIntero(MessageID) & ", '" & ParseSQL(cntRecapito, adChar) & "', '"& ParseSQL(cntName, adChar) &"', "& _
		  IIF(errorNumber <> 0, "0", "1") &")"
	CALL conn.execute(sql, , adExecuteNoRecords)
	
	SendSave = err.Number > 0
End Function


Public Function SendAdminSave(conn, adminId, adminEmail, adminName)
	dim sql, errorNumber
	
	sql = "SELECT COUNT(*) FROM log_cnt_email WHERE log_email_id = "& cIntero(MessageID) &" AND log_email LIKE '"& ParseSql(adminEmail, adChar) &"'"
	if CInt(GetValueList(conn, NULL, sql)) = 0 then
		errorNumber = SendOnError(adminEmail)
	else
		errorNumber = 0
	end if
	
	'inserisce riga su log
	sql = " INSERT INTO log_cnt_email (log_cnt_id, log_email_id, log_email, log_cnt_nominativo, log_inviato_ok) " & _
		  " VALUES (NULL , " & cIntero(MessageID) & ", '" & ParseSQL(adminEmail, adChar) & "', '"& ParseSQL(adminName, adChar) &"', "& _
		  IIF(errorNumber <> 0, "0", "1") &")"
	CALL conn.execute(sql, , adExecuteNoRecords)
	
	SendAdminSave = err.Number > 0
end function


%>