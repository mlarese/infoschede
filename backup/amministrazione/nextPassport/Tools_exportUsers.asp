<%
'*************************************************************************************************
'FUNZIONI DI EXPORT
'*************************************************************************************************

sub Export_Excel2000(conn, rs, sql, rsv)
	dim FileName, FSO, ExportPath, XLSConnString, XLSConn, XLSrs, Xsql, field, FileStream, rsN
	set rsN = server.CreateObject("ADODB.RecordSet")
	
	ExportPath = Application("IMAGE_PATH") & "temp\"
	FileName = Session.SessionID & ".XLS"
	
	'crea il file excel
	Set FSO = CreateObject("Scripting.FileSystemObject")
	
	FSO.CopyFile Server.Mappath("../library") & "/FilesTemplate/Master.xls", ExportPath + FileName , true

	'apre connessione a file EXCEL
	set XLSConn = Server.CreateObject("ADODB.Connection")
	XLSConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (ExportPath & FileName) &_
					";Extended Properties=""Excel 8.0;HDR=YES;"";Jet OLEDB:Engine Type=35;"
	XLSConn.open XLSConnString, "", ""
	
	'generazione sql per creazione tabella di export
	Xsql = "CREATE TABLE [Dati esportati] ( "
	for each field in rs.fields
		select case field.type
			case adInteger
				Xsql = Xsql & "[" & field.name & "] int, "
			case adBoolean
				Xsql = Xsql & "[" & field.name & "] bit, "
			case adDate 
				Xsql = Xsql & "[" & field.name & "] Date, "
			case adDBTimeStamp
				Xsql = Xsql & "[" & field.name & "] Date, "
			case adCurrency 
				Xsql = Xsql & "[" & field.name & "] Currency, "
			case adVarWChar
				if Field.DefinedSize>255 then
					Xsql = Xsql & "[" & field.name & "] varchar(" & IIF(Field.DefinedSize>250, 255, Field.DefinedSize) & "), "
				else
					Xsql = Xsql & "[" & field.name & "] text, "
				end if
		end select 
	next
	'aggiunge campi per numeri
	while not rsv.eof
		Xsql = Xsql & "[" & rsv("nome_tipoNumero") & "] text, "
		rsv.moveNext
	wend
	Xsql = Xsql & "Rubriche text)"
	
	'creazione tabella
	CALL XLSConn.execute(Xsql, 0, adExecuteNoRecords)
	
	'apre recordset su tabella
	set XLSrs = Server.CreateObject("ADODB.RecordSet")
	XLSrs.open "[Dati esportati]", XLSConn, adOpenStatic, adLockOptimistic, adCmdTable
	
	'riempie tabella
	while not rs.eof
		XLSrs.AddNew
		for each field in rs.fields
			XLSrs(field.name) = rs(field.name)
		next
		rsv.moveFirst
		while not rsv.eof
			field = rsv("nome_tipoNumero")
			sql = "SELECT ValoreNumero FROM tb_ValoriNumeri " &_
		 		  " WHERE id_TipoNumero=" & rsv("id_tipoNumero") & " AND  id_Indirizzario=" & rs("ID")
			XLSrs(field) = GetValueList(conn, rsN, sql)
			
			rsv.moveNext
		wend
		sql = " SELECT nome_rubrica FROM tb_rubriche " &_
		  	  " INNER JOIN rel_rub_ind ON tb_rubriche.id_rubrica=rel_rub_ind.id_rubrica " &_
			  " WHERE rel_rub_ind.id_indirizzo=" & rs("ID") 
		XLSrs("Rubriche") = cString(GetValueList(conn, rsN, sql))
		XLSrs.Update
		
		rs.movenext
	wend
	
	XLSrs.close
	XLSConn.close
	set XLSrs = nothing
	set XLSConn = nothing
	set rsN = nothing
	
	'reindirizza al file
	response.redirect "http://" & Application("IMAGE_SERVER") & "/temp/" & FileName
	
	'carica file su stream e lo restituisce
	
    'Set FileStream = Server.CreateObject("ADODB.Stream")
'    FileStream.Open
'   FileStream.Type = 1
'    FileStream.LoadFromFile(ExportPath & FileName)
'	
'	Response.Clear
'    response.contentType="application/vnd.ms-excel"
'	Response.BinaryWrite(FileStream.Read)
'
'	FileStream.Close
'    Set FileStream = Nothing
'	
'	'cancella file XLS creato
'	FSO.DeleteFile (ExportPath & FileName) , true
'	set FSO = nothing
	
end sub

sub Export_ExcelXP(conn, rs, sql, rsv)
	dim field, rsN
	set rsN = server.CreateObject("ADODB.RecordSet")
	response.contentType="application/vnd.ms-excel"
	response.clear%>
	<?xml version="1.0" encoding="UTF-8"?>
		<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">
	 		<DocumentProperties>
	  			<Author><%= Session("LOGIN_4_LOG") %></Author>
		  		<Created><%= DateIso(Date) %>T<%= Hour(date) %>:<%= Minute(date) %>:00Z</Created>
	  			<LastSaved><%= DateIso(Date) %>T<%= Hour(date) %>:<%= Minute(date) %>:00Z</LastSaved>
	 		</DocumentProperties>
			<Worksheet ss:Name="Dati esportati">
		  		<Table ss:DefaultColumnWidth="100">
			   		<Row>
			   			<% for each field in rs.fields%>
			    			<Cell><Data ss:Type="String"><%= field.name%></Data></Cell>
			   			<% next 
						while not rsv.eof%>
							<Cell><Data ss:Type="String"><%= rsv("nome_tipoNumero")%></Data></Cell>
							<%rsv.moveNext
						wend%>
						<Cell><Data ss:Type="String">Rubriche</Data></Cell>
			   		</Row>
			   		<%while not rs.eof%>
			    		<Row>
			   	 			<% for each field in rs.fields
					   			select case field.type
			   						case adInteger%>
										<Cell><Data ss:Type="Number"><%= rs(field.name) %></Data></Cell>
			    					<% case adBoolean
										if rs(field.name) then%>
											<Cell><Data ss:Type="String">Si</Data></Cell>
										<%else%>
											<Cell><Data ss:Type="String"></Data></Cell>
										<%end if
									case adDate
										if isDate(rs(field.name)) then%>
											<Cell><Data ss:Type="DateTime"><%= DateIso(rs(field.name)) %>T00:00:00.000</Data></Cell>
										<% else %>
											<Cell><Data ss:Type="String"></Data></Cell>
										<% end if %>
								<% 	case adCurrency %>
										<Cell><Data ss:Type="Number"><%= replace(rs(field.name), ",", ".") %></Data></Cell>
								<% 	case else %>
										<Cell><Data ss:Type="String"><%= Trim("" & rs(field.name)) %></Data></Cell>
								<% end select
				 			next
							rsv.moveFirst
							while not rsv.eof
								sql = "SELECT ValoreNumero FROM tb_ValoriNumeri " &_
							 		  " WHERE id_TipoNumero=" & rsv("id_tipoNumero") & " AND  id_Indirizzario=" & rs("ID")%>
								<Cell><Data ss:Type="String"><%= GetValueList(conn, rsN, sql) %></Data></Cell>
								<%rsv.moveNext
							wend
							sql = " SELECT nome_rubrica FROM tb_rubriche " &_
							  	  " INNER JOIN rel_rub_ind ON tb_rubriche.id_rubrica=rel_rub_ind.id_rubrica " &_
								  " WHERE rel_rub_ind.id_indirizzo=" & rs("ID") %>
							<Cell><Data ss:Type="String"><%= cString(GetValueList(conn, rsN, sql)) %></Data></Cell>
			   			</Row>
			    		<%rs.movenext
			   		wend %>
		  		</Table>
	 		</Worksheet>
		</Workbook>
	<%set rsN = nothing
end sub


sub Export_HTML(conn, rs, sql, rsv)
	dim Field, rsN
	set rsN = server.CreateObject("ADODB.RecordSet")%>
	
	<html>
		<head>
		<title>Export dati in formato HTML</title>
		<meta name="robots" content="noindex,nofollow" />
		<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
		</head>
		<body>
		<table border="1" bordercolor="#000000" cellspacing="0">
			<tr>
				<%for each Field in rs.Fields%>
					<th><%= Field.name %>&nbsp;</th>
				<%next 
				while not rsv.eof%>
					<th><%= rsv("nome_tipoNumero") %>&nbsp;</th>
					<%rsv.moveNext
				wend%>
				<th>Rubriche</th>
			</tr>
			<% while not rs.eof %>
				<tr>
					<%for each Field in rs.Fields%>
					<td><%= Field.value %>&nbsp;</td>
				<%next
				rsv.moveFirst
				while not rsv.eof
					sql = "SELECT ValoreNumero FROM tb_ValoriNumeri " &_
				 		  " WHERE id_TipoNumero=" & rsv("id_tipoNumero") & " AND  id_Indirizzario=" & rs("ID")%>
					<td><%= GetValueList(conn, rsN, sql) %>&nbsp;</td>
					<%rsv.moveNext
				wend
				sql = " SELECT nome_rubrica FROM tb_rubriche " &_
				  	  " INNER JOIN rel_rub_ind ON tb_rubriche.id_rubrica=rel_rub_ind.id_rubrica " &_
					  " WHERE rel_rub_ind.id_indirizzo=" & rs("ID") %>
				<td><%= cString(GetValueList(conn, rsN, sql)) %>&nbsp;</td>
				</tr>
				<% rs.MoveNext
			wend %>
		</table>
		</body>	
	</html>

	<%set rsN = nothing
end sub



sub Export_TXT(conn, rs, sql, rsv)
	dim Field, rsN, values
	set rsN = server.CreateObject("ADODB.RecordSet")
	response.ContentType = "text/plain"
	
	'scrive testata
	for each Field in rs.Fields
		response.write Field.name & ";"
	next 
	while not rsv.eof
		response.write rsv("nome_tipoNumero") & ";"
		rsv.moveNext
	wend
	response.write "Rubriche;"
	response.write vbCrLf
	
	'scrive righe
	while not rs.eof
		values = ""
		rsv.moveFirst
		while not rsv.eof
			sql = "SELECT ValoreNumero FROM tb_ValoriNumeri " &_
				  " WHERE id_TipoNumero=" & rsv("id_tipoNumero") & " AND  id_Indirizzario=" & rs("ID")
			values = values & GetValueList(conn, rsN, sql) & ";"
			sql = " SELECT nome_rubrica FROM tb_rubriche " &_
			  	  " INNER JOIN rel_rub_ind ON tb_rubriche.id_rubrica=rel_rub_ind.id_rubrica " &_
				  " WHERE rel_rub_ind.id_indirizzo=" & rs("ID") 
			values = values & GetValueList(conn, rsN, sql) & ";"
			rsv.moveNext
		wend
		response.write rs.GetString(adClipString, 1, ";", "", "")
		response.write values & vbCrLf
	wend
	
	set rsN = nothing
end sub
%>
