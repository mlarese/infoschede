<%
'*******************************************************************************************************************
'*******************************************************************************************************************
'DEFINIZIONE CLASSE
'*******************************************************************************************************************
class Import_OBJ

	'variabili utilizzate nelle proprieta'
	Public Name_Table			'Nome della tabella
	Public Name_Import			'Nome dell'import


	'variabili interne
	Private rs, sql
	Private CloseConnection
	
	
	Private Sub Class_Initialize()
		set rs = Server.CreateObject("ADODB.RecordSet")
	end sub

	private sub Class_Terminate()
	
	end sub

	
	'************************************************************************************************************
	'FUNZIONI DI INTERFACCIA PUBBLICA
	'******************************************************
	
	Public Sub WriteLogBeginImport(Conn)
		CALL WriteLogAdmin(Conn, Name_Table, 0, Name_Import, "Inizio import")
	End Sub
	
	Public Sub WriteLogEndImport(Conn)
		CALL WriteLogAdmin(Conn, Name_Table, 0, Name_Import, "Fine import")
	End Sub
	
	Public Function GetLastDateImport(Conn)
		dim lastDateImport
		sql = "SELECT TOP 1 log_data FROM log_framework WHERE log_table_nome LIKE '"&Name_Table&"' AND log_descrizione LIKE 'Fine import' ORDER BY log_id DESC"
		rs.open sql, Conn, adOpenStatic, adLockOptimistic, adCmdText
		if not rs.eof then
			lastDateImport = rs("log_data")
		else
			'lastDateImport = DateSerial(2000, 1, 1)
			lastDateImport = NULL
		end if
		rs.close
		GetLastDateImport = lastDateImport
	End Function
	

	'************************************************************************************************************
	'FUNZIONI PRIVATE'
	'******************************************************


end class

%>