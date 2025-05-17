<!--#INCLUDE FILE="Tools.asp" -->
<!--#INCLUDE FILE="Tools4Admin.asp" -->
<!--#INCLUDE FILE="ClassCryptography.asp"-->
<%


'***************************************************************************************************************************
'***************************************************************************************************************************
'COSTANTI E DICHIARAZIONI
'***************************************************************************************************************************
'***************************************************************************************************************************

'impostazione variabili per gestione del cookie
private CookieName
private CookieApplication
CALL InitializeCookie(CookieName, CookieApplication)

'dichiarazione oggetto per parametri dell'applicazione
dim Parametri
set Parametri = new ApplicationParameters


'***************************************************************************************************************************
'***************************************************************************************************************************
'FUNZIONI DI GESTIONE DEL COOKIE
'***************************************************************************************************************************
'***************************************************************************************************************************

'Inizializza le variabili di sessione per il login
Sub InitSex(byval id)
	CALL PreserveInitSex(id, false)
	
	'imposta la lingua
	if id = 18 then
		session("LINGUA") = session("LINGUA")
	else
		'session("LINGUA") = LINGUA_ITALIANO
		session("LINGUA") = session("LINGUA")
	end if
End Sub


'***************************************************************************************************************************
'***************************************************************************************************************************
'FUNZIONI DI GESTIONE DEL COOKIE
'***************************************************************************************************************************
'***************************************************************************************************************************


'...............................................................................
'.. classe per la gestione ed il controllo dei parametri delle applicazioni
'...............................................................................
Class ApplicationParameters
	
	Private parameters
	Public Declared

	Private Sub Class_Initialize()
		Declared = false
		set parameters = Server.CreateObject("Scripting.Dictionary")
		parameters.CompareMode = vbTextCompare
	end sub
	
	Private Sub Class_Terminate()
		set parameters = nothing
	End Sub
	
	Public Sub Clear()
		Declared = false
		set parameters = Server.CreateObject("Scripting.Dictionary")
		parameters.CompareMode = vbTextCompare
	end sub
	
	
	'aggiunge parametro con relativo valore di default
	Public sub Add(name, DefaultValue)
		Declared = true
		if parameters.Exists(cString(name)) then
			parameters(cString(name)) = DefaultValue
		else
			parameters.Add cString(name), DefaultValue
		end if
	end sub
	
	
	'restituisce true se sono attivi i nuovi parametri per l'applicazione corrente.
	Public function IsParamsNewVersion(conn, rs, SITO_ID)
		dim sql
		sql = "SELECT COUNT(*) FROM rel_siti_descrittori WHERE rsd_sito_id = "& SITO_ID
		if cIntero(GetValueList(conn, rs, sql))>0 then
			'nuovi parametri in uso
			IsParamsNewVersion = true
		else
			'vecchi parametri in uso
			IsParamsNewVersion = false
		end if
	end function
	
	
	'funzione che verifica i dati dei parametri: 
	'	se non ci sono li inserisce con il valore di default
	' 	se ce ne sono in piu' li cancella
	Public Sub Check(conn, rs, SITO_ID)
		dim sql
		'verifica se tutti i parametri impostati per l'applicaizione sono gia' presenti
		dim parameter
		for each parameter in parameters.keys
			'verifica se ogni parametro e' gia' presente
			sql = " SELECT * FROM tb_siti_parametri WHERE par_sito_id=" & SITO_ID & _
				  " AND par_key LIKE '" & parameter & "'"
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
			if rs.eof then
				rs.AddNew
				rs("par_sito_id") = SITO_ID
				rs("par_key") = parameter
				rs("par_value") = parameters(parameter)
				rs.Update
			end if
			rs.close
		next
		
		'cancella i parametri in piu'
		if not IsParamsNewVersion(conn, rs, SITO_ID) then
			'solo se per l'applicativo non e' ancora stata attivata la nuova gestione dei parametri, altrimenti lascia le vecchie copie attive.
			sql = "SELECT * FROM tb_siti_parametri WHERE par_sito_id=" & SITO_ID
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
			if rs.recordcount > parameters.Count then
				'il numero dei parametri presenti e' superiore ai parametri necessari: cancella quelli in pi&ugrave;
				while not rs.eof
					if not parameters.exists(cString(rs("par_key"))) then
						'il parametro non e' utilizzato: lo cancella
						sql = "DELETE FROM tb_siti_parametri WHERE par_id=" & rs("par_id")
						CALL conn.execute(sql, , adExecuteNoRecords)
					end if
					rs.movenext
				wend
			end if
			rs.close
		end if
	end sub
	
	'cairca i parametri del sito, indipendentemente dalla loro versione (nuovi o vecchi)
	Public sub LoadAllParams(conn, rs, idSito)
		dim sql, rsCreated, connCreated
		if not IsObjectCreated(conn) then
			connCreated = true
			set conn = server.createobject("adodb.connection")
			conn.open Application("DATA_ConnectionString"),"",""
		else
			connCreated = false
		end if
		if not IsObjectCreated(rs) then
			rsCreated = true
			set rs = server.createobject("adodb.recordset")
		else
			rsCreated = false
		end if
		
		'imposta valori parametri dell'applicazione
		sql = "SELECT COUNT(*) FROM rel_siti_descrittori WHERE rsd_sito_id = "& idSito
		if CIntero(GetValueList(conn, rs, sql)) > 0 then						'se abilitata la nuova gestione dei parametri
			CALL LoadNew(conn, rs, idSito)
		else
			CALL Load(conn, rs, idSito)
		end if
			
		if rsCreated then
			set rs = nothing
		end if
		if connCreated then
			conn.close
			set conn = nothing
		end if
	end sub
	
	'carica i parametri del sito in sessione, se rs = null lo crea
	Public Sub Load(conn, rs, idSito)
		dim sql, rsNew
		if IsNull(rs) then
			set rs = server.createobject("adodb.recordset")
			rsNew = true
		else
			rsNew = false
		end if
		
		sql = "SELECT par_key, par_value FROM tb_siti_parametri WHERE par_sito_ID="& idSito
		rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		while not rs.eof
			Session(rs("par_key")) = rs("par_value")
			'response.write rs("par_key") & " = " & Session(rs("par_key"))
			rs.movenext
		wend
		rs.close
		'response.end
		if rsNew then
			set rs = nothing
		end if
	End Sub
	
	'carica i NUOVI parametri del sito in sessione, se rs = null lo crea
	Public Sub LoadNew(conn, rs, idSito)
		dim sql, rsNew
		if IsNull(rs) then
			set rs = server.createobject("adodb.recordset")
			rsNew = true
		else
			rsNew = false
		end if
		
		sql = " SELECT * FROM tb_siti_descrittori d"& _
			  " INNER JOIN rel_siti_descrittori r ON d.sid_id = r.rsd_descrittore_id"& _
			  " WHERE rsd_sito_id = "& idSito			  
		rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		while not rs.eof
			Session(rs("sid_codice")) = DesValue(rs, "sid_tipo", "rsd_", LINGUA_ITALIANO)
			'response.write rs("sid_codice") & " = " & DesValue(rs, "sid_tipo", "rsd_", LINGUA_ITALIANO)
			rs.movenext
		wend
		rs.close
		'response.end
		if rsNew then
			set rs = nothing
		end if
	End Sub
	
end class
%>
