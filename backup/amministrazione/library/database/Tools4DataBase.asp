<!--#INCLUDE FILE="Update__library__b2b.asp" -->
<!--#INCLUDE FILE="Update__library__b2b_tour.asp" -->
<!--#INCLUDE FILE="Update__library__Info.asp" -->
<!--#INCLUDE FILE="Update__library__banner.asp" -->
<!--#INCLUDE FILE="Update__library__banner2.asp" -->
<!--#INCLUDE FILE="Update__library__memo.asp" -->
<!--#INCLUDE FILE="Update__library__memo2.asp" -->
<!--#INCLUDE FILE="Update__library__booking.asp" -->
<!--#INCLUDE FILE="Update__library__booking2.asp" -->
<!--#INCLUDE FILE="Update__library__booking3.asp" -->
<!--#INCLUDE FILE="Update__library__flat.asp" -->
<!--#INCLUDE FILE="Update__library__realestate.asp" -->
<!--#INCLUDE FILE="Update__library__travel2.asp" -->
<!--#INCLUDE FILE="Update__library__framework_core.asp" -->
<!--#INCLUDE FILE="Update__library__comment.asp" -->
<!--#INCLUDE FILE="Update__library__guestbook.asp" -->
<!--#INCLUDE FILE="Update__library__questionario.asp" -->
<!--#INCLUDE FILE="Update__library__infoschede.asp" -->
<!--#INCLUDE FILE="Update__library__commesse.asp" -->
<!--#INCLUDE FILE="Install__library__paypal.asp" -->
<!--#INCLUDE FILE="Install__library__virtualpay.asp" -->
<%

'procedura che verifica se l'utente corrente e' abilitato, altrimenti riporta al login
sub VerificaPermessiUtente(IsNewWindow)
	if cString(Session("UTENTE_MANUTENZIONE"))="" then
		if IsNewWindow then %>
			<script language="JavaScript" type="text/javascript">
				//esegue reload della finestra padre 
				try { opener.location.reload(true);}
				catch(e){/*istruzione messa solo per sintassi*/}
		
				//chiude la finestra corrente
				window.close();
			</script>
			<%response.end
		else
			response.redirect "default.asp"
		end if
	end if
end sub


'ritorna il numero di versione del database collegato alla connessione
function ReadCurrentDbVersion(conn)
	if VersionTableExists(conn) then
		dim sql, rs
		set rs = server.CreateObject("ADODB.RecordSet")
		sql = "SELECT * FROM AA_Versione"
		rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		ReadCurrentDbVersion = rs("Versione")
		rs.close
		set rs = nothing
	else
		ReadCurrentDbVersion = "tabella mancante"
	end if
end function


'verifica se esiste la tabella di controllo della versione sul database corrente
function VersionTableExists(conn)
    VersionTableExists = TableExists(conn, "AA_Versione")
end function


'compatta il database della connessione passata
sub CompactDatabase(conn)
	dim BasePath, dbToCompact, dbCompacted, ConnStringToCompact, ConnStringCompacted
	dim Engine, FSO
	
	if DB_Type(conn) = DB_Access then
		
		'calcola e genera i percorsi di base delle connessioni
		BasePath = left(conn.Properties("Data Source"), instrRev(conn.Properties("Data Source"), "\")) 		'"
		
		dbToCompact = conn.Properties("Data Source")
		dbCompacted = BasePath & Session.SessionID & ".mdb"
		
		ConnStringToCompact = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & dbToCompact
		ConnStringCompacted = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & dbCompacted
		
		conn.close
		
		'compatta il database creandone uno di nuovo
		Set Engine = CreateObject("JRO.JetEngine")
	 	Engine.CompactDatabase ConnStringToCompact, ConnStringCompacted
		
		'esegue operazioni sui file di database
		set FSO = CreateObject("Scripting.FileSystemObject")
		
		'cancella il dababase di origine da compattare
		FSO.DeleteFile(dbToCompact)
		
		'crea una copia del file compattato con il nome del database di origine da compattare
		FSO.CopyFile dbCompacted, dbToCompact, true
		
		'cancella il dababase temporaneo compattato
		FSO.DeleteFile(dbCompacted)
			
		conn.open ConnStringToCompact, "", ""
	else
		dim sql, DbmsVersion
		conn.CommandTimeout = 360
		DbmsVersion = DB_SQL_version(conn)
		if DbmsVersion = DB_SQL_2000 OR _
		   DbmsVersion = DB_SQL_2005 then
			'parte di script per database sql server 2000 e 2005
			
			'esegue backup per troncare i log
			sql = "BACKUP LOG [" + GetDatabaseName(conn) + "] WITH NO_LOG "
			if DB_Type(conn) = DB_SQL then 
				if cIntero(GetValueList(conn, NULL, "SELECT REPLACE(SUBSTRING(CAST(SERVERPROPERTY('productversion') as varchar(100)),1,2),'.','')"))>9 then
					'versione Microsoft SQL 2008 o successive
					sql = "BACKUP LOG [" + GetDatabaseName(conn) + "] WITH INIT, COMPRESSION "
				end if
			end if
			CALL conn.execute(sql)
			
			'compatta databae
			sql = "DBCC SHRINKDATABASE ([" + GetDatabaseName(conn) + "] )"
			CALL conn.execute(sql)
		else
			'script eseguito per 2008
			dim rsf, sqlCmd
			set rsf = server.CreateObject("ADODB.RecordSet")
			
			'recupera elenco dei file che compongono il database (generalmente 2: log e dati)
			'sql = "SELECT * FROM sys.database_files WHERE type_desc LIKE 'ROWS' OR type_desc LIKE 'LOG'"
			
			'Modificato da Nicola il 02/05/2012: andiamo a compattare solo il file di log
			sql = "SELECT * FROM sys.database_files WHERE type_desc LIKE 'LOG'"
			rsf.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			
			while not rsf.eof
				'per ogni file esegue la routine di compattazione
				sqlCmd = sqlCmd + _
						 " USE master " + vbcrLF + _
						 " ALTER DATABASE [" + GetDatabaseName(conn) + "] SET Recovery Simple WITH NO_WAIT" + vbCrlf + _
						 " USE [" + GetDatabaseName(conn) + "] " + vbCrLf + _
						 " DBCC SHRINKFILE (N'" + rsf("name") + "' , 1) " + vbCrLf + _
						 " USE [master] " + vbCrLf + _
						 " ALTER DATABASE [" + GetDatabaseName(conn) + "] SET Recovery Full WITH NO_WAIT" + vbCrLF + _
						 " USE [" + GetDatabaseName(conn) + "] " + vbCrLf + _
						 " ; " + vbCrLF
				rsf.movenext
			wend
			
			rsf.close
			set rsf = nothing
		
			'esegue la compattazione
			CALL ExecuteMultipleSql(conn, sqlCmd, true)
			
		end if
	end if
end sub


'restituisce la dimensione del file del database
function DatabaseSize(conn)
	if DB_Type(conn) = DB_Access then
		dim fso, f
		set FSO = CreateObject("Scripting.FileSystemObject")
		set f = FSO.GetFile(Conn.Properties("Data Source"))
        'DatabaseSize = FormatPrice((f.size / 1024), 0, true) & " kb"
        DatabaseSize = File_Dimension( f.size )
		set f = nothing
		set FSO = nothing
	else
		dim rs, size
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open "sp_helpfile", conn, adOpenStatic, adLockOptimistic, adCmdStoredProc
		size = 0
		while not rs.eof
			size = size + cInteger(left(rs("size"), instr(1, rs("size"), " ", vbTextCompare)))
			rs.movenext
		wend
        DatabaseSize = File_Dimension( size * 1024 )
		rs.close
		set rs = nothing
	end if
end function


function GetDatabaseName(conn)
	if conn.properties("Current Catalog")<>"" then
		'database sql
		GetDatabaseName = conn.properties("Current Catalog")
	else
		'altro database (access)
		GetDatabaseName = conn.properties("Data Source")
		GetDatabaseName = right(GetDatabaseName, (len(GetDatabaseName) - instrRev(GetDatabaseName, "\")))   '"
		GetDatabaseName = left(GetDatabaseName, instr(1, GetDatabaseName, ".", vbtextCompare)-1)
	end if
end function


'funzione che verifica se esiste o meno un oggetto (VALIDO SOLO PER LE TABELLE )
function TableExists(conn, objName)
    if cString(objName)<>"" then
        dim rs
        set rs = conn.OpenSchema(adSchemaTables, Array(Empty, Empty, objName))
        TableExists = not rs.eof
	    set rs = nothing
    else
        TableExists = false
    end if
end function


'funzione che verifica se esiste o meno un oggetto (VALIDO SOLO PER LE VISTE )
function ViewExists(conn, objName)
	if cString(objName)<>"" then
        dim rs
		if DB_Type(conn) = DB_SQL then
			dim sql 
			sql = "SELECT * FROM sysobjects WHERE id = object_id('" + ParseSql(objName, adChar) + "')"
			set rs = Server.CreateObject("ADODB.Recordset")
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			
			if rs.eof AND DB_SQL_version(conn) >= DB_SQL_2005 then
				rs.close
				sql = "SELECT * FROM sys.indexes WHERE name LIKE '" + ParseSql(objName, adChar) + "'"
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			end if
		else
			set rs = conn.OpenSchema(adSchemaViews, Array(Empty, Empty, objName))
		end if
        ViewExists = not rs.eof
		rs.close
	    set rs = nothing
    else
        ViewExists = false
    end if
end function


'restituisce il codice di reset dell'identity
function SQLSERVER_ReseedIdentity(conn, tableName, identityField)
	SQLSERVER_ReseedIdentity = _
		" DECLARE @maxid int " + vbCrLf + _
		" SELECT @maxid = max(" + identityField + ") FROM " + tableName + vbCrLF + _
		" DBCC CHECKIDENT ( " + tableName + ", RESEED, @maxid ) ; "
end function


'produce codice per la cancellazione di un oggetto, verificandone l'esistenza prima della cancellazione
function DropObject(conn, objName, objType)
    Select case uCase(objType)
        case "TABLE"
            if TableExists(conn, objName) then
                DropObject = "DROP " & objType & " " & objName
            else
                DropObject = ""
            end if
        case else
			if ViewExists(conn, objName) then
				if DB_Type(conn) = DB_SQL then
					if uCase(objType)<>"INDEX" then
						DropObject = "IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('" + ParseSql(objName, adChar) + "') "
					else
						DropObject = "IF EXISTS (SELECT * FROM sys.indexes WHERE name LIKE '" + ParseSql(objName, adChar) + "' "
					end if
				end if
				Select Case uCase(objType)
					case "VIEW"
						DropObject = DropObject + _
									 IIF(DB_Type(conn) = DB_SQL, " AND sysstat & 0xf = 2) " + vbCrLF, "") + _
									 " DROP VIEW " + objName
					case "PROCEDURE"
						DropObject = DropObject + _
									 IIF(DB_Type(conn) = DB_SQL, " AND sysstat & 0xf = 4) " + vbCrLF, "") + _
									 " DROP PROCEDURE " + objName
					case "INDEX"
						dim sql 
						sql = "SELECT sysobjects.name FROM sysobjects " + _
							  " INNER JOIN sys.indexes ON sysobjects.id = sys.indexes.object_id " +  _
							  " WHERE sys.indexes.name = '" + ParseSql(objName, adChar) + "' "
						DropObject = DropObject + _
									 IIF(DB_Type(conn) = DB_SQL, ") " + vbCrLF, "") + _
									 " DROP INDEX " + objName + " ON " + GetValueList(conn, NULL, sql)
					case "TRIGGER"
						DropObject = DropObject + _
									 IIF(DB_Type(conn) = DB_SQL, " AND sysstat & 0xf = 8) " + vbCrLF, "") + _
									 " DROP TRIGGER " + objName
					case "FUNCTION"
						DropObject = DropObject + _
									 IIF(DB_Type(conn) = DB_SQL, " and xtype in (N'FN', N'IF', N'TF')) " + vbCrLF, "") + _
									 " DROP FUNCTION " + objName
					case else
						DropObject = ""
				end select
			end if
    end select
    
    DropObject = DropObject + ";" + vbCrLf
end function


'funzione che restituisce il codice sql per definire un campo testuale in access o in SQL server.
function SQL_CharField(Conn, lenght)
	if DB_Type(conn) = DB_SQL then
		if cIntero(lenght)> 0 then
		  	SQL_CharField = " nvarchar(" & lenght & ") "
		else
			SQL_CharField = " ntext "
		end if
	else
	  	if cIntero(lenght)> 0 then
		  	SQL_CharField = " TEXT(" & IIF(lenght>255, 255, lenght) & ") WITH COMPRESSION "
		else
			SQL_CharField = " TEXT WITH COMPRESSION "
		end if
	end if
end function


'funzione che restituisce dbo. per gli oggetti sql server
function SQL_Dbo(Conn)
	if DB_Type(conn) = DB_SQL then
		SQL_Dbo = " dbo."
	else 
		SQL_Dbo = ""
	end if
end function 


'Restituisce il codice SQL per la creazione di un campo chiave primaria e del relativo vincolo
Function SQL_PrimaryKey(conn, tabellaNome)
	select case DB_Type(conn)
		case DB_Access
			SQL_PrimaryKey = "COUNTER CONSTRAINT PK_" + tabellaNome + " PRIMARY KEY"
		case DB_SQL
			SQL_PrimaryKey = "INT IDENTITY(1, 1) NOT NULL CONSTRAINT PK_" + tabellaNome + " PRIMARY KEY CLUSTERED "
	end select
End Function


'Restituisce il codice SQL per la creazione di un campo chiave primaria non contatore ma intero
Function SQL_PrimaryKeyInt(conn, tabellaNome)
	select case DB_Type(conn)
		case DB_Access
			SQL_PrimaryKeyInt = "INT NOT NULL CONSTRAINT PK_" + tabellaNome + " PRIMARY KEY"
		case DB_SQL
			SQL_PrimaryKeyInt = "INT NOT NULL CONSTRAINT PK_" + tabellaNome + " PRIMARY KEY CLUSTERED "
	end select
End Function


Function SQL_AddColumn(conn)
	select case DB_Type(conn)
		case DB_Access
			SQL_AddColumn = "ADD COLUMN"
		case DB_SQL
			SQL_AddColumn = "ADD"
	end select
End Function


function SQL_MultiLanguageField(fieldDefinition)
    SQL_MultiLanguageField = replace(fieldDefinition, "<lingua>", "it") + ", " + vbCrLF + _
                             replace(fieldDefinition, "<lingua>", "en") + ", " + vbCrLF + _
                             replace(fieldDefinition, "<lingua>", "fr") + ", " + vbCrLF + _
                             replace(fieldDefinition, "<lingua>", "de") + ", " + vbCrLF + _
                             replace(fieldDefinition, "<lingua>", "es") + vbCrLF
end function


function SQL_MultiLanguageFieldNew(fieldDefinition)
    SQL_MultiLanguageFieldNew = replace(fieldDefinition, "<lingua>", "ru") + ", " + vbCrLF + _
								replace(fieldDefinition, "<lingua>", "cn") + ", " + vbCrLF + _
								replace(fieldDefinition, "<lingua>", "pt") + vbCrLF
end function


function SQL_MultiLanguageFieldComplete(conn, fieldDefinition)
	Dim sql, rs
	sql = SQL_MultiLanguageField(fieldDefinition)
	
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open "SELECT * FROM tb_cnt_lingue", conn, adOpenKeySet, adLockOptimistic, adCmdText
	while not rs.eof 
		if rs("lingua_codice")<>"it" AND rs("lingua_codice")<>"en" AND rs("lingua_codice")<>"fr" AND rs("lingua_codice")<>"de" AND rs("lingua_codice")<>"es" then
			sql = sql & ", " &  replace(fieldDefinition, "<lingua>", rs("lingua_codice"))
		end if
		rs.moveNext
	wend
	rs.close
	
    SQL_MultiLanguageFieldComplete = sql
	
	set rs = nothing
end function


'attiva il service broker per la cache di ASP.net
function SQL_Activate_Service_Broker(conn)
	SQL_Activate_Service_Broker = "ALTER DATABASE " + GetDatabaseName(conn) + " SET ENABLE_BROKER ; "
end function



'aggiunge relazione tra tabelle
function SQL_AddForeignKey(conn, Table, FKField, ReferencedTable, ReferencedField, integrity, ConstraintNameNote)
	SQL_AddForeignKey = SQL_AddForeignKeyExtended(conn, Table, FKField, ReferencedTable, ReferencedField, integrity, true, ConstraintNameNote)
end function 


function SQL_AddForeignKeyExtended(conn, Table, FKField, ReferencedTable, ReferencedField, integrity, cascade, ConstraintNameNote)
	dim sql, ConstraintName
	if integrity OR DB_Type(conn) = DB_SQL then
		if uCase(Table) <> uCase(ReferencedTable) then
			ConstraintName = "FK_" + Table + "__" + ReferencedTable
		else
			ConstraintName = "FK_" + Table + "__" + FKField
		end if
		ConstraintName = ConstraintName + IIF(ConstraintNameNote<>"", "_" + ConstraintNameNote, "")
		
		sql = " ALTER TABLE " + SQL_dbo(conn) + Table + _
			  IIF(not integrity AND DB_Type(conn) = DB_SQL, " WITH NOCHECK ", "") + _
			  " ADD CONSTRAINT " + ConstraintName + _
			  "	FOREIGN KEY (" + FKField + ") " + _
			  " REFERENCES " + ReferencedTable + "(" + ReferencedField + ") "
		
		if integrity then
			if cascade then
				sql = sql + " ON UPDATE CASCADE ON DELETE CASCADE "
			else
				sql = sql + " ON UPDATE NO ACTION ON DELETE NO ACTION "
			end if
		elseif DB_Type(conn) = DB_SQL then
			sql = Sql + "; " + _
					    " ALTER TABLE " + SQL_dbo(conn) + Table + " NOCHECK CONSTRAINT " + ConstraintName
		end if
		
		SQL_AddForeignKeyExtended = sql + ";"
	else
		SQL_AddForeignKeyExtended = ""
	end if
end function 


'rimuove la relazione tra tabelle
function SQL_RemoveForeignKey(conn, Table, FKField, ReferencedTable, integrity, ConstraintName)
    dim sql
	if integrity OR DB_Type(conn) = DB_SQL then
        if ConstraintName = "" then
            if uCase(Table) <> uCase(ReferencedTable) then
			    ConstraintName = "FK_" + Table + "__" + ReferencedTable
    		else
	    		ConstraintName = "FK_" + Table + "__" + FKField
		    end if
        end if
        sql = " ALTER TABLE " + SQL_dbo(conn) + Table + _
		    	  " DROP CONSTRAINT " + ConstraintName + "; "
        SQL_RemoveForeignKey = sql
    else
        SQL_RemoveForeignKey = ""
    end if
end function


'prefix: 		prefisso della tabella senza "_" finale
Function AddInsModFields(prefix)
	AddInsModFields = _
		"	"& prefix &"_insData DATETIME NULL, " + vbCrLF + _
		"	"& prefix &"_insAdmin_id INT NULL, " + vbCrLF + _
		"	"& prefix &"_modData DATETIME NULL, " + vbCrLF + _
		"	"& prefix &"_modAdmin_id INT NULL"
End Function


'tab:			nome della tabella
Function AddInsModRelations(conn, tab, prefix)
	if DB_Type(conn) = DB_SQL then
		AddInsModRelations = _
			SQL_AddForeignKey(conn, tab, prefix + "_insAdmin_id", "tb_admin", "id_admin", false, "ins") + _
			SQL_AddForeignKey(conn, tab, prefix + "_modAdmin_id", "tb_admin", "id_admin", false, "mod")
	end if
End Function



sub AddParametroSito(conn, codice, raggruppamento_id, nome, unita, tipo, principale, immagine, admin, personalizzato, sito_id, valore_it, valore_en, valore_fr, valore_de, valore_es)
	CALL AddParametroSitoNew(conn, codice, raggruppamento_id, nome, unita, tipo, principale, immagine, admin, personalizzato, sito_id, valore_it, valore_en, valore_fr, valore_de, valore_es, "", "")
end sub

sub AddParametroSitoNew(conn, codice, raggruppamento_id, nome, unita, tipo, principale, immagine, admin, personalizzato, sito_id, valore_it, valore_en, valore_fr, valore_de, valore_es, valore_ru, valore_cn)
	CALL AddParametroSitoNew2(conn, codice, raggruppamento_id, nome, unita, tipo, principale, immagine, admin, personalizzato, sito_id, valore_it, valore_en, valore_fr, valore_de, valore_es, valore_ru, valore_cn, "")
end sub


'funzione che aggiunge un parametro alla tabella di parametri delle applicazioni
sub AddParametroSitoNew2(conn, codice, raggruppamento_id, nome, unita, tipo, principale, immagine, admin, personalizzato, sito_id, valore_it, valore_en, valore_fr, valore_de, valore_es, valore_ru, valore_cn, valore_pt)
	dim sql, rs, sid_id
	set rs = Server.CreateObject("ADODB.Recordset")
	
	'salva descrittore
	sql = "SELECT * FROM tb_siti_descrittori WHERE sid_codice LIKE '" & ParseSQL(codice, adChar) & "'"
	rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
	if rs.eof then
		rs.AddNew
	end if
	rs("sid_codice") = codice
	rs("sid_raggruppamento_id") = raggruppamento_id
	rs("sid_nome_it") = nome
	rs("sid_tipo") = tipo
	rs("sid_unita_it") = unita
	rs("sid_principale") = principale
	rs("sid_img") = immagine
	rs("sid_admin") = admin
	rs("sid_personalizzato") = personalizzato
	rs.update
	
	sid_id = rs("sid_id")
	rs.close
	
	'salva valore del descrittore
	sql = "SELECT * FROM rel_siti_descrittori WHERE rsd_sito_id=" & sito_id & " AND rsd_descrittore_id=" & sid_id
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	if rs.eof then
		rs.AddNew
		rs("rsd_sito_id") = sito_id
		rs("rsd_descrittore_id") = sid_id
	end if
	if tipo = adLongVarChar then
		'salva nei campi memo
		rs("rsd_memo_it") = cString(valore_it)
		rs("rsd_memo_en") = cString(valore_en)
		rs("rsd_memo_fr") = cString(valore_fr)
		rs("rsd_memo_de") = cString(valore_de)
		rs("rsd_memo_es") = cString(valore_es)
		if FieldExists(rs, "rsd_memo_ru") then
			rs("rsd_memo_ru") = cString(valore_ru)
		end if
		if FieldExists(rs, "rsd_memo_cn") then
			rs("rsd_memo_cn") = cString(valore_cn)
		end if
		if FieldExists(rs, "rsd_memo_pt") then
			rs("rsd_memo_pt") = cString(valore_pt)
		end if
	else
		'salva campi testo
		rs("rsd_valore_it") = cString(valore_it)
		rs("rsd_valore_en") = cString(valore_en)
		rs("rsd_valore_fr") = cString(valore_fr)
		rs("rsd_valore_de") = cString(valore_de)
		rs("rsd_valore_es") = cString(valore_es)
		if FieldExists(rs, "rsd_valore_ru") then
			rs("rsd_valore_ru") = cString(valore_ru)
		end if
		if FieldExists(rs, "rsd_valore_cn") then
			rs("rsd_valore_cn") = cString(valore_cn)
		end if
		if FieldExists(rs, "rsd_valore_pt") then
			rs("rsd_valore_pt") = cString(valore_pt)
		end if
	end if
	rs.update
	rs.close
end sub


'esegue la copia dei dati della tabella SourceTable del database SourceConn nel database DestConn.
'Copia tutte le righe, aggiornando quelle esistenti, ed inserendo anche gli id
'ATTENZIONE: le due tabelle DEVONO essere strutturalmente identiche
sub CopyTableData(SourceConn, DestConn, SourceTable, DestTable, PrimaryKeyField, IdentityInsertForced)
    dim sql, rsS, rsD, field, IsInserting
    set rsS = Server.CreateObject("ADODB.Recordset")
    set rsD = Server.CreateObject("ADODB.Recordset")
    
    if DestTable = "" then
        DestTable = SourceTable
    end if
    
    if IdentityInsertForced then
        sql = " SET IDENTITY_INSERT " + DestTable + " ON "
        CALL DestConn.execute(sql, ,adCmdText)
    end if
    
    sql = "SELECT * FROM " + SourceTable
    %><!-- 
    Copia della tabella: <%= DestTable %>
    Sorgente: <%= sql %>
    --><%
    rsS.open sql, SourceConn, adOpenStatic, adLockOptimistic, adCmdText
    while not rsS.eof
        sql = "SELECT * FROM " + DestTable + " WHERE " + PrimaryKeyField + "=" & cIntero(rsS(PrimaryKeyField))
        rsD.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
        if rsD.eof then
            rsD.AddNew
            IsInserting = true
        else
            IsInserting = false
        end if
        
        for each field in rsS.Fields
            if IsInserting OR _
               lCase(field.name) <> lCase(PrimaryKeyField) then
                rsD(field.name) = rsS(field.name)
            end if
        next
        
        rsD.Update
        rsD.close
        rsS.MoveNext
    wend
    rsS.Close
    
    if IdentityInsertForced then
        sql = " SET IDENTITY_INSERT " + DestTable + " OFF "
        CALL DestConn.execute(sql, ,adCmdText)
    end if
    
    set rss = nothing
    set rsd = nothing
end sub


'*****************************************************************************************************
class UpdateDabase

	Public objConn
	Public logConn
	Public last_update_executed
    Public count_update_executed
    Public Terminate_SQL
	Private TimerStart
	Private TimerPrevious
    
	Private Sub Class_Initialize()
        Terminate_SQL = ""
        count_update_executed = 0
        last_update_executed = false
		TimerStart = Timer()
		TimerPrevious = TimerStart
		
    end sub
    
	Private Sub Class_Terminate()
         dim fso
        'esegue aggiornamento di chiusura per ricostruzione viste e procedure
        CALL Rebuild_OnTerminate(Terminate_SQL) %>
        <tr><th class="l2_center" colspan="5">Ripulitura directory temporanee e file non validi</th></tr>
        <%
        set fso = Server.CreateObject("Scripting.FileSystemObject")
        'ripulisce directory temporanee
    	CALL ClearTempDir(fso)
        
    	'rimuove file inutili dalle directory (qualsiasi directory)
    	CALL FileRemove(fso, Application("IMAGE_PATH"), "thumbs.db", true)
    	CALL FileRemove(fso, Application("IMAGE_PATH"), "pspbrwse.jbf", true)
        
        set fso = nothing
        %>
        <tr>
            <td class="content_center">ripulitura files</td>
            <td class="content">Ripulitura eseguita CORRETTAMENTE</td>
            <td class="content_center">&nbsp;</td>
			<% CALL TimeTracer() %>
        </tr>
        <% if Session("ERRORE")<>"" then %>
			<tr>
				<td colspan="5" class="errore"><%= Session("ERRORE") %></td>
			</tr>
			<% Session("ERRORE") = ""
		end if %>
			<tr>
				<td class="label_no_width" colspan="3">
					fine aggiornamento: <%= DateTimeIta(Now()) %>:<%= FixLenght(Second(Now()), "0", 2) %>
				</td>
				<td class="content_right" colspan="2">
					durata totale: <strong><%= FormatPrice(((Timer() - TimerStart)), 2, true) %></strong> secondi
				</td>
			</tr>
		</table>
		<script language="JavaScript" type="text/javascript">
			try{
				opener.document.location.reload(true);
			} catch(e){
			}
		</script>
		<%
		
		CALL WriteLogAdminHttp(logConn, "AA_versione", 0, "AGGIORNAMENTO_DATABASE_COMPLETATO", "Aggiornamenti database completato", true)
		
	End Sub


	'crea la tabella di controllo della versione e imposta la versione a 0
	Private Sub CreateVersionTable()
		
		sql = "CREATE TABLE dbo.AA_Versione (versione int)"
		CALL objConn.execute(sql)
		
		sql = "INSERT INTO AA_Versione(versione) VALUES (0)"
		CALL objConn.execute(sql)
		
	end sub
	
    
	Public sub Init(conn)
		
		conn.CommandTimeout = 600

		set objConn = conn 
		
		if request("ConnString") = "DATA_ConnectionString" then
			set logConn = objConn
		else
			
			set logConn = server.createobject("adodb.connection")
			logConn.open Application("DATA_ConnectionString"),"",""	
			
		end if
		
		CALL WriteLogAdminHttp(logConn, "AA_versione", 0, "AGGIORNAMENTO_DATABASE_AVVIO", "Avvio aggiornamenti database " + request("ConnString"), true)
		%>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
			<caption class="border">Sequenza aggiornamenti</caption>
			<tr><td class="label_no_width" colspan="5">inizializzazione: <%= DateTimeIta(Now()) %>:<%= FixLenght(Second(Now()), "0", 2) %></td></tr>
			<tr>
				<th rowspan="2" class="center">Aggiornamento n&ordm;</th>
				<th rowspan="2">Esito</th>
				<th rowspan="2" class="center">Versione corrente</th>
				<th class="center" colspan="2" style="border-bottom:0px;">tempo di esecuzione (secondi)</th>
			</tr>
			<tr>
				<th class="center">step</th>
				<th class="center">da inizio</th>
			</tr>
		<%'verifica presenza tabella oppure la crea
		if not VersionTableExists(objConn) then
			CreateVersionTable()%>
			<tr>
				<th class="l2_center" colspan="5">Creata tabella per mantenimento versione</th>
			</tr>
		<% else %>
			<tr>
				<th class="l2_center" colspan="5">Tabella versione esistente</th>
			</tr>
		<%end if %>
		
		<tr>
			<td class="label_no_width" colspan="3">
				avvio aggiornamenti: <%= DateTimeIta(Now()) %>:<%= FixLenght(Second(Now()), "0", 2) %>
			</td>
			<% CALL TimeTracer() %>
		</tr>
	<%end sub
	
	
	'esegue aggiornamento
	Public sub Execute(script_list, version)
	
	 	CALL ProtectedExecute(script_list, version, false)
		
	end sub
	
	
	'esegue aggiornamento proteggendo l'aggiornamento stesso da errori.
	Public sub ProtectedExecute(script_list, version, IsProtected)
		CALL ProtectedExecuteRebuild(script_list, version, IsProtected, false)
	end sub
	
	'esegue l'aggiornamento e, se specificato, esegue anche il rebuild delle viste.
	Public sub ProtectedExecuteRebuild(script_list, version, IsProtected, RebuildViews)
		if IsProtected then
			on error resume next
		end if

		last_update_executed = false %>
		<tr>
			<td class="content_center"><%= version %></td>
				<% 'verifica se aggiornamento esguibile
				if cInteger(ReadCurrentDbVersion(objConn)) + 1 <> cInteger(version) then
					'sequenza di aggiornamenti non corretta
					%>
					<td class="content">Aggiornamento non eseguito: 
						<% if cInteger(ReadCurrentDbVersion(objConn)) > cInteger(version) then %>
							VERSIONE DB SUCCESSIVA
						<% elseif cInteger(ReadCurrentDbVersion(objConn)) = cInteger(version) then %>
							VERSIONE CORRENTE
						<% else %>
							<b>VERSIONI NON CONSECUTIVE</b>
						<% end if %>
					</td>
				<% else
					
					'scrive log per aggiornamento
					CALL WriteLogAdmin(logConn, "AA_versione", version, "AGGIORNAMENTO_DATABASE_" & version, "Codice Eseguito:" + vbCrLf + vbCrLf + vbCrLf + script_list)
					
					'esegue aggiornamento
					CALL ExecuteMultipleSql(objConn, script_list, true)

					if err.number<>0 then
						'ramo eseguito solo se in esecuzione protetta da errori
						last_update_executed = false + "" %>
						<td>
							<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
								<tr>
									<td colspan="2" class="errore">ERRORE NELL'ESECUZIONE DELL'AGGIORNAMENTO PROTETTO:</td>
								</tr>
								<tr>
									<td colspan="2" class="note">ATTENZIONE: l'aggiornamento verr&agrave; comunque considerato corretto ed incrementer&agrave; la versione!</td>
								</tr>
								<tr><th colspan="2" class="L2">err object dumping</th></tr>
								<tr>
									<td class="label">Number</td>
									<td class="content"><%= err.number %></td>
								</tr>
								<tr>
									<td class="label">Source</td>
									<td class="content"><%= err.source %></td>
								</tr>
								<tr>
									<td class="label">Description</td>
									<td class="content"><%= err.description %></td>
								</tr>
								<% dim errore
								set errore = server.GetLastError %>
								<tr><th colspan="2" class="L2">asperror object dumping</th></tr>
								<tr>
									<td class="label">ASPCode</td>
									<td class="content"><%= errore.ASPCode %></td>
								</tr>
								<tr>
									<td class="label">Number </td>
									<td class="content"><%= errore.Number %></td>
								</tr>
								<tr>
									<td class="label">Category</td>
									<td class="content"><%= errore.Category %></td>
								</tr>
								<tr>
									<td class="label">ASPDescription</td>
									<td class="content"><%= errore.ASPDescription %></td>
								</tr>
								<tr>
									<td class="label">Description</td>
									<td class="content"><%= errore.Description %></td>
								</tr>
								<tr>
									<td class="label">File </td>
									<td class="content"><%= errore.File %></td>
								</tr>
								<tr>
									<td class="label">Line </td>
									<td class="content"><%= errore.Line %></td>
								</tr>
								<tr>
									<td class="label">Column</td>
									<td class="content"><%= errore.Column %></td>
								</tr>
								<tr>
									<td class="label">Source </td>
									<td class="content"><%= errore.Source %></td>
								</tr>
							</table>
						</td>
						<% 
						err.clear
					else
						last_update_executed = true 
                        count_update_executed = count_update_executed + 1%>
						<td class="content">Aggiornamento eseguito CORRETTAMENTE</td>
					<%end if
					
					if err.number = 0 then
						'aggiorna versione
						CALL objConn.execute("UPDATE AA_Versione SET Versione = Versione + 1", 0, adExecuteNoRecords)
					end if
					
				end if%>
			<td class="content_center"><%= ReadCurrentDbVersion(objConn) %></td>
			<% CALL TimeTracer() %>
		</tr>	
		
		<%if IsProtected then
			on error goto 0
		end if
		
		if RebuildViews then
			'richiama la ricostruzione delle viste
			CALL SqlServer_VIEWS_REBUILD(version)
		end if
		
	end sub
	
	
	'chiude la transazione precedente e ne apre una nuova per sbloccate tutte le risorse precedenti
    'solo se la versione del database e' quella indicata o se e' zero.
	Public sub ReSyncTransaction()
	    if DB.last_update_executed then
            ReSyncTransactionAlways()
		end if
	end sub
	
	'chiude la transazione precedente e ne apre una nuova per sbloccate tutte le risorse precedenti
	Public Sub ReSyncTransactionAlways()
		'chiude transazione precedente
   		objConn.CommitTrans%>
   		
   		<tr>
   			<th class="l2_center" colspan="3"> <strong>Chiusura</strong> e <strong>riapertura</strong> della <strong>transazione</strong> per liberare le risorse bloccate dagli aggiornamenti precedenti</th>
			<% CALL TimeTracer() %>
   		</tr>
   		
   		<%'riapre nuova transazione
   		objConn.BeginTrans
	end sub
	
    'funzione che verifica l'esistenza di un campo nella tabella
	public function FieldExistsInTable(table, field)
		dim sql, rs
		set rs = server.CreateObject("ADODB.RecordSet")
		sql = "SELECT TOP 1 * FROM " & table
		rs.open sql, conn, adOpenstatic, adLockOptimistic, adCmdText
		FieldExistsInTable = FieldExists(rs, field)
		rs.close
		set rs = nothing
	end function
	
	
    Private Sub Rebuild_OnTerminate(TerminateSql)
		if TerminateSql<>"" then %>
            <tr><th class="l2_center" colspan="5">Inizio aggiornamento finale per VISTE e STORED PROCEDURE</th></tr>
            <%CALL ExecuteMultipleSql(objConn, TerminateSql, true) %>
            <tr>
                <td class="content_center">finale</td>
                <td class="content">Aggiornamento eseguito CORRETTAMENTE</td>
                <td class="content_center">&nbsp;</td>
				<% CALL TimeTracer() %>
			</tr>
       	<%else%>
            <tr><th class="l2_center" colspan="5">aggiornamento finale di chiusura non presente</th></tr>
       	<%end if
	end sub
	
	
'******************************************************************************************************************************************************************************************************
'PROCEDURE PUBBLICHE DI GESTIONE DELL'INDICE
'******************************************************************************************************************************************************************************************************
	
	'procedura che aggiorna tutti i contenuti dell'indice del tipo indicato, secondo l'ordinamento specificato
	Public sub RebuildIndex_RefreshContents(tabella, ordinamento) %>
		<tr>
			<th class="l2_center" colspan="5"> Aggiornamento contenuti dell'indice per la tabella <%= tabella %></th>
		</tr>
		<%
		dim rsc, field_chiave
		set rsc = server.createObject("ADODB.recordset")
		if DB.last_update_executed then %>
			<%
			sql = "SELECT TOP 1 * FROM tb_siti_tabelle WHERE tab_name LIKE '" + ParseSql(tabella, adChar) + "'"
			rs.open sql, objConn, adOpenStatic, adLockOptimistic, adCmdText
			
			while not rs.eof 
				if lcase(tabella) = lcase(tabRaggruppamentoTable) then
					field_chiave = "idx_content_id"
				else
					field_chiave = lcase(rs("tab_field_chiave"))
				end if
				
				sql = " SELECT * FROM " + tabella  + " ORDER BY " + ordinamento + IIF(lcase(ordinamento) <>field_chiave, ", " & field_chiave, "")
				rsc.open sql, objConn, adOpenStatic, adLockOptimistic, adCmdText %>
				<tr>
					<td class="content">&nbsp;</td>
					<td class="content" colspan="4">Aggiornamento di n&ordm; <%= rsc.recordcount %> contenuti</td>
				</tr>
				<%while not rsc.eof

					CALL Index_UpdateItem(objConn, rs("tab_name"), rsc(field_chiave), false)

					rsc.movenext
				wend
				rsc.close
				rs.movenext
			wend %>
			<tr>
                <td class="content_center"></td>
                <td class="content">Aggiornamento contenuti tabella <%= tabella %> eseguito correttamente</td>
                <td class="content_center">&nbsp;</td>
				<% CALL TimeTracer() %>
			</tr>
			<% 
			rs.close
		else %>
			<tr>
                <td class="content_center"></td>
                <td class="content warning">Aggiornamento non eseguito: l'aggiornamento precedente non &egrave; stato eseguito.</td>
                <td class="content_center">&nbsp;</td>
				<% CALL TimeTracer() %>
			</tr>
		<% end if
		set rsc = nothing
	end sub
	
	
	'procedura che esegue un aggiornamento bulk dell'indice
	Public Sub RebuildIndex_BULK()%>
		<tr>
			<th class="l2_center" colspan="5"> Esecuzione dell'aggiornamento completo dell'indice in modalità sequenziale.</th>
		</tr>
		<%if DB.last_update_executed then
			sql = " SELECT top 5 idx_id, co_F_key_id, tab_name " + _
				  " FROM v_indice " + _
				  " WHERE idx_modAdmin_id <> 9999 " + _
				  " ORDER BY idx_livello, idx_id "
			rs.open sql, objConn, adOpenDynamic, adLockOptimistic, adCmdText
			
			while not rs.eof
				
				Index.DisableRicorsione = true
				
				Session("ID_ADMIN") = 9999
				CALL Index_UpdateItem(DB.objconn, rs("tab_name"), rs("co_F_key_id"), false)
				
				rs.movenext
			wend
			rs.close %>
			<tr>
                <td class="content_center"></td>
                <td class="content">aggiornamento eseguito correttamente.</td>
                <td class="content_center">&nbsp;</td>
				<% CALL TimeTracer() %>
			</tr>
		<% else %>
			<tr>
                <td class="content_center"></td>
                <td class="content warning">aggiornamento indice non eseguito: l'aggiornamento precedente non &egrave; stato eseguito.</td>
                <td class="content_center">&nbsp;</td>
				<% CALL TimeTracer() %>
			</tr>
		<% end if
	end sub
	
	
	'procedura che esegue l'aggiornamento di tutte le voci dell'indice eseguendo le operazioni ricorsive a partire dalla root
	Public Sub RebuildIndex_OperazioniRicorsive() %>
		<tr>
			<th class="l2_center" colspan="5"> Esecuzione delle operazioni ricorsive su tutto l'indice.</th>
		</tr>
		<%if DB.last_update_executed then
			
			sql = "SELECT co_F_key_id, tab_name FROM v_indice WHERE idx_livello = 0"
			rs.open sql, objConn, adOpenDynamic, adLockOptimistic, adCmdText
			
			while not rs.eof
				CALL Index_UpdateItem(DB.objconn, rs("tab_name"), rs("co_F_key_id"), false)
				rs.movenext
			wend
			
			rs.close %>
			<tr>
                <td class="content_center"></td>
                <td class="content">operazioni ricorsive indice eseguite correttamente.</td>
                <td class="content_center">&nbsp;</td>
				<% CALL TimeTracer() %>
			</tr>
		<% else %>
			<tr>
                <td class="content_center"></td>
                <td class="content warning">operazioni non eseguite: l'aggiornamento precedente non &egrave; stato eseguito.</td>
                <td class="content_center">&nbsp;</td>
				<% CALL TimeTracer() %>
			</tr>
		<% end if
	end sub
	
	
	
'******************************************************************************************************************************************************************************************************
'PROCEDURE PRIVATE
'******************************************************************************************************************************************************************************************************

	private sub TimeTracer()
		dim NowTime
		NowTime = timer %>
		<td class="content_right">
			<%= FormatPrice(((NowTime - TimerPrevious)), 2, true) %>
		</td>
		<td class="content_right">
			<%= FormatPrice(((NowTime - TimerStart)), 2, true) %>
		</td>
		<%TimerPrevious = NowTime	
	end sub


'******************************************************************************************************************************************************************************************************
'TOOLS GENERICI
'******************************************************************************************************************************************************************************************************

	public sub SqlServer_VIEWS_REBUILD(versione)
		CALL SqlServer_VIEWS_REBUILD_EX(versione,"")
	end sub

	'esegue il rebuil delle viste collegate
	' vista è il nome della vista da aggiornare
	public sub SqlServer_VIEWS_REBUILD_EX(versione, vista)
		dim value
		
		if cInteger(ReadCurrentDbVersion(objConn)) = cInteger(versione) AND _
		   DB_Type(objConn) = DB_SQL then%>
			<tr>
				<th class="l2_center" colspan="5"> <strong>Ricostruzione</strong> ed <strong>aggiornamento</strong> delle <strong>VISTE</strong> presenti nel database</th>
			</tr>
	   		<% dim sql, rs, rsT, RebuildSQL
			set rs = server.CreateObject("ADODB.RecordSet")
	   		sql = "SELECT name, xtype FROM sysobjects WHERE (( xtype = 'V' OR xtype = 'FN') OR xtype = 'IF') "
			if vista <> "" then
				sql = sql & " AND name LIKE '" & vista & "'"
			end if
			rs.open sql, objConn, adOpenstatic, adLockOptimistic, adCmdText

			RebuildSQL = ""
			
			while not rs.eof
				if instr(1, rs("name"), "sys", vbTextCompare)<1 then %>
					<tr>
						<td class="content">&nbsp;</td>
						<td class="content" colspan="4"><%= rs("name") %></td>
					</tr>
					<%
					RebuildSQL = RebuildSQL + _
								 DropObject(objConn, rs("name"), IIF(instr(1, cString(rs("xtype")), "V", vbTextCompare)>0, "VIEW", "FUNCTION"))
					set rsT = objConn.execute("EXEC sp_helptext '" & rs("name") & "' ")
					while not rsT.eof
						value = cString(rsT("text"))
						value = trim(value)
						value = replace(value, vbLf, "")
						value = replace(value, vbCr, "")
						if value<>"" then
							RebuildSQL = RebuildSQL + value + vbCrLf
						end if
						rst.movenext
					wend
					RebuildSQL = RebuildSQL + ";"
				end if
				rs.movenext
			wend
			rs.close
	
			CALL ExecuteMultipleSql(objConn, RebuildSQL, true)
			
			set rs = nothing
		end if
   	end sub
	
	
	'esegue ricostruzione dei campi testo uniformando il codice del collate al "database default"
	public sub SqlServer_COLLATE_REBUILD(versione)
		dim UpdateSQL, sql
		
		if cInteger(ReadCurrentDbVersion(objConn)) = cInteger(versione) AND _
		   DB_Type(objConn) = DB_SQL then%>
			<tr>
				<th class="l2_center" colspan="3">Aggiornamento COLLATE dei campi testo</th>
			</tr>
	   		<%'aggiornamento corrente per database SQL: genera sql per aggiornamento
			UpdateSQL = ""
			sql = " SELECT (sysobjects.name) AS table_name, " + _
				  "	(syscolumns.name) AS column_name, " + _
				  " (syscolumns.xtype) AS column_type, " + _
				  " (syscolumns.length) AS column_size " + _
				  " FROM syscolumns INNER JOIN sysobjects ON syscolumns.id = sysobjects.id " + _
				  " WHERE  sysobjects.xtype = 'U' AND " + _
				  "		   syscolumns.xtype IN (167, 175, 231, 239, 35, 99) AND " + _
				  "		   syscolumns.isnullable = 1 AND " + _
				  "        NOT (sysobjects.name LIKE '%temp_%') AND " + _
				  "        NOT EXISTS(SELECT 1 FROM sysreferences " + _
				  				    " WHERE (sysreferences.fkeyid = syscolumns.id AND sysreferences.fkey1 = syscolumns.colid) OR " + _
								          " (sysreferences.rkeyid = syscolumns.id AND sysreferences.rkey1 = syscolumns.colid)) " + _
				  " ORDER BY table_name, column_name "
			rs.open sql, DB.objConn, adOpenstatic, adLockOptimistic, adCmdText
			
			rs.movefirst
			while not rs.eof
				Select case cInteger(rs("column_type")) 
					case 35, 99
						'crea colonna temporanea, copia i dati, cancella vecchia colonna e la ricrea.
						UpdateSQL = UpdateSQL + _
									" ALTER TABLE " + rs("table_name") + " ADD " + _
									" temp_column_collate_" + rs("column_name") + IIF(cInteger(rs("column_type"))=35, " text ", " ntext ") + " COLLATE database_default NULL ; " + vbCrLf + _
									" UPDATE " + rs("table_name") + " SET temp_column_collate_" + rs("column_name") + " = " + rs("column_name") + "; " + vbCrLf + _
									" ALTER TABLE " + rs("table_name") + " DROP COLUMN " + rs("column_name") + "; " + vbCrLf + _
									" ALTER TABLE " + rs("table_name") + " ADD " + _
									rs("column_name") + IIF(cInteger(rs("column_type"))=35, " text ", " ntext ") + " COLLATE database_default NULL ; " + vbCrLf + _
									" UPDATE " + rs("table_name") + " SET " + rs("column_name") + " = temp_column_collate_" + rs("column_name") + "; " + vbCrLf + _
 									" ALTER TABLE " + rs("table_name") + " DROP COLUMN temp_column_collate_" + rs("column_name") + "; " + vbCrLf
					case else
						UpdateSQL = UpdateSQL + _
									" ALTER TABLE " + rs("table_name") + " ALTER COLUMN " + rs("column_name")
						Select case cInteger(rs("column_type")) 
							case 167
								UpdateSQL = UpdateSQL + " varchar(" & rs("column_size") & ") "
							case 175
								UpdateSQL = UpdateSQL + " char(" & rs("column_size") & ") "
							case 231
								UpdateSQL = UpdateSQL + " nvarchar(" & rs("column_size")\2 & ") "
							case 239
								UpdateSQL = UpdateSQL + " nchar(" & rs("column_size")\2 & ") "
						end select
						UpdateSQL = UpdateSQL + " COLLATE database_default NULL ; " + vbCrLf
				end select
				rs.movenext
			wend
			rs.close
			
			CALL ExecuteMultipleSql(objConn, UpdateSQL, true)
			
		end if
	end sub
	
end class


%>