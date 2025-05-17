<%
'************************************************************************************************
'VERSIONE CON I COMPONENTI:
const OBJ_PERSITS = "PERSITS"
const OBJ_SOFTARTISANS = "SOFTARTISANS"
const OBJ_SOFTARTISANS_OLD = "SOFTARTISANS_OLD"		'vecchia versione del softartisans vers. <4


'Ultimo aggiornamento:		25/11/2004
'Autore						Nicola

'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'classi dichiarate nel file:

'CLASS UpLoad		==>		Gestione completa in una pagina di tutte le operazioni di Upload dei file
'CLASS DeleteFile	==>		Gestione delle operazioni di cancellazione di un file legato o meno a database
'CLASS UploadFile	==>		Gestione delle operazioni di salvataggio di un file legato o meno a database


'***********************************************************************
'INTERFACCE
'***********************************************************************
'PARAMETRI COMUNI:
'request.quersytring("ID")		Indica l'id del record


'***********************************************************************
'CLASS UPLOAD	********************************************************

'PARAMETRI
'	Public Connection_String	'connessione al DB
'	Public Table_Name			'nome della tabella
'	Public ID_Field				'campo Identity della tabella
'	Public SQL_Nominativo		'Stringa sql per comporre il nome del record
'	Public File_Field			'Nome campo contenente il file
'	Public File_Path			'Percorso FISICO COMPLETO directory

'	Public ShowConsigli	'indica se devono essre mostrati i "consigli dei file"
'	Public OverWrite		'indica se nell'upload si puo' sovrascrivere il file se gia' esistente
'	Public OnlyExtensionAllowed ' indica se deve esseer attivato il controllo sulle estensioni
	
'	Public Border_color			'colore bordi tabelle
'	Public Bg_testata			'colore sfondo bordi
	
'	Public Stile_Input			'stile input di testo
'	Public Stile_Submit			'stile pulsanti submit
'	Public Stile_Titoli			'stile testo titoli
'	Public Stile_Testata		'stile testo testata
'	Public Stile_testo			'stile testo normale
'	Public OperationOK			'Indica se l'operazione e' andata a buon fine o meno

'METODI
'public sub Gestione_Completa_Record()
	'Legge i dati da DB e, se presente il file gestisce la cancellazione, altrimenti gestisce l'upload.
	'per fare la cancellazione e l'upload utilizza le sottoclassi.
	'si appoggia al database. Puo' lavorare su un campo di un record, oppuer considerare il record come
	'registrazione di un file.

	

'***********************************************************************
'CLASS DeleteFile  *****************************************************

'PARAMETRI	
'	Public File_Path			'Percorso FISICO COMPLETO directory
'	Public File_Name			'Nome del file
	
	'campi usati solo per aggiornamento record
'	Public Connection_String	'connessione al DB
'	Public Table_Name			'nome della tabella
'	Public ID_Field				'campo Identity della tabella
'	Public File_Field			'Nome campo contenente il file
'	Public Update_Record		'DEFAULT = "" Indica se alla fine deve essere eseguita la cancellazione del file
								'dal DB. Se Update_Record = "UPDATE" viene aggiornato il record
								'Se Update_Record = "DELETE" viene cancellato il record completo
	
'	Public Stile_Input			'stile input di testo
'	Public Stile_Submit			'stile pulsanti submit
'	Public Stile_Titoli			'stile testo titoli
'	Public Stile_testo			'stile testo normale.
'	Public OperationOK			'Indica se l'operazione e' andata a buon fine o meno

'METODI
'public sub Delete()
	'gestisce completamente la cancellazione. Se necessario legge il file da DB.
	'Utilizza i metodi FormDelete() per creare il form di richiesta di cancellazione e
	'il metodo FileDelete() per eseguire la cancellazione.
	
'public sub FormDelete()
	'genera il form per la richiesta di cancellazione. se il nome del file e' vuoto lo va a leggere dal database,
	' ma se i parametri di accesso al database non sono impostati genera un errore.
	
'public sub FileDelete()
	'cancella il file e genera un report di cancellazione
	'se richiesto e se i parametri sono impostati, aggiorna il database cancellando il record o modificandolo.

	
'***********************************************************************
'CLASS UploadFile  *****************************************************

'PARAMETRI	
'	Public Connection_String'connessione al DB
'	Public Table_Name		'nome della tabella
'	Public ID_Field			'campo Identity della tabella
'	Public File_Field		'Nome campo contenente il file
'	Public File_Path		'Percorso FISICO COMPLETO directory

'	Public ShowConsigli	'indica se devono essre mostrati i "consigli dei file"
'	Public OverWrite		'indica se nell'upload si puo' sovrascrivere il file se gia' esistente
'	Public OnlyExtensionAllowed ' indica se deve esseer attivato il controllo sulle estensioni

'	Public Update_Record	'DEFAULT = "" Indica se alla fine deve essere eseguita la cancellazione del file
							'dal DB. Se Update_Record = "UPDATE" viene aggiornato il record
							'Se Update_Record = "INSERT" viene inserito un nuovo record
	
'	Public Border_color		'colore bordi tabelle
'	Public Stile_Input		'stile input di testo
'	Public Stile_Submit		'stile pulsanti submit
'	Public Stile_Titoli		'stile testo titoli
'	Public Stile_testo		'stile testo normale	
'	Public OperationOK			'Indica se l'operazione e' andata a buon fine o meno

'METODI
'sub public Upload()
	'Gestisce completamente l'invio di un file su server. Se necessario si connette al DB.
	'Utilizza il metodo FormUpload() per generare il form di richiesta del file
	'Utilizza il metodo FileUpload() per fare il salvataggio del file, se richiesto aggiorna/aggiunge il record nel DB
	
'sub public FormUpload()
	'Genera il form per la richiesta di un file.

'sub public FileUpload()
	'Invia e salva il file su server.
	'Se richiesto e se i parametri sono impostati, aggiorna il database aggiungendo un nuovo record o modificando
	'il record esistente.
%>

<%



'*******************************************************************************************************
'*******************************************************************************************************
'*******************************************************************************************************
'*******************************************************************************************************
class UpLoad
	
	Public Connection_String'connessione al DB
	Public Table_Name		'nome della tabella
	Public ID_Field			'campo Identity della tabella
	Public SQL_Nominativo	'Stringa sql per comporre il nome del record
	Public File_Field		'Nome campo contenente il file
	Public File_Path		'Percorso FISICO COMPLETO directory
	Public External_Key		'Nome delle chiave esterna (es. ID_AZIENDA su B2B)
	public External_Key_Value ' Valore da preimpostare per la chiave esterna
	
	Public ShowConsigli	'indica se devono essre mostrati i "consigli dei file"
	Public OverWrite		'indica se nell'upload si puo' sovrascrivere il file se gia' esistente
	Public OnlyExtensionAllowed ' indica se deve esseer attivato il controllo sulle estensioni
	
	Public Border_color		'colore bordi tabelle
	Public Bg_testata		'colore sfondo bordi
	
	Public Stile_Input		'stile input di testo
	Public Stile_Submit		'stile pulsanti submit
	Public Stile_Titoli		'stile testo titoli
	Public Stile_Testata	'stile testo testata
	Public Stile_testo		'stile testo normale
	Public OperationOK		'Indica se l'operazione e' andata a buon fine o meno

Private conn, rs, sql, DEL_OBJ , UPL_OBJ
	
Private Sub Class_Initialize()
	ShowConsigli = true
	OverWrite = TRUE
	OnlyExtensionAllowed = FALSE
	OperationOK = FALSE
end sub

private sub Class_Terminate()

end sub

'******************************************************
'FUNZIONI DI INTERFACCIA PUBBLICA'
'******************************************************
public sub Gestione_Completa_Record()
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open Connection_String,"",""

	set rs = Server.CreateObject("ADODB.RecordSet")
	
	sql = "SELECT (" & SQL_Nominativo & ") AS NOMINATIVO, " & File_Field 
	sql = sql + " FROM " & table_Name & " WHERE " & ID_Field & "=" & cIntero(request("ID"))
	rs.open sql, conn, adOpenStatic, adLockOptimistic
	
	response.write "<table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
	if rs.recordcount > 0 then
		response.write "	<tr><td height=""20"" bgcolor=""" & Bg_testata & """ style=""border:1px solid " & Border_Color & ";"">"
		response.write "		<table align=""right"" cellspacing=""0"" cellpadding=""0"" border=""0"" width=""100%"">"
		response.write "			<tr><td align=""left""><font " & Stile_testata & ">&nbsp;" & rs("NOMINATIVO") & "</font></td>"
		response.write "				<td style=""padding-right:1px;"" align=""right"" ><input type=""button"" name=""close"" value=""CHIUDI"" onclick=""window.close();"" " & stile_submit & "></td></tr>"
		response.write "		</table>"
		response.write "	</td></tr>"
		response.write "	<tr><td ><font style=""font:5px Arial"">&nbsp;</font></td></tr>"
		response.write "<tr><td height=""24"" style=""border:1px solid " & Border_Color & ";"">"
			if NOT ISNULL(rs(File_Field)) AND rs(File_Field)<>"" then
				
				'File gia' presente: gestione cancellazione
				SET DEL_OBJ = New DeleteFile
				
					DEL_OBJ.File_Path = File_Path		
					DEL_OBJ.File_Name = rs(File_Field)	
					DEL_OBJ.Connection_String = Connection_String
					DEL_OBJ.Table_Name = Table_Name
					DEL_OBJ.ID_Field = ID_Field
					DEL_OBJ.File_Field = File_Field
					DEL_OBJ.Update_Record = "UPDATE"
					DEL_OBJ.Stile_Input = Stile_Input
					DEL_OBJ.Stile_Submit = Stile_Submit
					DEL_OBJ.Stile_Titoli = Stile_Titoli
					DEL_OBJ.Stile_testo	= Stile_Testo
				
				DEL_OBJ.Delete()
				
				OperationOK = DEL_OBJ.OperationOK
			else
				'File non presente: unica operazione invio
				
				SET UPL_OBJ = New UploadFile
					UPL_OBJ.Connection_String = Connection_String
					UPL_OBJ.Table_Name = Table_Name
					UPL_OBJ.ID_Field = ID_Field
					UPL_OBJ.File_Field = File_Field
					UPL_OBJ.File_Path = File_Path
					UPL_OBJ.ShowConsigli = ShowConsigli
					UPL_OBJ.OverWrite = OverWrite
					UPL_OBJ.OnlyExtensionAllowed = OnlyExtensionAllowed
					UPL_OBJ.Update_Record = "UPDATE"
					UPL_OBJ.Border_color = Border_color
					UPL_OBJ.Stile_Input = Stile_Input
					UPL_OBJ.Stile_Submit = Stile_Submit
					UPL_OBJ.Stile_Titoli = Stile_Titoli
					UPL_OBJ.Stile_testo = Stile_testo
					'Public External_Key		'Nome delle chiave esterna (es. ID_AZIENDA su B2B)
					'public External_Key_Value ' Valore da preimpostare per la chiave esterna
					UPL_OBJ.External_Key = External_Key
					UPL_OBJ.External_Key_Value = External_Key_Value
				UPL_OBJ.Upload()
				OperationOK = UPL_OBJ.OperationOK
				
			end if
		response.write "</td></tr>"
		response.write "	<tr><td ><font style=""font:5px Arial"">&nbsp;</font></td></tr>"
	else
		response.write "<tr><td height=""25"" align=""center""><font " & Stile_titoli & ">ERRORE NELL'APPLICAZIONE</font></td></tr>"
	end if
	response.write "	<tr><td height=""20"" bgcolor=""" & Bg_testata & """style=""border:1px solid " & Border_Color & ";"">"
	response.write "		<table align=""right"" cellspacing=""0"" cellpadding=""0"" border=""0"" width=""100%"">"
	response.write "			<tr><td colspan=""2"" style=""padding-right:1px;"" align=""right"" ><input type=""button"" name=""close"" value=""CHIUDI"" onclick=""window.close();"" " & stile_submit & "></td></tr>"
	response.write "		</table>"
	response.write "	</td></tr>"
	response.write "</table>"
	rs.close
	conn.close
end sub


'******************************************************
'FUNZIONI PRIVATE'
'******************************************************

end class




'*******************************************************************************************************
'*******************************************************************************************************
'*******************************************************************************************************
'*******************************************************************************************************
'CLASSE PER LA CANCELLAZIONE DI UN FILE

class DeleteFile
	
	Public File_Path		'Percorso FISICO COMPLETO directory
	Public File_Name		'Nome del file
	
	'campi usati solo per aggiornamento record
	Public Connection_String'connessione al DB
	Public Table_Name		'nome della tabella
	Public ID_Field			'campo Identity della tabella
	Public File_Field		'Nome campo contenente il file
	Public Update_Record	'DEFAULT = "" Indica se alla fine deve essere eseguita la cancellazione del file
							'dal DB. Se Update_Record = "UPDATE" viene aggiornato il record
							'Se Update_Record = "DELETE" viene cancellato il record completo
	
	Public Stile_Input		'stile input di testo
	Public Stile_Submit		'stile pulsanti submit
	Public Stile_Titoli		'stile testo titoli
	Public Stile_testo		'stile testo normale
	Public OperationOK		'Indica se l'operazione e' andata a buon fine o meno
	
	
Private Sub Class_Initialize()
	Update_Record = ""
	OperationOK = FALSE
end sub

private sub Class_Terminate()

end sub

'******************************************************
'FUNZIONI DI INTERFACCIA PUBBLICA'
'******************************************************

public sub Delete()
	
	dim rs, sql, conn
	
	if Connection_String <> "" AND File_Name="" and table_Name<>"" and File_Field<>"" and ID_Field<>"" then
	
		set conn = Server.CreateObject("ADODB.Connection")
		conn.open Connection_String,"",""
	
		'lettura nome file se non passato dal parametro
		set rs = Server.CreateObject("ADODB.RecordSet")
		
		sql = "SELECT " & File_Field & " FROM " & table_Name & " WHERE " & ID_Field & "=" & cIntero(request.querystring("ID"))
		rs.open sql, conn, adOpenStatic, adLockOptimistic
		if rs.recordcount > 0 then
		
			File_Name = rs(File_Field)
			
		end if
		rs.close
		conn.close
	end if
	
	if File_Name <> "" then
		if Request.ServerVariables("REQUEST_METHOD")<>"POST" then
			
			'prima esecuzione pagina: Form per richiesta cancellazione
			FormDelete()
			
		else
			
			'seconda esecuzione pagina: esecuzione cancellazione
			FileDelete()
			
		end if
	else
	
		response.write "	<table align=""right"" cellspacing=""0"" cellpadding=""0"" border=""0"" width=""100%"">"
		response.write "		<tr><td align=""left""><font " & stile_titoli & ">&nbsp;PARAMETRI NON CORRETTI&nbsp;</font></td></tr>"
		response.write "	</table>"
		
	end if
		
end sub


'**********************************************************************
Public Sub FormDelete()

	'scrive il form per la richiesta di cancallazione
	if File_Name <> "" then%>
		<form method="POST" action="" id="form1" name="form1">
			<table align="right" cellspacing="0" cellpadding="0" border="0" width="100%">
				<tr><td style="font-size:5px;">&nbsp;</td></tr>
				<tr>
					<td align="left" class="<%= stile_testo %>">
						&nbsp;FILE:&nbsp;
						<input type="text" <%= stile_input %> size="<%= (len(File_Name) + 5) %>" name="FILE1" value="<%= File_Name %>" READONLY>
					</td>
					<td style="padding-right:4px;" align="right">
						<input type="submit" name="B1" value="CANCELLA" <%= stile_Submit %>>
					</td>
				</tr>
				<tr><td style="font-size:5px;">&nbsp;</td></tr>
			</table>
		</form>
	<%else%>
		<table align="right" cellspacing="0" cellpadding="0" border="0" width="100%">
				<tr>
					<td align="left" class="<%= stile_titoli %>">
						&nbsp;NOME FILE NON PRESENTE&nbsp;
					</td>
				</tr>
		</table>
	<%end if
	
end Sub


'************************************************************************
Public Sub FileDelete()
	dim Fso, Path, sql, conn
	'cancella il file e aggiorna i dati (se richiesto)
	
	if File_Name <> "" AND (Update_Record="" OR (Update_Record="UPDATE" AND table_Name<>"" and File_Field<>"" and ID_Field<>"") OR (Update_Record="DELETE" AND table_Name<>"" and ID_Field<>"")) then
		
		'creazione oggetto filesystem
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		Path = File_path & "\" & File_Name
		Path = replace(Path, "\\", "\")
		
		response.write "	<table align=""right"" cellspacing=""0"" cellpadding=""0"" border=""0"" width=""100%"">"
		response.write "		<tr><td align=""left""><font " & stile_titoli & ">&nbsp;"
		
		if fso.FileExists(Path) then
		
			'cancellazione del file
			fso.DeleteFile(Path)
			response.write "	File """ & File_Name & """cancellato correttamente."
			
			OperationOK = TRUE
			
		else
			response.write "	File """ & File_Name & """ non trovato."
		
		end if
					
		'cancellazione file da DB (se richiesta)
		if Update_record<>"" AND request.Querystring("ID")<>"" then
		
			'Apertura connessione
			set conn = Server.CreateObject("ADODB.Connection")
			conn.open Connection_String,"",""
			
			if Update_Record="UPDATE" AND table_Name<>"" and File_Field<>"" and ID_Field<>"" then
			
				'se presenti tutti i campi esegue UPDATE su record
				sql = "UPDATE " & table_Name & " SET " & File_Field & "='' WHERE " & ID_Field & "=" & cIntero(request.Querystring("ID"))
				conn.execute(sql)
				
			elseif Update_Record="DELETE" AND table_Name<>"" and ID_Field<>"" then
			
				'se presenti solo nome tabella ed ID cancella il record
				sql = "DELETE * FROM " & table_Name & " WHERE " & ID_Field & "=" & cIntero(request.Querystring("ID"))
				conn.execute(sql)
				
			end if
			
			conn.close
		end if					
					
		response.write "		&nbsp;</font></td></tr>"
		response.write "	</table>"
		
	else
		
		response.write "	<table align=""right"" cellspacing=""0"" cellpadding=""0"" border=""0"" width=""100%"">"
		response.write "		<tr><td align=""left""><font " & stile_titoli & ">&nbsp;PARAMETRI NON CORRETTI&nbsp;</font></td></tr>"
		response.write "	</table>"
		
	end if
	
end sub	
	

'******************************************************
'FUNZIONI PRIVATE'
'******************************************************

	
end class



'*******************************************************************************************************
'*******************************************************************************************************
'*******************************************************************************************************
'*******************************************************************************************************
'CLASSE PER L'UPLOAD DI UN FILE

class UploadFile

	Public Connection_String'connessione al DB
	Public Table_Name		'nome della tabella
	Public ID_Field			'campo Identity della tabella
	Public File_Field		'Nome campo contenente il file
	Public File_Path		'Percorso FISICO COMPLETO directory
	Public External_Key		'Nome delle chiave esterna (es. ID_AZIENDA su B2B)
	public External_Key_Value ' Valore da preimpostare per la chiave esterna
	
	Public ShowConsigli	'indica se devono essre mostrati i "consigli dei file"
	Public OverWrite		'indica se nell'upload si puo' sovrascrivere il file se gia' esistente
	Public OnlyExtensionAllowed ' indica se deve esseer attivato il controllo sulle estensioni
	
	Public Update_Record	'DEFAULT = "" Indica se alla fine deve essere eseguita la cancellazione del file
							'dal DB. Se Update_Record = "UPDATE" viene aggiornato il record
							'Se Update_Record = "INSERT" viene inserito un nuovo record
	
	Public Border_color		'colore bordi tabelle
	Public Stile_Input		'stile input di testo
	Public Stile_Submit		'stile pulsanti submit
	Public Stile_Titoli		'stile testo titoli
	Public Stile_testo		'stile testo normale	
	
	Public OperationOK		'Indica se l'operazione e' andata a buon fine o meno
	Public File_Name		'nome del file caricato
		
Private Sub Class_Initialize()
	ShowConsigli = true
	OverWrite = TRUE
	OnlyExtensionAllowed = FALSE
	OperationOK = FALSE
end sub

private sub Class_Terminate()

end sub

'******************************************************
'FUNZIONI DI INTERFACCIA PUBBLICA'
'******************************************************

public sub Upload()
	if Request.ServerVariables("REQUEST_METHOD")<>"POST" then
		'Form per la scelta del file
		
		FormUpload()
		
	else
		'Esecuzione Invio File
		
		FileUpload()
		
	end if
end sub
	

	public sub FormUpload()%>
		<form enctype="multipart/form-data" method="post" action="?ID=<%= request.querystring("ID") %>&CLASS=1" id="form1" name="form1">
			<input type="hidden" name="RelativePath" value="<%= replace(File_Path, Application("IMAGE_PATH"), "") %>">
			<input type="hidden" name="ReplacedPath" value="<%= instr(1, File_Path, Application("IMAGE_PATH"), vbTextCompare) %>">
			<table align="right" cellspacing="0" cellpadding="0" border="0" width="100%">
				<% if ShowConsigli then %>
					<tr>
						<td align="left" <%= stile_titoli %> colspan="2">
							&nbsp;Caratteristiche consigliate dei file:
						</td>
					</tr>
					<tr>
						<td colspan="2" style="border-bottom:1px solid <%= Border_color %>;">
							<ul type="disc" style="margin-bottom:4px;">
								<li <%= stile_testo %>>il nome del file non deve contenere spazi al suo interno;</li>
								<li <%= stile_testo %>>il peso massimo consigliato &egrave; di 20kb;</li>
								<li <%= stile_testo %>>le immagini devono avere una risoluzione di 72  pixel/cm;</li>
							</ul>
						</td>
					</tr>
				<% end if %>
				<tr><td style="font-size:5px;">&nbsp;</td></tr>
				<tr>
					<td align="left" <%= stile_testo %>>
						&nbsp;FILE:&nbsp;
						<input type="File" <%= stile_input %> name="INPUT_FILE" value="">
					</td>
					<td align="right" style="padding-right:2px;">
						<input type="submit" name="B1" value="INVIA" <%= stile_Submit %>>
					</td>
				</tr>
				<% if isNull(OverWrite) then %>
					<tr>
						<td <%= stile_testo %>>
							Sovrascrivi il file se gi&agrave; esistente
							<input type="checkbox" name="OVERWRITE" value="1" class="checkbox">
						</td>
					</tr>
				<% end if %>
				<tr><td style="font-size:5px;">&nbsp;</td></tr>
			</table>
		</form>
	<%end sub


public sub FileUpload()
	dim upl, InternalFilePath
	
	OperationOK = true
	set upl = new UploadObject
	upl.OverWrite = true		'imposta sovrascrittura (eventuali blocchi vengono eseguiti da codice)
	
	if cInt("0" & upl.Form("ReplacedPath"))>0 then
		InternalFilePath = Application("IMAGE_PATH") & "\" & upl.Form("RelativePath")
	else
		InternalFilePath = File_Path
	end if 
	InternalFilePath = replace(InternalFilePath, "\\", "\")

'response.write InternalFilePath
'response.end

	%>
	<table cellspacing="0" cellpadding="0" width="100%">
		<tr>
			<td>
				<% if upl.isEmpty then 
					'file non caricato
					OperationOK = false%>
					<span <%= stile_titoli %>>&nbsp;File mancante: Scegliere prima un file e quindi inviarlo.</span>
				<% else
					'file caricato correttamente
					'recupera nome file e dimensione
					File_Name = upl.FileName
					if request.querystring("class")="1" then		'indica se il form che salva il file viene generato dalla classe
						'esegue controlli per correttezza file
						
						'verifica se il file caricato ha una estensione riconosciuta e valida.
						if OnlyExtensionAllowed then
							dim Extension
							Extension = File_Extension( upl.FileName )
							if instr(1, EXTENSION_ALLOWED, " " & Extension & " " , vbTextCompare)<1 then
								'file non riconosciuto
								OperationOK = false %>
								<span <%= stile_titoli %>>
									Per problemi di sicurezza il tipo di file che si sta tentando di caricare non
									&egrave; stato riconosciuto dal server come "tipo sicuro". Se possibile comprimere il file in un formato ZIP.<br>
									Per ulteriori informazioni contattare il webmaster.
								</span>
							<%end if
						end if

						'verifica se il file esiste gia'
						if (not OverWrite) OR (isNull(OverWrite) AND upl.form("overwrite")="") then
							if upl.FileExists(InternalFilePath) then
								'file gia' esistente
								OperationOK = false %>
								<span <%= stile_titoli %>>
									Impossibile caricare il file perch&egrave; &egrave; presente un file con lo stesso nome.<br>
									<% if isNull(OverWrite) then %>
										Per sovrascrivere il file selezionare l'opzione "Sovrascrivi file se gi&agrave; esistente" al momento del caricamento.
									<% else %>
										Per caricare il file cambiarne il nome o cancellare prima il file esistente.<br>
									<% end if %>
								</span>
							<%end if
						end if
					end if
					
					if OperationOK then
						'controlli eseguiti correttamente: salvo definitivamente il file
						'On error resume next
						err.clear
						
						'salva il file
						upl.Save(InternalFilePath)
						
						on error goto 0
						if err.number <> 0 then
							OperationOK = false%>
							<span <%= stile_titoli %>>
								Errore interno nel trasferimento del file (Cod. <%= Err.number %>).<br>
								Ritentare il caricamento, se il problema persiste contattare il webmaster.
							</span>
						<%else
							OperationOK = true%>
							<span <%= stile_titoli %>>
								Trasferimento concluso correttamente.<br>
								E' stato caricato il file "<%= upl.FileName %>" di dimensione <%= File_Dimension( upl.FileSize ) %>.
							</span>
						<%end if
						
					end if
				end if%>
			</td>
		</tr>
	</table>
	<%
	
	if OperationOK then
		'registra i dati del file
		
		if (Update_Record="UPDATE" OR Update_Record="INSERT") AND table_Name<>"" and File_Field<>"" and ID_Field<>"" then
				dim conn, rs, sql
				'Apertura connessione
				set conn = Server.CreateObject("ADODB.Connection")
				conn.open Connection_String,"",""
				
				set rs= Server.CreateObject("ADODB.RecordSet")
				sql ="SELECT * FROM " & table_Name
				if request.Querystring("ID")<>"" then
					sql = sql & " WHERE " & ID_Field & "=" & cIntero(request.Querystring("ID"))
				end if 
				rs.open sql, conn, adOpenStatic, adLockOptimistic
				
				if Update_Record="INSERT" then
					'Inserimento nuovo record
					rs.addNew
					' gestisce l'eventuale chiave esterna
					' Public External_Key		'Nome delle chiave esterna (es. ID_AZIENDA su B2B)
					' public External_Key_Value ' Valore da preimpostare per la chiave esterna
					if External_Key<>"" and External_Key_Value<>"" then
						rs(External_Key) = External_Key_Value
					end if
				end if
				
				'imposta nome file
				rs(File_Field) = upl.FileName
				rs.Update
				
				rs.close
				conn.close
			end if
		
	end if
	
end Sub

end class

'---------------------------------------------------------------------------------------------------------------------
'-- Astrazione oggetto upload
'---------------------------------------------------------------------------------------------------------------------
'classe che astrae l'oggetto dalla vera implementazione
class UploadObject
	public Obj				'oggetto upload
	public fso				'oggetto filesystem per controlli
	public ObjectType		'tipo di oggetto istanziato
	public isEmpty 			'True se almeno un file è stato caricato
	public NumOfFile		'Numero dei file caricati

'---------------------------------------------------------------------------------------------------------------------
'-- Creazione oggetti	
'---------------------------------------------------------------------------------------------------------------------
	'crea l'istanza del primo oggetto upload valido
	Private Sub Class_Initialize()
		dim objVersion
		ObjectType = ""
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		On error resume next
		
		'....................................................... OGGETTO PERSIST UPLOAD
		if ObjectType= "" then		'oggetto precedrente non creato
			err.clear
			Set Obj = Server.CreateObject("Persits.Upload.1")
			if err.number = 0 then
				'oggetto creato correttamente
				ObjectType = OBJ_PERSITS
				Obj.RegKey = "74412-43017-14228"
				'salva temporaneamente il file in memoria per accedere al resto del form
				NumOfFile = obj.Save
				isEmpty = NumOfFile <1
			end if
		end if
		
		'....................................................... OGGETTO SOFTARTISAN FILEUP
		if ObjectType= "" then		'oggetto precedrente non creato
			err.clear
			Set Obj = Server.CreateObject("SoftArtisans.FileUp")
			if err.number = 0 then
				'verifica se l'oggetto in uso supporta le nuove funzionalita' del SAfileUP.
				err.clear 
				objVersion = obj.version
				if err.number <> 0 or cInt(left(objVersion, 1))<4 then
					ObjectType = OBJ_SOFTARTISANS_OLD
				else
					ObjectType = OBJ_SOFTARTISANS
				end if
				isEmpty = obj.isEmpty
				NumOfFile = obj.item.count ' da verificare se funziona
				obj.MaxBytes = 0
			end if
		end if
		
		response.write "<!-- " & ObjectType & " -->"
		
		On error goto 0
	end sub
	
	private sub Class_Terminate()
		set fso = nothing
		set Obj = nothing
	end sub
	
'---------------------------------------------------------------------------------------------------------------------
'-- Dichiarazione metodi
'---------------------------------------------------------------------------------------------------------------------
	public Sub SaveFile(Path,n)
		Path = replace(path, "\\", "\")

		Select case ObjectType
			 case OBJ_PERSITS
			 	if right(Trim(path), 1) <> "\" then
					Path = Path & "\"
				end if
				obj.Files(n).SaveAs(Path & obj.Files(n).ExtractFileName)
			case OBJ_SOFTARTISANS
				obj.path = Path
				obj.Save
			case OBJ_SOFTARTISANS_OLD
				obj.SaveAs(Path & FileName())
		end select
	end sub

	' Salva il primo
	public Sub Save(Path)
		CALL SaveFile(Path,1)
	end sub
	
	public function FileExists(Path)
		Path = replace(path, "\\", "\")
		FileExists = fso.FileExists(path & FileName()) 
	end function
	
	public function ExistsFileNo(Path, n)
		Path = replace(path, "\\", "\")
		FileExists = fso.FileExists(path & NameOfFile(n)) 
	end function
	
	public function SizeOfFile( n )
		Select case ObjectType
			 case OBJ_PERSITS
				SizeOfFile = obj.Files(n).Size
			case OBJ_SOFTARTISANS
				SizeOfFile = obj.Form(n).TotalBytes
			case OBJ_SOFTARTISANS_OLD
				SizeOfFile = obj.TotalBytes
		end select
	end function
	
	public function NameOfFile( n )
		Select case ObjectType
			 case OBJ_PERSITS
			 	NameOfFile = obj.Files(n).ExtractFileName
			case OBJ_SOFTARTISANS
				if n=1 then
					NameOfFile = obj.Form("INPUT_FILE").ShortFilename
				else
					NameOfFile = obj.Form("INPUT_FILE_" & n).ShortFilename
				end if
			case OBJ_SOFTARTISANS_OLD
				NameOfFile = ExtractFileName(obj.UserFileName)
		end select
	end function
	
'---------------------------------------------------------------------------------------------------------------------
'-- Dichiarazione proprieta'
'---------------------------------------------------------------------------------------------------------------------
		
	'imposta sovrascrittura dei file
	property Let OverWrite(ByVal Flag)
		Select case ObjectType
			 case OBJ_PERSITS
				obj.OverWriteFiles = Flag
			case OBJ_SOFTARTISANS
				obj.OverWriteFiles = Flag
			case OBJ_SOFTARTISANS_OLD
				obj.OverWriteFiles = Flag
		end select
	end property
	
	'restituisce la dimensione del file
	property Get FileSize()
		FileSize = SizeOfFile( 1 )
	end property
	
	
	
	'restituisce il nome del file
	property Get FileName()
		Select case ObjectType
			 case OBJ_PERSITS
			 	FileName = obj.Files(1).ExtractFileName
			case OBJ_SOFTARTISANS
				FileName = obj.Form("INPUT_FILE").ShortFilename
			case OBJ_SOFTARTISANS_OLD
				FileName = ExtractFileName(obj.UserFileName)
		end select
	end property
	
	'restituisce il valore del dell'elemento del form richiesto
	property Get Form(element)
		Select case ObjectType
			 case OBJ_PERSITS
			 	Form = obj.Form(element).value
			case OBJ_SOFTARTISANS
				Form = obj.Form(element)
			case OBJ_SOFTARTISANS_OLD
				Form = obj.Form(element)
		end select
	end property


'---------------------------------------------------------------------------------------------------------------------
'-- Dichiarazione metodi / funzioni private
'---------------------------------------------------------------------------------------------------------------------
	private function ExtractFileName(LongFileName)
		ExtractFileName = mid(LongFileName, (instrrev(longFileName, "\", -1,vbTextCompare)) + 1, len(LongFileName))
	end function
end class

%>