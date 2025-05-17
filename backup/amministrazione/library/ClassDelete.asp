<!--#INCLUDE FILE="class_testata.asp" -->
<!--#INCLUDE FILE="Tools4admin.asp" -->
<!--#INCLUDE FILE="Tools.asp" -->
<%

class OBJ_Delete

'parametri da impostare per ogni sezione
Public Message							'Messaggio visualizzato quando si carica il form di conferma
Public Note								'eventuali note di cancellazione da visualizzare nel form
Public Name_Field						'campo o espressione SQL per nome record
Public ID_Field							'Campo ID record
Public Table							'Nome tabella
Public IndexTable						'Nome tabella per l'indice
Public Caption							'titolo sezione
Public MsgSql							'sql di lettuta del record (CALCOLATO AUTOMATICAMENTE IN DELETE CONFIRM SE NON IMPOSTATO)
Public OnlyRelations					'se TRUE non esegue la delete ma solo le routine "Delete_Relazioni"

'parametri da impostare per sito
Public Section							'parametro con il nome della sezione da cancellare
Public ID_Value							'valore ID record
Public PageName							'Nome della pagina corrente
Public ReloadOpener						'Indica se alla fine della cancellazione deve essere fatto 
										'il reload dela finestra padre
Public ConnString						'stringa di connessione a database
Public LinkStyle						'stile dei link di conferma/chiusura della finestra
Public MessageStyle						'stile messaggio di testo
Public CaptionStyle						'stile titolo sezione
Public CaptionColor						'colore di sfondo del titolo di sezione
Public BorderDarkColor					'colore del bordo piu' scuro
Public BorderLightColor					'colore del bordo piu' chiaro
Public BackgroundColor					'Colore di sfondo cella principale
Public OperationOK						'Esito della cancellazione
Public DeleteRelations					'Indica se sono gestite anche le relazioni
Public AfterDelete						'Indica se dopo la cancellazione deve essere fatta qualche operazione

Private oIndex

Private Options_Labels
Private Options_Values
Private Options_Notes

Public conn, rs, sql

Private Sub Class_Initialize()
	OnlyRelations = FALSE
	Set Conn = Server.CreateObject("ADODB.Connection")
End Sub

private sub Class_Terminate()
	conn.close
	'set CheckList = nothing
	set conn = nothing
end sub

'gestione completa della cancellazione del record
public Sub Delete_Manager()
	OperationOK = False
	if request.Querystring("MODE")="CANC" then
		'esegue cancellazione
		Delete_Execute()
	else
		'chiede conferma
		Delete_Confirm()
	end if
end Sub

'disegna il form per richiesta di conferma cancellazione
public Sub Delete_Confirm()
	dim intestazione
	'disegna intestazione
	set intestazione= New testata
	intestazione.sezione = Caption & ChooseValueByAllLanguages(Session("LINGUA"), " - cancellazione", " - deleting", "", "", "", "", "", "")
	intestazione.scrivi_ridotta()
    
	if MsgSql="" then
		MsgSql = "SELECT (" & Name_Field & ") AS NOMINATIVO FROM " & table
	end if
	MsgSql = MsgSql + " WHERE " & ID_Field & "=" & cIntero(ID_Value)
	
'response.write MsgSql
	
	set rs = conn.execute(MsgSql)
	
	if not rs.eof then
		Message = replace(Message, "<RECORD>", "<b>""" & rs(0) & """</b>")
	else
		Message = ChooseValueByAllLanguages(Session("LINGUA"), "ERRORE NELL'APPLICAZIONE: Record non individuato.", "APPLICATION ERROR: Record not found.", "", "", "", "", "", "")
	end if%>
	<div id="confirm" style="position:absolute; top:90px; width:400px; z-index:4">
	<table cellpadding="1" cellspacing="0" style="border: 1px solid <%= BorderDarkColor %>" width="99%" align="center">
		<tr>
			<td bgcolor="<%= CaptionColor %>" style="padding-left:10px; border-bottom:1px solid <%= BorderDarkColor %>">
				<font <%= CaptionStyle %>>
					<%= Caption %>
				</font>
			</td>
		</tr>
		<tr bgcolor="<%= BackgroundColor %>">
			<td>
				<table cellpadding="4" cellspacing="0" width="100%" border="0">
					<% if not isEmpty(Options_Labels) then %>
						<form action="" method="get" id="delete" name="delete">
						<input type="hidden" name="MODE" value="CANC">
						<input type="hidden" name="ID" value="<%= ID_Value %>">
						<input type="hidden" name="SEZIONE" value="<%= Section %>">
					<% end if %>
					<tr>
						<td align="center" style="border-bottom:1px solid <%= BorderLightColor %>">
							<img src="<%= GetAmministrazionePath() %>grafica/alert_anim.gif">
						</td>
					</tr>
					<tr>
						<td align="left" <%= MessageStyle %>>
							<%= Message %>
							<% if not isEmpty(Options_Labels) then %>
                                <table cellpadding="0" cellspacing="0" width="100%">
                                    <% dim check
                                    for each check in Options_Labels %>
                                        <tr>
                                            <td style="width:6%; vertical-align:baseline;"><input type="checkbox" class="checkbox" name="<%= check %>" <%= chk(Options_values(check)) %>></td>
                                            <td style="padding-bottom:3px;">
                                                <%= Options_Labels(check) %>
                                                <% if Options_Notes(check)<>"" then %>
                                                    <span class="note"> ( <%= Options_Notes(check) %> ) </span>
                                                <% end if %>
                                            </td>
                                        </tr>
                                    <% next %>
                                </table>
							<% end if %>
						</td>
					</tr>
					<tr>
						<td style="height:25px;">
							<% if not rs.eof then 
								if not isEmpty(Options_Labels) then %>
									<table cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td nowrap style="width:45%; padding-right:10px; text-align:right;">	
												<input type="submit" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "CONFERMA", "OK", "", "", "", "", "", "")%>" class="button" name="conferma" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Conferma ed esegue la cancellazione dell'elemento", "Confirm deleting element", "", "", "", "", "", "")%>" style="width:70px;">
											</td>
											<td style="width:10%;"><font style="font:10px Arial;">&nbsp;</font></td>
											<td nowrap style="width:45%; padding-left:10px; text-align:left;">	
												<input type="button" value="<%= ChooseValueByAllLanguages(Session("LINGUA"), "ANNULLA", "CANCEL", "", "", "", "", "", "")%>" class="button" name="annulla" onclick="window.close();" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Chiude la finestra ed annulla la cancellazione", "Close the window and delete operation", "", "", "", "", "", "")%>" style="width:70px;">
											</td>
										</tr>
									</table>
								<% else 
									'mantiene la versione normale
									%>
									<table cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td nowrap style="width:45%; padding-right:10px; text-align:right;">
												<% if instr(1,PageName, "?",vbTextCompare)>0 then
													PageName = PageName & "&"
												else
													PageName = PageName & "?"
												end if %>
												<a accesskey="c" tabindex="1" href="<%= PageName %>SEZIONE=<%= Section %>&ID=<%= ID_Value %>&MODE=CANC" <%= LinkStyle %>
												   title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Conferma ed esegue la cancellazione dell'elemento", "Confirm deleting element", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %> id="primo_elemento">
													<%= ChooseValueByAllLanguages(Session("LINGUA"), "CONFERMA", "OK", "", "", "", "", "", "")%>
												</a>
											</td>
											<td style="width:10%;"><font style="font:10px Arial;">&nbsp;</font></td>
											<td nowrap style="width:45%; padding-left:10px; text-align:left;">
												<a accesskey="a" tabindex="2" href="#"  onclick="window.close();" <%= LinkStyle %>
												   title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Chiude la finestra ed annulla la cancellazione", "Close the window and delete operation", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
													<%= ChooseValueByAllLanguages(Session("LINGUA"), "ANNULLA", "CANCEL", "", "", "", "", "", "")%>
												</a>
											</td>
										</tr>
									</table>
								<% end if
							else %>
								<table cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td nowrap style="width:45%; padding-left:10px; text-align:left;">
											<a href="#" onclick="window.close();" <%= LinkStyle %>
											   title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Chiude la finestra ed annulla la cancellazione", "Close the window and delete operation", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
												<%= ChooseValueByAllLanguages(Session("LINGUA"), "ANNULLA", "CANCEL", "", "", "", "", "", "")%>
											</a>
										</td>
									</tr>
								</table>
							<% end if %>
						</td>
					</tr>
					<% if Note<>"" then %>
						<tr>
							<td class="note"><%= Note %></td>
						</tr>
					<% end if %>
					<% if not isEmpty(Options_Labels) then %>
						</form>
					<% end if %>
				</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= CaptionColor %>" style="text-align:right; padding-left:10px; border-top:1px solid <%= BorderDarkColor %>">
				<a href="#" onclick="window.close();" <%= LinkStyle %> title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "chiudi questa finestra", "close this window", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
					<%= ChooseValueByAllLanguages(Session("LINGUA"), "CHIUDI", "CLOSE", "", "", "", "", "", "")%>
				</a>
			</td>
		</tr>
	</table>
	</div>
	<br>
	<script language="JavaScript" type="text/javascript">
		FitWindowSize_delete(this);
		
		PageOnLoad_FocusSet_delete();
		
        //copia delle funzioni presenti in utils.js
		function FitWindowSize_delete(obj){
			var displace = 22;
			//calcola offset per ridimensionare la finestra
			var width_offset = obj.document.body.scrollWidth - obj.document.body.clientWidth
			var height_offset = displace + obj.document.body.scrollHeight - obj.document.body.clientHeight
				
			//verifica se la finestra sfora dallo schermo in altezza
			if ((obj.screenTop + obj.document.body.clientHeight + height_offset) > obj.screen.availHeight){
				//la finestra sfora in altezza dallo schermo
				var top_offset = obj.screen.availHeight - displace - (obj.screenTop + obj.document.body.clientHeight + height_offset);
				//sposta la finestra in alto per mantenere spazio sufficente
				obj.moveBy(0, top_offset);
			}
			//ridimensiona la finestra
			obj.resizeBy(width_offset, height_offset);
		}
		
		function PageOnLoad_FocusSet_delete(){
			window.onload = PageOnLoad_Focus_delete;
		}
		
		function PageOnLoad_Focus_delete(){
			window.focus();
			var elemento = document.getElementById("primo_elemento");
            if (elemento){
				elemento.focus();
			}
		}
	</script>
	<%
	rs.close
	conn.close
	Set rs = nothing
end Sub

'Esegue la cancellazione e, se necessario fa il reload finestra padre e chiude finestra corrente.
public Sub Delete_Execute()
	Conn.beginTrans
	
	'on error goto 0
	
	if DeleteRelations OR onlyRelations then
	
        response.write ChooseValueByAllLanguages(Session("LINGUA"), "cancellazione relazioni in corso...<br>", "deleting relations...<br>", "", "", "", "", "", "")
	
        'relazioni del record
		CALL Delete_Relazioni(conn, ID_Value)
	end if
	
	'gestione cancellazione eventuale contenuto ed indice
	if IsObject(oIndex) AND instr(1, table, "tb_contents_index", vbTextCompare)<1 then

		response.write ChooseValueByAllLanguages(Session("LINGUA"), "cancellazione indice in corso...<br>", "deleting index...<br>", "", "", "", "", "", "")
		if cString(IndexTable) = "" then
			IndexTable = table
		end if
		CALL oIndex.content.DeleteAll(IndexTable, id_value)
	end if
    
	if NOT onlyRelations then
		
        response.write ChooseValueByAllLanguages(Session("LINGUA"), "cancellazione in corso...<br>", "deleting...<br>", "", "", "", "", "", "")
        
		sql = "DELETE FROM " & table & " WHERE " & ID_Field & "=" & cIntero(ID_Value)
		Conn.execute(sql)
	end if
    
	if AfterDelete then
		'operazioni successive all'aggiornamento
        response.write ChooseValueByAllLanguages(Session("LINGUA"), "cancellazione dati collegati in corso...<br>", "deleting linked data...<br>", "", "", "", "", "", "")
		CALL Operations_AfterDelete(conn, ID_Value)
	end if
    
	'esegue il reload della pagina padre
	response.write "<script language=""JavaScript"">" & vbCrLf
	if ReloadOpener then
		response.write "	opener.location.reload(true);" & vbCrLf
	end if
	response.write "	window.close();" & vbCrLf
	response.write "</script>" & vbCrLf
	
	OperationOK = TRUE
	Conn.CommitTrans

end sub

'******************************************************
'Stringa di connessione al database
Public Property Get ConnectionString()
    ConnectionString = ConnString
End Property

Public Property Let ConnectionString(ByVal vData)
	ConnString = vData
	Conn.open ConnString, "", ""
End Property


'******************************************************
'Oggetto di gestione dell'indice
Public Property Get Index()
    set Index = oIndex
End Property

Public Property Let Index(obj)
    set oIndex = obj
    set conn = oIndex.conn
End Property


'metodo per l'aggiunta di opzioni per la cancellazione del record (lista di checkbox in conferma)
Public Sub AddOption(check_name, check_label, check_DefaultValue, check_note)
	if isEmpty(Options_Labels) then
		set Options_Labels = Server.CreateObject("Scripting.Dictionary")
        set Options_Values = Server.CreateObject("Scripting.Dictionary")
        set Options_Notes = Server.CreateObject("Scripting.Dictionary")
	end if

    Options_Labels.Add check_name, cString(check_label)
    Options_Values.Add check_name, cBool(check_DefaultValue)
    Options_Notes.Add check_name, cString(check_note)
end sub


end class
%>
