<% 
'...................................................................................................
'classe per la gestione del timer per la misura dei tempi di esecuzione di una pagina
'...................................................................................................
class Crono

	private times
	private StartTime
	private PreviousTime

	
	Private Sub Class_Initialize()
		StartTime = Timer
		PreviousTime = Timer
		times = ""
	End Sub
	
	Private Sub Class_Terminate()
		Trace("END PAGE")
		times = "<table border=""1"">" & vbCrLf & _
				"<caption>TEMPO ESCUZIONE PAGINA IN SECONDI</caption>" & vbCrLf & _
				"<colgroup align=""left"">" & vbCrLf & _
				"<colgroup align=""right"">" & vbCrLf & _
				"<colgroup align=""right"">" & vbCrLf & _
				"<colgroup align=""right"">" & vbCrLf & _
				"<tr><th>AZIONE</th>" & _
					"<th>da inizio pagina (millesimi)</th>" & _
					"<th>tempo azione (millesimi)</th>" & _
					"<th>timer (secondi)</th></tr>" & vbCrlf & _
				"<tr><td>START PAGE</td><td>0</td><td>0</td><td>" & FormatPrice(StartTime, 2, True) & "</td></tr>" & vbCrLF & _
				times & _
				"</table><br><br><br><br><br>" & vbCrLf
		response.write times
	End Sub
	
	Public Sub Trace(label)
		dim NowTime
		NowTime = timer
		times = times & "<tr><td>" & label & "</td>" & _
						"<td>" & FormatPrice(((NowTime - StartTime)*1000), 2, true) & "</td>" & _
						"<td>" & FormatPrice(((NowTime - PreviousTime)*1000), 2, true) & "</td>" & _
						"<td>" & FormatPrice(NowTime, 2, true) & "</td></tr>" & vbCrLF
		PreviousTime = NowTime
	end sub

end class


'visualizza il contenuto della richiesta (form e querystring) in una tabella HTML
function ListRequest()
	dim var%>
	<!--
	QUERYSTRING
	<%for each var in request.querystring%>
		<%= var %>:"<%= request.querystring(var) %>"
	<% next %>
	FORM
	<%for each var in request.form%>
		<%= var %>:"<%= request.form(var) %>"
	<% next %>
	-->
	<table cellspacing="0" cellpadding="3" border="1">
		<tr><th colspan="2">QUERYSTRING</th></tr>
		<%for each var in request.querystring%>
			<tr>
				<td bgcolor="#FFFFFF"><%= var %></td>
				<td bgcolor="#FFFFFF"><%= request.querystring(var) %></td>
			</tr>
		<% next %>
		<tr><th colspan="2">FORM</th></tr>
		<%for each var in request.form%>
			<tr>
				<td bgcolor="#FFFFFF"><%= var %></td>
				<td bgcolor="#FFFFFF"><%= request.form(var) %></td>
			</tr>
		<% next %>
	</table>
	<%
end function


'visualizza il contenuto delle servervariables in una tabella HTML
function ListServerVariables()
	dim var%>
	<table cellspacing="0" cellpadding="3" border="1">
		<tr><th colspan="2">request.ServerVariables</th></tr>
		<%for each var in request.ServerVariables%>
			<tr>
				<td bgcolor="#FFFFFF"><%= var %></td>
				<td bgcolor="#FFFFFF"><%= request.ServerVariables(var) %></td>
			</tr>
		<% next %>
	</table>
<%end function


'visualizza il contenuto dei cookies in una tabella HTML
function ListCookies()
	dim var, var2%>
	<table cellspacing="0" cellpadding="3" border="1">
		<tr><th colspan="2">request.Cookies</th></tr>
		<%for each var in request.Cookies%>
			<tr>
				<td rowspan="2" bgcolor="#FFFFFF"><%= var %></td>
				<td bgcolor="#FFFFFF"><%= request.Cookies(var) %></td>
			</tr>
			<tr>
				<td bgcolor="#00ffff">
					<table cellspacing="0" cellpadding="1" border="1">
						<%for each var2 in request.Cookies(var) %>
							<tr>
								<td><%= var2 %></td>
								<td><%= request.Cookies(var)(var2) %></td>
							</tr>
						<% next %>
					</table>
				</td>
			</tr>
		<% next %>
	</table>
<%end function


'visualizza il contenuto della sessione in una tabella HTML
function ListSession()
	dim var%>
	<!--
	<%for each var in Session.contents
		if not isObject(Session(var)) AND not isArray(Session(var)) then %>
			<%= var %>:"<%= Session(var) %>"
		<% elseif isObject(Session(var)) then%>
			<%= var %>: Object (<%= TypeName(Session(var)) %>)
		<% elseif isArray(Session(var)) then%>
			<%= var %>: Array
		<% end if %>
	<% next %>
	FORM
	<%for each var in request.form%>
		<%= var %>:"<%= request.form(var) %>"
	<% next %>
	-->
	<table cellspacing="0" cellpadding="3" border="1">
		<tr><th colspan="2">Session</th></tr>
		<%for each var in Session.contents%>
			<tr>
				<% if isObject(Session(var)) then %>
					<td bgcolor="#DCFFB9"><%= var %></td>
					<td bgcolor="#DCFFB9">Object (<%= TypeName(Session(var)) %>)</td>
				<% elseif isArray(Session(var)) then %>
					<td bgcolor="#00ffff"><%= var %></td>
					<td bgcolor="#00ffff"><% ListArray(Session(var)) %></td>
				<%else %>
					<td bgcolor="#FFFFFF"><%= var %></td>
					<td bgcolor="#FFFFFF"><%= Session(var) %></td>
				<% end if %>
			</tr>
		<% next %>
	</table>
<%end function


'visualizza il contenuto della riga in una tabella html semplice
function ListDictionary(dict,VerticalDir)
	dim key%>
	<table cellspacing="0" cellpadding="3" border="1">
		<tr>
			<%
				response.write TypeName(dict)
				for each Key in dict.keys%>
				<% if instr(1,TypeName(dict(Key)),"Dict",vbTextCompare)>0 then %>
					<td bgcolor="#DCFFB9" title="<%= TypeName(dict(key)) %>"><%= Key %></td>
					<td bgcolor="#DCFFB9"><% ListDictionary(dict(key)) %></td>
				<% elseif isArray(dict(key)) then %>
					<td bgcolor="#00ffff" title="array"><%= Key %></td>
					<td bgcolor="#00ffff"><% ListArray(dict(key)) %></td>
				<%else %>
					<td bgcolor="#FFFFFF" title="semplice">key: <%= Key %></td>
					<td bgcolor="#FFFFFF">value: <%= dict(key) %></td>
				<% end if %>
				<% if VerticalDir then %>
					</tr><tr>
				<% end if %>
			<% next %>
		</tr>
	</table>
<%end function

'visualizza il contenuto dell'array in una tabella html semplice
function ListArray(a)
	dim i
	if isArray(a) then%>
		<table cellspacing="0" cellpadding="4" border="1">
			<% for i = lbound(a) to ubound(a) %>
				<tr>
					<% if isObject(a(i)) then %>
						<td bgcolor="#DCFFB9"><strong><%= i %></strong></td>
						<td bgcolor="#DCFFB9"><% ListDictionary(a(i)) %></td>
					<% elseif isArray(a(i)) then %>
						<td bgcolor="#00ffff"><strong><%= i %></strong></td>
						<td bgcolor="#00ffff"><% ListArray(a(i)) %></td>
					<%else %>
						<td bgcolor="#FFFFFF"><strong><%= i %></strong></td>
						<td bgcolor="#FFFFFF"><%= a(i) %></td>
					<% end if %>
				</tr>
			<% next %>
		</table>
	<%end if
end function


'visualizza il contenuto della query
function ListQuery(conn, sql)
	CALL ListRecordset(conn.execute(sql), true)
end function


'visualizza il contenuto del recordset in una tabella html semplice
'se ListAll = true lo visualizza fino alla fine, altrimenti solo il record corrente
function ListRecordset(rs, ListAll)
	dim field%>
	<table cellspacing="0" cellpadding="4" border="5">
		<tr>
			<td bgcolor="#ffffe0" colspan="<%= rs.fields.count * 2%>">
				<%= rs.source %>
			</td>
		</tr>
		<%do while not rs.eof %>
			<tr>
				<% for each field in rs.fields %>
					<td bgcolor="#FFFFFF"><%= field.name %></td>
					<td bgcolor="#FFFFFF">
						<% if field.type=203 then
							response.write "&lt;blob value&gt;"
						else
							response.write field.value
						end if %>
					</td>
				<% next %>
			</tr>
			<%if not ListAll then
				exit do
			end if
			rs.movenext
		loop
		if ListAll then
			rs.movefirst		
		end if
%>
	</table>
<%end function


'visualizza tutti i parametri ed i relativi valori di un command
function ListCommand(CMD) 
	dim parameter%>
	<table cellspacing="0" cellpadding="4" border="1">
		<tr><td colspan="3"><strong><%= CMD.CommandText %></strong></td></tr>
		<% for each parameter in CMD.Parameters %>
			<tr>
				<td><%= parameter.name %></td>
				<td><%= TypeName(parameter.type) %> (ADO:<%= parameter.type %>)</td>
				<td><%= parameter.value %></td>
			</tr>
		<% next %>	
	</table>
<%end function
%>