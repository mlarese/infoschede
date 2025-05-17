<% 
'*******************************************************************************
'classe per la gestione dell'ordinamento
'*******************************************************************************


class OrderSelector
	
	'variabili interne
	private Sqls
	private Labels
	private DefaultOrder
	private CurrentOrder
	
	private SessionVarName
	private SessionVarPrefix

	Private Sub Class_Initialize()
		set SQLs = Server.CreateObject("scripting.dictionary")
		set Labels = Server.CreateObject("scripting.dictionary")
		DefaultOrder = ""
		CurrentOrder = ""
		
		'costruzione nome variabile di sessione che mantiene l'ordine corrente
		if request("PAGINA")<>"" then
			SessionVarName = "ORDER_" & cIntero(request("PAGINA"))
		else
			SessionVarName = GetPageName()
			SessionVarName = "ORDER_" & left(SessionVarName, instr(1, SessionVarName, ".", vbTextCompare)-1)
		end if
		if Session("OrderBy_VarPrefix")<>"" then
			SessionVarName = Session("OrderBy_VarPrefix") & SessionVarName
			SessionVarPrefix = Session("OrderBy_VarPrefix") & SessionVarName
			Session.Contents.Remove("OrderBy_VarPrefix")
		end if
		
		if request.querystring("order_by") <> "" then
			CurrentOrder = request.querystring("order_by")
		end if
	end Sub
	
	Private Sub Class_Terminate()
		set Sqls = nothing
		set Labels = nothing
	End Sub
	
	
'DEFINIZIONE METODI:*********************

	'aggiunge una nuova opzione di ordinamento
	Public Function AddOrder(key, label, sql, isDefault)
		if Labels.Exists(key) then
			AddOrder = false
			Labels(key) = Label
			Sqls(key) = sql
		else
			AddOrder = true
			CALL Labels.Add(key, label)
			CALL Sqls.Add(key, sql)
		end if
		if isDefault then
			DefaultOrder = Key
			CALL SetCurrentOrder()
		end if
	end function
	
	
	'azzera variabili e gestione dei dati 
	Public Sub Reset()
		Session(SessionVarName) = DefaultOrder
		CurrentOrder = DefaultOrder
	end sub
	
	'genera drop down per scelta ordinamento
	Public Sub DropDown(lingua)
		%>
		<script language="JavaScript" type="text/javascript">
			function select_order_OnChange(obj){
				document.location = "?<%= GetQueryString() %>" + obj.value;
			}
		</script>
		<%CALL DropDownDictionary(Labels, "select_order", CurrentOrder, true, " class=""order"" onchange=""select_order_OnChange(this)""", lingua)
	end sub
	
	Public sub DrawTable(lingua, DistinguiCelle)
		dim key, QueryString
		if Labels.Count > 0 then
			QueryString = GetQueryString()%>
			<table cellpadding="0" cellspacing="0" class="orderby">
				<tr>
					<td class="label"><%= ChooseValueByAllLanguages(lingua, "ordina per", "order by", "Auftrag für", "Commande pour", "Orden para", "заказ", "命令", "ordem") %>:</td>
					<% for each key in Labels.keys
						if instr(1, CurrentOrder, Key, vbTextCompare)>0 then %>
							<td class="<%= IIF(DistinguiCelle, IIF(instr(1, Key, "asc", vbTextCompare), "asc_", "desc_"), "") %>selected" title="<%= ChooseValueByAllLanguages(lingua, "elenco ordinato per", "list ordered for", "Liste auftrag für", "liste command&eacute;e pour", "lista ordenada para", "Список отсортирован", "列表进行排序", "lista ordenada") %>: <%= Labels(key) %>"><%= Labels(key) %></td>
						<% else %>
							<td <%= IIF(DistinguiCelle, "class=""" & IIF(instr(1, Key, "asc", vbTextCompare), "asc", "desc") & """", "") %>><a href="?<%= QueryString %><%= Key %>" title="<%= ChooseValueByAllLanguages(lingua, "ordina per", "order by", "Auftrag für", "Commande pour", "Orden para", "заказ", "命令", "ordem") %>: <%= Labels(key) %>"><%= Labels(key) %></a></td>
						<% end if
					next %>
				</tr>
			</table>
		<%end if
	end sub
	
	
'DEFINIZIONE PROPRIETA':*********************
	
	'restituisce la clausola order by corrispondente all'ordinmanto corrente
	Public Property Get SQL()
		CALL setCurrentOrder()
		if not Sqls.Exists(CurrentOrder) then
			CurrentOrder = DefaultOrder
			SetCurrentOrder()
		end if
		SQL = " ORDER BY " & Sqls(CurrentOrder)
	end Property
	
	'restituisce il nome dell'ordinamento corrente
	Public Property Get Label()
		CALL setCurrentOrder()
		Label = Labels(CurrentOrder)
	end Property
	
	'restituisce il codice dell'ordinamento corrente
	Public Property Get Order()
		Order = CurrentOrder
	end Property
	
	'restituisce il nome della variabile di sessione utilizzata per registrare l'ordinamento
	Public Property Get SessionVariableName()
		SessionVariableName = SessionVarName
	End Property
	
	'cambia il nome della variabile di sessione che registra l'ordinamento scelto preservandone il valore
	Public Property Let SessionVariableName(name)
		Session(name) = Session(SessionVarName)
		Session.Contents.Remove(SessionVarName)
		SessionVarName = name
	End Property
	
	
'DEFINIZIONE FUNZIONI PRIVATE:*********************

	Private Sub SetCurrentOrder()
		if CurrentOrder = "" then
			if Session(SessionVarName)="" then
				CurrentOrder = DefaultOrder
			else
				CurrentOrder = Session(SessionVarName)
			end if
		end if
	end sub
	
	
	Private function GetQueryString()
		GetQueryString = request.ServerVariables("QUERY_STRING")
		if instr(1, GetQueryString, "order_by=", vbTextCompare)>0 then
			GetQueryString = replace(GetQueryString, "&order_by=" & request.querystring("order_by"), "")
			GetQueryString = replace(GetQueryString, "order_by=" & request.querystring("order_by") & "&", "")
			GetQueryString = replace(GetQueryString, "order_by=" & request.querystring("order_by"), "")
		end if
		GetQueryString = GetQueryString + IIF(GetQueryString <> "", "&", "") +"order_by="
	end function
	
end class

%>