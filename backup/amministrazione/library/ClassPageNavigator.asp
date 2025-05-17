<% 
'*******************************************************************************
'classe per la gestione della paginazione con memoria
'*******************************************************************************
'......... METODI PER LA GESTIONE DELLA PAGINAZIONE CON MEMORIA:
'Reset()								:porta la visualizzazzione e la memoria alla prima pagina
'_Render(PageCount, _
'		 TextStyle, _
'		 PageLinkStyle, _
'		 PageSelectedStyle)				:visualizza la paginazione

'DESCRIZIONE PRORIETA'*******************
'PageNo									: numero della pagina corrente nella paginazione
'ParentOnfiguration						: oggetto configurazione per prendere i dati della paginazione(opzionale)
										' se senza oggetto calcola il link con il nome della pagina corrente se con indirizzo relativo


class PageNavigator
	'variabili interne
	Private SessionVarName
	Private Param
	Private ParentConfigObject
	Private BaseUrl
	Private PageName
	
	'variabili utilizzate nella gestione dei recordset avanzata per gruppo di navigazione
	Public PageCount
	Public RecordCount
	Public CommandTimeout
	
	Public QueryString
	Public PageNo
	Public Lingua
	Public offset
	Public LinkSelChar ' contiene la maschera da usare per il link selezionato es. [n]
	
	'inizializza la navigazione per pagina 
	Private Sub Class_Initialize()

		'costruzione nome variabile di sessione che mantiene la paginazione
		if request("PAGINA")<>"" then
			SessionVarName = "@page_" & cIntero(request("PAGINA"))
			QueryString = "&PAGINA=" & cIntero(request("PAGINA"))
		else
			SessionVarName = GetPageName()
			SessionVarName = "@page_" & left(SessionVarName, instr(1, SessionVarName, ".", vbTextCompare)-1)
			QueryString = ""
		end if
		if Session("PageNavigator_VarPrefix")<>"" then
			SessionVarName = Session("PageNavigator_VarPrefix") & SessionVarName
			Session.Contents.Remove("PageNavigator_VarPrefix")
		end if

		'ciclo che mantiene le altre variabili passate per querystring
		for each param in request.querystring
			if UCase(param) <> "PAGINA" AND UCase(param) <> "PAGE_NO" AND UCase(param) <> "OFFSET" then
				QueryString = Querystring & "&"& param &"="& request.querystring(param)
			end if
		next
		
		PageNo = cInteger(Request("page_no"))
		offset = cInteger(Request("offset"))
		LinkSelChar = ""
		if PageNo = 0 Then
			PageNo = cInteger(Session(SessionVarName))
			if PageNo = 0 Then
				PageNo = 1
			end if
		end if
		
		Session(SessionVarName) = PageNo
		
		Lingua = Session("LINGUA")
		
		BaseUrl = ""
		PageName = GetPageName()
	end sub
	

'DEFINIZIONE METODI:*********************
	
	'resetta i dati della navigazione corrente e riporta in prima pagina
	Public Function Reset()
		
		PageNo = 1
		offset = 1
		Session(SessionVarName) = PageNo
		
	end function
	
	'resetta i dati di tutte le pagine e di tutte le sezioni (Non riconosce il parametro Session("PageNavigator_VarPrefix"))
	Public Function ResetALL()
		dim VarName
		
		'resetta tutte le variabili di sessione
		for each VarName in Session.Contents
			if instr(1, VarName, "@page_", vbTextCompare)>0 then
				'variabile di sessione del pager
				Session(VarName) = 1
			end if
		next
		
		Reset
		
	end function
	
	
	
	'disegna la paginazione di tutte le pagine
	Public Function Render_PageNavigator(PageCount, TextStyle, PageLinkStyle, PageSelectedStyle)
		dim i, LabelPages, LabelGotoPage, LabelCurrentPage
		
		LabelPages = ChooseValueByAllLanguages(lingua, "Pagine", "Pages", "Seiten", "Pages", "P&aacute;ginas", "Страницы", "页", "Páginas")
		LabelCurrentPage = ChooseValueByAllLanguages(lingua, "Pagina <n> di", "Page <n> of", "Seite <n> von", "Page <n> de", "P&aacute;gina <n> de", "страница", "<n>的页的", "<n> da página") & " " & PageCount
		LabelGotoPage = ChooseValueByAllLanguages(lingua, "Vai alla pagina <n> di", "Go to page <n> of", "Gehen Sie zur Seite <n> von", "Allez à la page <n> de", "Vaya a la página <n> de", "Перейти на страницу из <n>", "转到第页<n>的对", "Ir para a página de <n>") & " " & PageCount
		
		if TextStyle<>"" then
			TextStyle = "class=""" & TextStyle & """ "
		end if
		if PageLinkStyle<>"" then
			PageLinkStyle = "class=""" & PageLinkStyle & """ "
		end if
		if PageSelectedStyle<>"" then
			PageSelectedStyle = "class=""" & PageSelectedStyle & """ "
		end if%>
		<table border="0" cellspacing="0" cellpadding="1">
			<tr>
				<td <%= TextStyle %> valign="top"><%= LabelPages %>:</td>
				<td <%= TextStyle %> valign="top">
					<% for i = 1 to PageCount %>
						<span >
							<span style="font-size:1px">
								&nbsp;
							</span>
							<%if i=PageNo then%>
								<a <%= PageSelectedStyle %> title="<%= replace(LabelCurrentPage, "<n>", i) %>" <%= ACTIVE_STATUS %>>
								<% if LinkSelChar<>"" then %>
									<%= replace(LinkSelChar,"n",cstr(i)) %>
								<% else %>
									<%= i %>
								<% end if %>
								</a>
							<% else %>
								<a <%= PageLinkStyle %> href="<%= BaseURL %><%= PageName %>?page_no=<%= i %><%= Querystring %>" title="<%= replace(LabelGotoPage, "<n>", i) %>" <%= ACTIVE_STATUS %>>
									<%= i %></a>
							<% end if%>
						</span>
					<%next %>
				</td>
			</tr>
		</table>
	<%end function
	
	
	'disegna la paginazione di tutte le pagine elencate per gruppi
	Public Function Render_GroupNavigator(PageForGroup, PageCount, TextStyle, PageLinkStyle, PageSelectedStyle)
		dim displace, i
		
		if TextStyle<>"" then
			TextStyle = "class=""" & TextStyle & """ "
		end if
		if PageLinkStyle<>"" then
			PageLinkStyle = "class=""" & PageLinkStyle & """ "
		end if
		if PageSelectedStyle<>"" then
			PageSelectedStyle = "class=""" & PageSelectedStyle & """ "
		end if
		
		'recupera indice gruppo corrente di pagine
		if offset = 0 and PageNo <= PageForGroup then 
			offset = 1
		elseif offset=0 OR offset > PageCount then
			offset = ((PageNo \ PageForGroup) * PageForGroup) + 1
		end if

		'recupera indice di partenza della pagina
		if (PageCount - offset) > PageForGroup then 
			displace = offset + (PageForGroup - 1)
		else
			displace = PageCount
		end if%>
		<table border="0" cellspacing="0" cellpadding="0" style="width:100%;">
			<tr>
				<td <%= TextStyle %> valign="top" style="width: 8%; padding-right:4px; text-align:right;"><%= ChooseValueByAllLanguages(lingua, "Pagine", "Pages", "Seiten", "Pages", "P&aacute;ginas", "Страницы", "页", "Páginas") %>:</td>
				<td>
					<%if offset>1 then	'blocco precedente di pagine
					%>
						<span style="white-space:nowrap">
							<span style="font-size:1px">
								&nbsp;
							</span>
							<a <%= PageLinkStyle %> href="<%= BaseURL %><%= PageName %>?page_no=<%= offset - PageForGroup %>&offset=<%=offset - PageForGroup%><%= QueryString %>" 
							   title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Vai alle " & PageForGroup & " pagine precedenti", "Go to the " & PageForGroup & " previous pages", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
								&lt;&lt;
							</a>
						</span>
					<%end if 
					for i=offset to displace%>
						<span style="white-space:nowrap">
							<span style="font-size:1px">
								&nbsp;
							</span>
							<%if i=PageNo then%>
								<a <%= PageSelectedStyle %> title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "pagina corrente", "present page", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
									<% if LinkSelChar<>"" then %>
									<%= replace(LinkSelChar,"n",cstr(i)) %>
									<% else %>
									<%= i %>
								<% end if %>
								</a>
								
							<% else %>
								<a <%= PageLinkStyle %> href="<%= BaseURL %><%= PageName %>?page_no=<%= i %>&offset=<%=offset%><%= Querystring %>" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "vai alla pagina n&ordm; ", "go to page n&ordm; ", "", "", "", "", "", "")%><%= i %>" <%= ACTIVE_STATUS %>>
									<%= i %>
								</a>
							<% end if%>
						</span>
					<%next
					if i <= PageCount then	'blocco successivo di pagine
					%>
						<span style="white-space:nowrap">
							<span style="font-size:1px">
								&nbsp;
							</span>
							<a <%= PageLinkStyle %> href="<%= BaseURL %><%= PageName %>?page_no=<%= i %>&offset=<%=offset + PageForGroup%><%= QueryString %>" 
							   title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Vai alle " & PageForGroup & " pagine successive", "Go to the " & PageForGroup & " following pages", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
								&gt;&gt;
							</a>
						</span>
					<%end if%>
				</td>
			</tr>
		</table>
	<%end function
	
	
	'procedura che apre il recordset con la query richiesta ottimizzando il numero di record restituiti
	Public Sub OpenSmartRecordset(conn, rs, byval sql, byval PageSize)	
		dim i, sqlCount
		
		if cIntero(CommandTimeout)>0 then
			conn.CommandTimeout = CommandTimeOut
		end if
		
		'ritorna il conteggio dei record	
		'esegue il replace della prima parte del from
		'ATTENZIONE: solo fino alla prima istanza
		'response.write sql
		'response.end
		i = instr(1, sql, "FROM", vbTextCompare)		
		if (instr(7, sql, "SELECT", vbTextCompare)>i OR instr(7, sql, "SELECT", vbTextCompare)<1) AND instr(1, sql, "UNION", vbTextCompare)<1 then
			sqlCount = "SELECT COUNT(*) " + _
				 	   right(sql, len(sql) - (i-1))
			
			'rimuove la parte order by
			i = instrrev(sqlCount, "ORDER BY", -1, vbTextCompare)
			if i > 0 then
				sqlCount = left(sqlCount, i-1)
			end if
'response.write sqlCount & "<br><br>"			
			RecordCount = cInteger(GetValueList(conn, rs, sqlCount))		
			if cString(RecordCount) = "" then
				'calcolo recordcount non andato a buonfine
				RecordCount = null
			else
				'calcola il numero di pagine presenti
				PageSize = IIF(cInteger(PageSize)>0, cInteger(PageSize), 10)
				PageCount = RecordCount \ PageSize
				if Recordcount Mod PageSize > 0 then
					PageCount = PageCount + 1
				end if
				
				'costruisce query con limitazione records
				i = instr(1, sql, "SELECT", vbTextCompare)
				sql = " SELECT TOP " & (PageSize * PageNo) &  " " & _
					  right(sql, len(sql) - (i + 6))
'response.write sql & "<br><br>"
'response.end
				rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
				rs.PageSize = PageSize
			end if
		else
			RecordCount = null
		end if

		if isNull(RecordCount) then
			rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
			rs.PageSize = PageSize
			Recordcount = rs.recordcount
			PageCount = rs.PageCount
		end if
		if PageNo > PageCount then
			PageNo = PageCount
		end if
	end sub
	
	
	
'DEFINIZIONE PRORIETA'*******************
	
	Public Property Let ParentConfiguration(obj)
		set ParentConfigObject = obj
		Lingua = ParentConfigObject.Lingua
		BaseUrl = ParentConfigObject.BaseURL
		PageName = ParentConfigObject.PageName
	end property

'FUNZIONI PRIVATE************************

end class
%>