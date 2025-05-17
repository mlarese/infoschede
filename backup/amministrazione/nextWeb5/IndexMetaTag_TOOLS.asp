<%
'**************************************************************************************************
'	dichiarazione funzioni per gestione meta tag
'**************************************************************************************************

'dichiara tipi
const VISUAL_INDICE = " ok"
const VISUAL_CONTENUTO = " warning"
const VISUAL_CONTENUTOBASE = " alert"
const VISUAL_PAGINA = " visibile"
const VISUAL_SITO = "_disabled"
const VISUAL_VUOTO = " errore"


'dichiara legenda icone e significati
dim VisualTypes
set VisualTypes = Server.CreateObject("Scripting.dictionary")
CALL VisualTypes.Add(VISUAL_INDICE, "Valore derivato direttamente dalla voce dell'indice.")
CALL VisualTypes.Add(VISUAL_CONTENUTO, "Valore derivato dal contenuto associato alla voce dell'indice.")
CALL VisualTypes.Add(VISUAL_CONTENUTOBASE, "Valore derivato dal contenuto stesso e gestito dal modulo corrispondente.")
CALL VisualTypes.Add(VISUAL_PAGINA, "Valore derivato dalle impostazioni di default della pagina.")
CALL VisualTypes.Add(VISUAL_SITO, "Valore derivato dalle impostazioni di default del sito.")
CALL VisualTypes.Add(VISUAL_VUOTO, "Valore non impostato.")


'funzione che genera l'icona.
sub WriteIcon(visualType) 
	if VisualTypes.Exists(visualType) then%>
		<span class="icona_biggest<%= visualType %>" title="<%= VisualTypes(visualType) %>">&nbsp;</span>
	<% end if
end sub


sub WriteLegenda( style, colspan )
	dim visualType
	for each visualType in VisualTypes.keys %>
		<tr>
			<td class="content_center" <%= IIF(style <> "", " style=""" + style + """", "") %>>
				<% WriteIcon(visualType) %>
			</td>
			<td class="content" colspan="<%=colspan%>"><%= visualTypes(visualType) %></td>
		</tr>
	<% next
end sub


sub WriteSyncroLock(field)
	if cString(field)<>"" then %>
		<img src="<%= GetAmministrazionePath() %>grafica/padlock.gif" alt="Valore sincronizzato automaticamente dal contenuto.<%= vbCrLF %>Personalizzabile solo a livello di voce dell'indice.">
		&nbsp;
	<% end if
end sub



function GetTitle(IndexValue, ContentValue, PageValue, BaseValue, Lingua, byref Visual, byref Text, byref Title)
		if IndexValue <> "" then
			GetTitle = IndexValue
			Visual = VISUAL_INDICE
			Text = "Alt dall'indice"
		elseif ContentValue <> "" then
			GetTitle = ContentValue
			Visual = VISUAL_CONTENUTO
			Text = "Alt dal contenuto"
		elseif BaseValue<>"" then
			GetTitle = BaseValue
			Visual = VISUAL_CONTENUTOBASE
			Text = "Titolo del contenuto"
		else
			GetTitle = PageValue
			Visual = VISUAL_PAGINA
			Text = "Nome della pagina"
		end if
		if GetTitle="" then
			Visual = VISUAL_VUOTO
			Text = "non impostato"
		end if
		
		Title = "TITOLO FINALE VISIBILE SUL BROWSER (" & Lingua & "): " + vbCrLf + _
				"&quot;" + Server.HtmlEncode(GetTitle) + "&quot;" + vbCrLf + vbCrLF + _
				"DALL' INDICE: " + vbCrLf + _
				"&quot;" + Server.HtmlEncode(IndexValue) + "&quot;" + vbCrLf + vbCrLF + _
				"DAL CONTENUTO : " + vbCrLf + _
				"&quot;" + Server.HtmlEncode(ContentValue) + "&quot;" + vbCrLf + vbCrLF + _
				"DAL TITOLO DEL CONTENUTO: " + vbCrLf + _
				"&quot;" + Server.HtmlEncode(BaseValue) + "&quot;" + vbCrLf + vbCrLF
end function


function GetKeywords(IndexValue, ContentValue, PageValue, BaseValue, Lingua, byref Visual, byref Text, byref Title)
			if IndexValue <> "" then
				GetKeywords = IndexValue
				Visual = VISUAL_INDICE
				Text = "Dall'indice"
			elseif ContentValue <> "" then
				GetKeywords = ContentValue
				Visual = VISUAL_CONTENUTO
				Text = "Dal contenuto"
			elseif PageValue <> "" then
				GetKeywords = PageValue
				Visual = VISUAL_PAGINA
				Text = "Dalla pagina"
			else
				GetKeywords = BaseValue
				Visual = VISUAL_SITO
				Text = "Dal sito"
			end if
			if GetKeywords="" then
				Visual = VISUAL_VUOTO
				Text = "non impostato"
			end if
			Title = "KEYWORDS FINALI APPLICATE (" & Lingua & "): " + vbCrLf + _
					"&quot;" + Server.HtmlEncode(GetKeywords) + "&quot;" + vbCrLf + vbCrLF + _
					"DALL' INDICE: " + vbCrLf + _
					"&quot;" + Server.HtmlEncode(IndexValue) + "&quot;" + vbCrLf + vbCrLF + _
					"DAL CONTENUTO INDICE: " + vbCrLf + _
					"&quot;" + Server.HtmlEncode(ContentValue) + "&quot;" + vbCrLf + vbCrLF + _
					"DALLA PAGINA: " + vbCrLf + _
					"&quot;" + Server.HtmlEncode(PageValue) + "&quot;" + vbCrLf + vbCrLF + _
					"DAL SITO (keywords generiche): " + vbCrLf + _
					"&quot;" + Server.HtmlEncode(BaseValue) + "&quot;" + vbCrLf + vbCrLF
end function



function GetDescription(IndexValue, ContentValue, PageValue, BaseValue, Lingua, byref Visual, byref Text, byref Title)
			if IndexValue <> "" then
				GetDescription = IndexValue
				Visual = VISUAL_INDICE
				Text = "Dall'indice"
			elseif ContentValue <> "" then
				GetDescription = ContentValue
				Visual = VISUAL_CONTENUTO
				Text = "Dal contenuto"
			elseif PageValue <> "" then
				GetDescription = PageValue
				Visual = VISUAL_PAGINA
				Text = "Dalla pagina"
			else
				GetDescription = BaseValue
				Visual = VISUAL_SITO
				Text = "Dal sito"
			end if
			if GetDescription="" then
				Visual = VISUAL_VUOTO
				Text = "non impostato"
			end if
			Title = "DESCRIPTION FINALE APPLICATA (" & Lingua & "): " + vbCrLf + _
					"&quot;" + Server.HtmlEncode(GetDescription) + "&quot;" + vbCrLf + vbCrLF + _
					"DALL' INDICE: " + vbCrLf + _
					"&quot;" + Server.HtmlEncode(IndexValue) + "&quot;" + vbCrLf + vbCrLF + _
					"DAL CONTENUTO INDICE: " + vbCrLf + _
					"&quot;" + Server.HtmlEncode(ContentValue) + "&quot;" + vbCrLf + vbCrLF + _
					"DALLA PAGINA: " + vbCrLf + _
					"&quot;" + Server.HtmlEncode(PageValue) + "&quot;" + vbCrLf + vbCrLF + _
					"DAL SITO (keywords generiche): " + vbCrLf + _
					"&quot;" + Server.HtmlEncode(BaseValue) + "&quot;" + vbCrLf + vbCrLF
end function




'**************************************************************************************************
'	dichiarazione funzioni per motore di ricerca sull'indice
'**************************************************************************************************

function IndexSearchEngineSetFilter(conn, ElencoByUnion)
	dim sql
	'imposta ricerca
	if Request.ServerVariables("REQUEST_METHOD")="POST" then
		Pager.Reset()
		CALL SearchSession_Reset("idx_")
		if not(request("tutti")<>"") then
			CALL SearchSession_Set("idx_")
		end if
	end if
	
	
	'filtra per titolo
	if Session("idx_titolo")<>"" then
		sql = sql & " AND " & SQL_FullTextSearch(Session("idx_titolo"), FieldLanguageList(IIF(ElencoByUnion, "TIP_C0.", "") + "co_titolo_"))
	end if
	
	'filtra per codice
	if Session("idx_chiave")<>"" then
		sql = sql & " AND " & SQL_FullTextSearch(Session("idx_chiave"), FieldLanguageList(IIF(ElencoByUnion, "TIP_C0.", "") + "co_chiave_"))
	end if
	
	'filtra per livello
	if session("idx_livello")<>"" then
		sql = sql & " AND " + IIF(ElencoByUnion, "TIP_L0.", "") + "idx_livello=" & session("idx_livello")
	end if
	
	'ricerca per stato visibile
	if Session("idx_visibile")<>"" then
		if not (instr(1, Session("idx_visibile"), "1", vbTextCompare)>0 AND _
			    instr(1, Session("idx_visibile"), "0", vbTextCompare)>0 ) then
			if instr(1, Session("idx_visibile"), "1", vbTextCompare)>0 then
				'visibile
				sql = sql & " AND ("& SQL_IsTrue(conn, IIF(ElencoByUnion, "TIP_L0.", "") + "idx_visibile_assoluto") & " AND " + _
							  "(" & SQL_DateDiff(conn, "d", IIF(ElencoByUnion, "TIP_C0.", "") + "co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, IIF(ElencoByUnion, "TIP_C0.", "") + "co_data_pubblicazione") & ") AND " + _
							  "("& SQL_DateDiff(conn, "d", IIF(ElencoByUnion, "TIP_C0.", "") + "co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, IIF(ElencoByUnion, "TIP_C0.", "") + "co_data_scadenza") & ") ) "
			elseif instr(1, Session("idx_visibile"), "0", vbTextCompare)>0 then
				'non visibile
				sql = sql & " AND NOT ("& SQL_IsTrue(conn, IIF(ElencoByUnion, "TIP_L0.", "") + "idx_visibile_assoluto") & " AND " + _
							  "(" & SQL_DateDiff(conn, "d", IIF(ElencoByUnion, "TIP_C0.", "") + "co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, IIF(ElencoByUnion, "TIP_C0.", "") + "co_data_pubblicazione") & ") AND " + _
							  "("& SQL_DateDiff(conn, "d", IIF(ElencoByUnion, "TIP_C0.", "") + "co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, IIF(ElencoByUnion, "TIP_C0.", "") + "co_data_scadenza") & ") ) "
			end if
		end if
	end if
	
	'ricerca per categoria padre
	if Session("idx_categoria")<>"" then	
		sql = sql & " AND " + IIF(ElencoByUnion, "TIP_L0.", "") + "idx_padre_id = " & cInteger(Session("idx_categoria"))
	end if
	
	'ricerca per tabella contenuto
	if CIntero(session("idx_tipoContenuto")) > 0 then
		sql = sql &" AND " + IIF(ElencoByUnion, "TIP_C0.", "") + "co_F_table_id = "& session("idx_tipoContenuto")
	end if
	
	'ricerca per tabella contenuto
	if CIntero(session("idx_id")) > 0 then
		sql = sql &" AND idx_id = "& session("idx_id")
	end if
	
	if sql <> "" then
		IndexSearchEngineSetFilter = right(sql, len(sql) - 4)
	else
		IndexSearchEngineSetFilter = ""
	end if
	
end function
%>