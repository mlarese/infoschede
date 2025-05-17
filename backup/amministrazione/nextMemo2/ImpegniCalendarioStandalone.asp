<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="Tools_Categorie.asp" -->
<!--#INCLUDE FILE="Tools_Memo2.asp" -->
<!--#INCLUDE FILE="../library/class_testata.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->


<%
'*----------------------------------------------- PARAMETRI PASSATI DA QUERYSTRING (da .NET) -------------------------*
'ID_USER		ID UTENTE che ha effettuato l'accesso all'area riservata
'ID_ALTRO_USER	id UTENTE per vedere tutti gli impegni di un utente + tutti gli impegni pubblici (NON OBBLIGATORIO)
'ID_PROFILO		id del profilo per filtrare gli impegni visibili da quel profilo (NON OBBLIGATORIO)
'ID_PAGE_VIEW	id pagina sito della scheda dell'impegno
'FIRSTDATE		data formato gg/mm/aaaa per settare la visualizzazione del calendario
'LINGUA


dim ID_user, ID_altro_user, ID_profilo_qrst, ID_page_view, lingua
ID_user = cIntero(request("ID_USER"))
ID_altro_user = cIntero(request("ID_ALTRO_USER"))
ID_profilo_qrst = cIntero(request("ID_PROFILO"))
ID_page_view = cIntero(request("ID_PAGE_VIEW"))
lingua = ParseSQL(request("LINGUA"), adChar)

if cString(lingua) = "" then lingua = "en"


'*--------------------------------------------------------------------------------------------------------------------*



dim conn, rs, rsp, rsr, sql_filtri_ricerca, sql

dim intervallo
intervallo = cIntero(Session("AGENDA_INTERVALLO_CALENDARIO"))
if intervallo = 0 then
	intervallo = 30
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsp = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")



sql = ""

'filtro per id utente
if cIntero(ID_user)>0 OR cIntero(ID_altro_user)>0 then
	dim ut_ids
	ut_ids = ID_user
	if ID_altro_user<>"" then
		if cString(ut_ids) <> "" then
			ut_ids = ut_ids & "," & ID_altro_user
		else
			ut_ids = ID_altro_user
		end if
	end if
	sql = sql & " AND (((imp_id IN (SELECT riu_impegno_id FROM mrel_impegni_utenti WHERE riu_utente_id IN (" & ut_ids & "))) OR " & _
				" 		(imp_id IN (SELECT rip_impegno_id FROM mrel_impegni_profili WHERE rip_profilo_id IN " & _
				"											(SELECT rpu_profilo_id FROM mrel_profili_utenti WHERE rpu_utenti_id IN (" & ut_ids & "))))))"
end if

'ricerca per profilo collegato agli impegni
if cIntero(ID_profilo_qrst)>0 then
	sql = sql & " AND imp_id IN (SELECT rip_impegno_id FROM mrel_impegni_profili WHERE rip_profilo_id = " & ID_profilo_qrst & ")"
end if


	
sql_filtri_ricerca = "(SELECT imp_id FROM mtb_impegni "
sql_filtri_ricerca = sql_filtri_ricerca & " WHERE ((1=1) " & sql & ") "
if cIntero(ID_user)=0 then
	sql_filtri_ricerca = sql_filtri_ricerca & " OR (NOT " & SQL_IsTrue(conn, "imp_protetto") & ")"
end if
sql_filtri_ricerca = sql_filtri_ricerca & ")"


dim profili_attivi

sql = "SELECT pro_id FROM mtb_profili"
if cString(GetValueList(conn, NULL, sql)) <> "" then
	profili_attivi = true
else
	profili_attivi = false
end if

%>


<div id="content_calendario">
	
	<% 
	dim max_width, day_number, min_hour, max_hour, column_num, ora_first, ora_last, column_width, first_week_date, last_week_date, week_day, classe
	dim add_style, query, id_utente, chk_impegno_scritto, id_profilo, change_first_week_date, testo, colspan, min_hour_day, max_hour_day, last_id_mese
	dim controllo, i, a, query_cond, ok_write, id_imp_per_mod, force_new_line, conta, ultimo_id_scritto, data_per_confronto, int_inf, int_sup, title
	const max_column_num = 720
	const inactive_cell_color = "#dadbd5"
	dim calendario(8, 720)
	

	if request("FIRSTDATE")<>"" AND IsDate(request("FIRSTDATE")) then
		first_week_date = request("FIRSTDATE")
		week_day = WeekDay(first_week_date)
	else
		first_week_date = Date()
		week_day = WeekDay(Now())
	end if
	
	if week_day > 1 then
		first_week_date = DateAdd("d", -1*(week_day-1),first_week_date)
	end if
	last_week_date = DateAdd("d", 6,first_week_date)
	

	query = " FROM mtb_configurazione_impegni " & _
			" WHERE " & SQL_CompareDateTime(conn, "coi_dal", adCompareLessThan, last_week_date) & _
			" AND "& SQL_CompareDateTime(conn, "coi_al", adCompareGreaterThan, first_week_date)
	'imposto l'ora minore nella settimana selezionata				 
	sql = " SELECT MIN(CONVERT(datetime,CONVERT(nvarchar, DATEPART(hh, coi_dal)) + ':' + CONVERT(nvarchar, DATEPART(n, coi_dal)) + ':00')) AS Expr1 " & _
		  query & _
		  " AND coi_id IN (SELECT MAX(coi_id) "&query&" GROUP BY coi_giorno)"
	min_hour = "06:00"
	if cString(GetValueList(conn,NULL,sql))<>"" then
		min_hour = TimeIta(GetValueList(conn,NULL,sql))
	end if
	ora_first = DATE() & " "&min_hour&":00"
	'controllo se nella settimana selezionata c'è un impegno con orario minore a quello calcolato dalla configurazione
	sql = " SELECT MIN(CONVERT(datetime,CONVERT(nvarchar, DATEPART(hh, coi_dal)) + ':' + CONVERT(nvarchar, DATEPART(n, coi_dal)) + ':00')) AS Expr1 " & _
		  query
	sql = Replace(sql,"mtb_configurazione_impegni","mtb_impegni")
	sql = Replace(sql,"coi_dal","imp_data_ora_inizio")
	sql = Replace(sql,"coi_al","imp_data_ora_fine")
	if cString(GetValueList(conn,NULL,sql))<>"" then
		data_per_confronto = TimeIta(GetValueList(conn,NULL,sql))
		data_per_confronto = DATE() & " "&data_per_confronto&":00"
		if DateDiff("n",ora_first,data_per_confronto)<0 then
			min_hour = TimeIta(data_per_confronto)
		end if	
	end if
	
	
	'imposto l'ora maggiore nella settimana selezionata
	sql = " SELECT MAX(CONVERT(datetime,CONVERT(nvarchar, DATEPART(hh, coi_al)) + ':' + CONVERT(nvarchar, DATEPART(n, coi_al)) + ':00')) AS Expr1 " & _
		  query & _
		  " AND coi_id IN (SELECT MAX(coi_id) "&query&" GROUP BY coi_giorno)"
	max_hour = "21:00"
	if cString(GetValueList(conn,NULL,sql))<>"" then
		max_hour = TimeIta(GetValueList(conn,NULL,sql))
	end if
	ora_last = DATE() & " "&max_hour&":00"
	'controllo se nella settimana selezionata c'è un impegno con orario maggiore a quello calcolato dalla configurazione
	sql = " SELECT MAX(CONVERT(datetime,CONVERT(nvarchar, DATEPART(hh, coi_al)) + ':' + CONVERT(nvarchar, DATEPART(n, coi_al)) + ':00')) AS Expr1 " & _
		  query
	sql = Replace(sql,"mtb_configurazione_impegni","mtb_impegni")
	sql = Replace(sql,"coi_dal","imp_data_ora_inizio")
	sql = Replace(sql,"coi_al","imp_data_ora_fine")
	if cString(GetValueList(conn,NULL,sql))<>"" then
		data_per_confronto = TimeIta(GetValueList(conn,NULL,sql))
		data_per_confronto = DATE() & " "&data_per_confronto&":00"
		if DateDiff("n",ora_last,data_per_confronto)>0 then
			max_hour = TimeIta(data_per_confronto)
		end if	
	end if
	
	ora_first = DATE() & " "&min_hour&":00"
	ora_last = DATE() & " "&max_hour&":00"

	
	'calcolo il numero di colonne
	column_num = 1
	do while true
		calendario(0,column_num) = TimeITA(ora_first)
		ora_first = DateAdd("n", intervallo, ora_first)
		column_num = column_num + 1
		if Hour(ora_first) = Hour(ora_last) then
			controllo = Minute(ora_first)
			while controllo <= Minute(ora_last)
				calendario(0,column_num) = TimeITA(ora_first)
				ora_first = DateAdd("n", intervallo, ora_first)
				column_num = column_num + 1
				controllo = controllo + intervallo
				'calendario(0,column_num) = TimeITA(ora_first)
			wend
			exit do
		end if
	loop
	'--------------------
	column_num = column_num - 1
	'--------------------
	
	'max_width in percentuale
	max_width = 70
	column_width = Round(cReal(max_width/column_num+2), 2)
	'response.write column_width
	'response.end
	
	
	'riempio tutta la matrice che userò poi per il confronto con gli impegni
	for i=1 to 8
		for a=1 to 720
			calendario(i,a) = DateIta(DateAdd("d", i-1,first_week_date)) & " " & calendario(0,a)
		next
	next

	%>
	<table cellspacing="0" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption" align="center">
						<% if lingua = LINGUA_ITALIANO then %>
							da <%=DataEstesa(first_week_date,lingua)%> a <%=DataEstesa(last_week_date,lingua)%>
						<% else %>
							from <%=DataEstesa(first_week_date,lingua)%> to <%=DataEstesa(last_week_date,lingua)%>
						<% end if %>
					</td>
				</tr>
			</table>
		</caption>
		<!-- scrivo la riga con l'orario -->
		<tr>
			<!--
			<th style="width:90px;">&nbsp;</td>
			<th style="width:110px;">&nbsp;</td>
			-->
			<th class="colonna1">&nbsp;</td>
			<th class="colonna2">&nbsp;</td>
			<% for i=1 to column_num %>
				<% if i MOD 2 then %>
					<th colspan="2" class="<%=IIF(i=column_num,"last_column","")%>"><%=FixLenght(Hour(calendario(0,i)), "0", 2)%></td>
				<% end if %>
			<% next %>
		</tr>
		<% last_id_mese = Month(first_week_date) %>
		
		<% for day_number=1 to 7 %>
			<% 
			query = " FROM mtb_configurazione_impegni " & _
				    " WHERE " & SQL_CompareDateTime(conn, "coi_dal", adCompareLessThan, last_week_date) & _
				    " AND "& SQL_CompareDateTime(conn, "coi_al", adCompareGreaterThan, first_week_date) & _
				    " AND coi_giorno = " & day_number
			
			'cerco l'ora minore per ogni giorno in maniera da riuscire a marcare gli orari "non attivi"			
			sql = " SELECT MIN(CONVERT(datetime,CONVERT(nvarchar, DATEPART(hh, coi_dal)) + ':' + CONVERT(nvarchar, DATEPART(n, coi_dal)) + ':00')) AS Expr1 " & _
				  query & _
				  " AND coi_id IN (SELECT TOP 1 coi_id "&query&" ORDER BY coi_id DESC)"
			min_hour_day = min_hour
			if cString(GetValueList(conn,NULL,sql))<>"" then
				min_hour_day = TimeIta(GetValueList(conn,NULL,sql))
			else
				min_hour_day = max_hour
			end if
			'min_hour_day = DATE() & " "&min_hour_day&":00"
			
			
			'cerco l'ora maggiore per ogni giorno in maniera da riuscire a marcare gli orari "non attivi"
			sql = " SELECT MAX(CONVERT(datetime,CONVERT(nvarchar, DATEPART(hh, coi_al)) + ':' + CONVERT(nvarchar, DATEPART(n, coi_al)) + ':00')) AS Expr1 " & _
				  query & _
				  " AND coi_id IN (SELECT TOP 1 coi_id "&query&" ORDER BY coi_id DESC)"
			max_hour_day = max_hour
			if cString(GetValueList(conn,NULL,sql))<>"" then
				max_hour_day = TimeIta(GetValueList(conn,NULL,sql))
			else
				max_hour_day = min_hour
			end if
			'max_hour_day = DATE() & " "&max_hour_day&":00"

			
		
			'---- impegni degli UTENTI per giorno -----------------

			sql = " SELECT riu_utente_id, imp_tipo_id, imp_data_ora_inizio, imp_data_ora_fine, imp_titolo_it, imp_titolo_en, imp_id " & _
				  " FROM mtb_impegni INNER JOIN mrel_impegni_utenti ON mtb_impegni.imp_id = mrel_impegni_utenti.riu_impegno_id " & _
				  " WHERE " & SQL_CompareDateTime(conn, "imp_data_ora_inizio", adCompareLessThan, calendario(day_number,column_num)) & _
				  " AND "& SQL_CompareDateTime(conn, "imp_data_ora_fine", adCompareGreaterThan, calendario(day_number,1)) 
			sql = sql & " AND imp_id IN (" & sql_filtri_ricerca & ")" 
			sql = sql & " ORDER BY riu_utente_id, imp_data_ora_inizio, imp_data_ora_fine "
			rs.open sql, conn, adOpenStatic, adLockOptimistic 

			'--------------------------------------------------------------
			

			'---- Elenco impegni dei PROFILI per giorno -------------------
			
			sql = " SELECT rip_profilo_id, imp_tipo_id, imp_data_ora_inizio, imp_data_ora_fine, imp_titolo_it, imp_titolo_en, imp_id " & _
				  " FROM mtb_impegni INNER JOIN mrel_impegni_profili ON mtb_impegni.imp_id = mrel_impegni_profili.rip_impegno_id " & _
				  " WHERE " & SQL_CompareDateTime(conn, "imp_data_ora_inizio", adCompareLessThan, calendario(day_number,column_num)) & _
				  " AND " & SQL_CompareDateTime(conn, "imp_data_ora_fine", adCompareGreaterThan, calendario(day_number,1))
			sql = sql & " AND imp_id IN (" & sql_filtri_ricerca & ")" 
			sql = sql & " ORDER BY rip_profilo_id, imp_data_ora_inizio, imp_data_ora_fine "
			rsp.open sql, conn, adOpenStatic, adLockOptimistic 
			
			'--------------------------------------------------------------
			
			
			id_utente = ""
			%>
			
			<% if last_id_mese <> Month(calendario(day_number,1)) then %>
				<tr>
					<td style="line-height:1px;border-right:none;height:1px !important;background-color:#D1D1D1;padding:0px;" colspan="<%=column_num+2%>">&nbsp;</td>
				</tr>
			<% end if %>
			<% last_id_mese = Month(calendario(day_number,1)) %>
			<tr>
				<td class="giorno">
					<span class="nome_giorno"><%=Left(NomeGiorno(calendario(day_number,1),lingua), 3)%></span>
					<span class="note num_giorno">
						<% if lingua = LINGUA_ITALIANO then %>
							<%=DateIta(calendario(day_number,1))%>
						<% else %>
							<%=DateEN(calendario(day_number,1))%>
						<% end if %>
					</span>
				</td>
			
			<% 
			
			'impegni degli UTENTI
			chk_impegno_scritto = false
			force_new_line = false
			conta = 1
			ultimo_id_scritto = 0
			if not rs.eof then
				while not rs.eof
					
					id_utente = rs("riu_utente_id")
					chk_impegno_scritto = false
					ok_write = true
					colspan = 0
					
					if force_new_line then
						force_new_line = false
					end if 

					%>
					<%  if lingua = LINGUA_ITALIANO then
							title = "utente" 
						else
							title = "user"
						end if
					%>
						<% if conta=1 then %>
							<td class="utente" title="<%=title%>"><%=cString(GetNomeUtente(id_utente))%></td>
							<% ultimo_id_scritto = id_utente %>
						<% else %>
							<tr>
								<td class="giorno" style="border-top:0px;">&nbsp;</td>
								<% if ultimo_id_scritto = id_utente then %>
									<td class="utente" style="border-top:0px;">&nbsp;</td>
								<% else %>
									<td class="utente" title="<%=title%>"><%=cString(GetNomeUtente(id_utente))%></td>
									<% ultimo_id_scritto = id_utente %>
								<% end if %>
						<% end if %>
						
						<%
						for i=1 to column_num
							classe = "free"
							if TimeIta(calendario(day_number,i)) < TimeIta(min_hour_day) OR _
								TimeIta(calendario(day_number,i)) > TimeIta(max_hour_day) then
								add_style = ""
							else
								add_style = ""
							end if
							testo = "&nbsp;"
							title = ""
							if not rs.eof then
								if TimeIta(rs("imp_data_ora_inizio")) <= TimeIta(calendario(day_number,i)) AND _
										TimeIta(rs("imp_data_ora_fine")) => TimeIta(calendario(day_number,i)) AND _
											id_utente = rs("riu_utente_id") AND NOT force_new_line then
									classe = "selected"
									add_style = "background:"&GetColorTipologia(rs("imp_tipo_id"))
									chk_impegno_scritto = true
									
									'inizio "linea dell'impegno"
									int_inf = DateAdd("n", -(1)*intervallo, DateTimeIta(calendario(day_number,i)))
									int_sup = DateAdd("n", intervallo, DateTimeIta(calendario(day_number,i)))
									'if TimeIta(rs("imp_data_ora_inizio")) = TimeIta(calendario(day_number,i)) OR i=1 then
									if (TimeIta(rs("imp_data_ora_inizio")) > TimeIta(int_inf) AND TimeIta(rs("imp_data_ora_inizio")) < TimeIta(int_sup)) OR i=1 then
										ok_write = false
									end if
									
									colspan = colspan + 1
									
									'fine "linea dell'impegno"
									int_inf = DateAdd("n", -(1)*intervallo, DateTimeIta(calendario(day_number,i)))
									int_sup = DateAdd("n", intervallo, DateTimeIta(calendario(day_number,i)))
									'if TimeIta(rs("imp_data_ora_fine")) = TimeIta(calendario(day_number,i)) then
									if TimeIta(rs("imp_data_ora_fine")) > TimeIta(int_inf) AND TimeIta(rs("imp_data_ora_fine")) < TimeIta(int_sup) then
										ok_write = true
										testo = CBLL(rs,"imp_titolo",lingua)
										title = TimeIta(rs("imp_data_ora_inizio")) & " - " & TimeIta(rs("imp_data_ora_fine"))
										id_imp_per_mod = rs("imp_id")
									end if
									
									'se arrivo alla fine delle colonne e l'impegno non è ancora finito
									if i=column_num then
										testo = CBLL(rs,"imp_titolo",lingua)
										title = TimeIta(rs("imp_data_ora_inizio")) & " - " & TimeIta(rs("imp_data_ora_fine"))
										id_imp_per_mod = rs("imp_id")
										ok_write = true
										rs.moveNext
										if not rs.eof then
											if id_utente = rs("riu_utente_id") then
												force_new_line = true
											end if
										end if
									end if
								else
									if chk_impegno_scritto then
										'ho finito un impegno, passo al successivo
										rs.moveNext
										ok_write = true
										if not rs.eof then
											'se l'impegno successivo è consecutivo ("adiacente")
											int_inf = DateAdd("n", -(1)*intervallo, DateTimeIta(calendario(day_number,i)))
											int_sup = DateAdd("n", intervallo, DateTimeIta(calendario(day_number,i)))
											'if TimeIta(rs("imp_data_ora_inizio")) = TimeIta(calendario(day_number,i)) AND id_utente = rs("riu_utente_id") then
											if TimeIta(rs("imp_data_ora_inizio")) > TimeIta(int_inf) AND TimeIta(rs("imp_data_ora_inizio")) < TimeIta(int_sup) AND id_utente = rs("riu_utente_id") then
												ok_write = false
												i = i-1
											end if
											
											'se l'impegno successivo è già iniziato, lo scrivo su una nuova riga
											if TimeIta(rs("imp_data_ora_inizio")) < TimeIta(calendario(day_number,i)) AND id_utente = rs("riu_utente_id") then
												force_new_line = true
											end if
											
										end if
										chk_impegno_scritto = false
										colspan = 0
									end if
								end if
							end if
							
							if ok_write then
								'imposto il tag title prima di scrivere la cella
								if title = "" then
									title = TimeIta(calendario(day_number,i))
								end if
								%>
								<td colspan="<%=IIF(colspan=0,1,colspan)%>" class="<%=classe%>  <%=IIF(i=column_num," last_column","")%>" title="<%=title%>" style="<%=add_style%>">
									<% if testo="&nbsp;" then %>
										<span><%=testo%></span>
									<% else %>
										<!--<a href="<%=GetPageSiteUrl(conn, ID_page_view, lingua) & "&ID=" & id_imp_per_mod %>" title="<%=title%>" 
											style="display:block; color:black !important;"><%=testo%></a>-->
										<a href="<%=GetPageSiteUrl(conn, ID_page_view, lingua) & "&ID=" & id_imp_per_mod %>" title="<%=title%>"><%=testo%></a>
									<% end if %>
								</td>
								<%
							end if
						next
						%>
					</tr>
					<%
					conta = conta + 1
				wend
			elseif rsp.eof then
				classe = "free"
				%>
				<td class="no_utente">&nbsp;</td>
				<%
				for i=1 to column_num 
					if TimeIta(calendario(day_number,i)) < TimeIta(min_hour_day) OR _
						TimeIta(calendario(day_number,i)) > TimeIta(max_hour_day) then
						add_style = ""
					else
						add_style = ""
					end if
					title = TimeIta(calendario(day_number,i))
					%>
					<td class="<%=classe%> <%=IIF(i=column_num," last_column","")%>" title="<%=title%>">&nbsp;</td>
					<%
				next
				%>
				</tr>
				<%						
			end if
			rs.close

			
			
			'impegni dei PROFILI
			chk_impegno_scritto = false
			force_new_line = false
			if not rsp.eof then	
				while not rsp.eof
					id_profilo = rsp("rip_profilo_id")
					chk_impegno_scritto = false
					ok_write = true
					colspan = 0
					
					if force_new_line then
						force_new_line = false
					end if 
					
					%>
					<%  if lingua = LINGUA_ITALIANO then
							title = "gruppo di utenti" 
						else
							title = "group" 
						end if
					%>
						<% if conta = 1 then %>
							<td class="profilo" title="<%=title%>"><%=GetNomeProfilo(id_profilo, lingua)%></td>
							<% ultimo_id_scritto = id_profilo %>
						<% else %>
							<tr>
								<td class="giorno" style="border-top:0px;">&nbsp;</td>
								<% if  ultimo_id_scritto = id_profilo then %>
									<td class="profilo" style="border-top:0px;">&nbsp;</td>
								<% else %>
									<td class="profilo" title="<%=title%>"><%=GetNomeProfilo(id_profilo, lingua)%></td>
									 <% ultimo_id_scritto = id_profilo %>
								<% end if %>
						<% end if %>
						
						<%
						for i=1 to column_num
							classe = "free"
							if TimeIta(calendario(day_number,i)) < TimeIta(min_hour_day) OR _
								TimeIta(calendario(day_number,i)) > TimeIta(max_hour_day) then
								add_style = ""
							else
								add_style = ""
							end if
							testo = "&nbsp;"
							title = ""
							if not rsp.eof then
								if TimeIta(rsp("imp_data_ora_inizio")) <= TimeIta(calendario(day_number,i)) AND _
										TimeIta(rsp("imp_data_ora_fine")) => TimeIta(calendario(day_number,i)) AND _
											id_profilo = rsp("rip_profilo_id") AND NOT force_new_line then
									classe = "selected"
									add_style = "background:"&GetColorTipologia(rsp("imp_tipo_id"))
									chk_impegno_scritto = true
									
									'inizio "linea dell'impegno"
									int_inf = DateAdd("n", -(1)*intervallo, DateTimeIta(calendario(day_number,i)))
									int_sup = DateAdd("n", intervallo, DateTimeIta(calendario(day_number,i)))
									'if TimeIta(rsp("imp_data_ora_inizio")) = TimeIta(calendario(day_number,i)) OR i=1 then
									if (TimeIta(rsp("imp_data_ora_inizio")) > TimeIta(int_inf) AND TimeIta(rsp("imp_data_ora_inizio")) < TimeIta(int_sup)) OR i=1 then
										ok_write = false
									end if
									
									colspan = colspan + 1
									
									'fine "linea dell'impegno"
									int_inf = DateAdd("n", -(1)*intervallo, DateTimeIta(calendario(day_number,i)))
									int_sup = DateAdd("n", intervallo, DateTimeIta(calendario(day_number,i)))
									'if TimeIta(rsp("imp_data_ora_fine")) = TimeIta(calendario(day_number,i)) then
									if TimeIta(rsp("imp_data_ora_fine")) > TimeIta(int_inf) AND TimeIta(rsp("imp_data_ora_fine")) < TimeIta(int_sup) then
										ok_write = true
										testo = CBLL(rsp,"imp_titolo",lingua)
										title = TimeIta(rsp("imp_data_ora_inizio")) & " - " & TimeIta(rsp("imp_data_ora_fine"))
										id_imp_per_mod = rsp("imp_id")
									end if
									
									'se arrivo alla fine delle colonne e l'impegno non è ancora finito
									if i=column_num then
										testo = CBLL(rsp,"imp_titolo",lingua)
										title = TimeIta(rsp("imp_data_ora_inizio")) & " - " & TimeIta(rsp("imp_data_ora_fine"))
										id_imp_per_mod = rsp("imp_id")
										ok_write = true
										rsp.moveNext
										if not rsp.eof then
											if id_profilo = rsp("rip_profilo_id") then
												force_new_line = true
											end if
										end if
									end if
								else
									if chk_impegno_scritto then
										'ho finito un impegno, passo al successivo
										rsp.moveNext
										ok_write = true
										if not rsp.eof then
											'se l'impegno successivo è consecutivo ("adiacente")
											int_inf = DateAdd("n", -(1)*intervallo, DateTimeIta(calendario(day_number,i)))
											int_sup = DateAdd("n", intervallo, DateTimeIta(calendario(day_number,i)))
											'if TimeIta(rsp("imp_data_ora_inizio")) = TimeIta(calendario(day_number,i)) and id_profilo = rsp("rip_profilo_id") then
											if TimeIta(rsp("imp_data_ora_inizio")) > TimeIta(int_inf) AND TimeIta(rsp("imp_data_ora_inizio")) < TimeIta(int_sup) AND id_profilo = rsp("rip_profilo_id") then
												ok_write = false
												i = i-1
											end if
											
											'se l'impegno successivo è già iniziato, lo scrivo su una nuova riga
											if TimeIta(rsp("imp_data_ora_inizio")) < TimeIta(calendario(day_number,i)) AND id_profilo = rsp("rip_profilo_id") then
												force_new_line = true
											end if
										end if
										chk_impegno_scritto = false
										colspan = 0
									end if
								end if
							end if
							
							if ok_write then
								'imposto il tag title prima di scrivere la cella
								if title = "" then
									title = TimeIta(calendario(day_number,i))
								end if
								%>
								<td colspan="<%=IIF(colspan=0,1,colspan)%>" class="<%=classe%>  <%=IIF(i=column_num," last_column","")%>" title="<%=title%>" style="<%=add_style%>">
									<% if testo="&nbsp;" then %>
										<span><%=testo%></span>
									<% else %>
										<!--<a href="<%=GetPageSiteUrl(conn, ID_page_view, lingua) & "&ID=" & id_imp_per_mod %>" title="<%=title%>" 
											style="display:block; color:black !important;"><%=testo%></a>-->
										<a href="<%=GetPageSiteUrl(conn, ID_page_view, lingua) & "&ID=" & id_imp_per_mod %>" title="<%=title%>"><%=testo%></a>
									<% end if %>
								</td>
								<%
							end if
						next
						%>
					</tr>
					<%
					conta = conta + 1
				wend
			end if
			rsp.close
			%>

		<% next %>
	</table>

</div>
</body>
</html>