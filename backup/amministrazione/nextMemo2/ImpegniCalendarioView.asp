<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/ClassIndexAlberi.asp" -->
<%
dim dicitura, conn, rs, rsp, rsr, sql_filtri_ricerca
set dicitura = New testata

dicitura.iniz_sottosez(2)
dicitura.sottosezioni(1) = "TIPOLOGIE"
dicitura.links(1) = "ImpegniTipologie.asp"
dicitura.sottosezioni(2) = "CONFIGURAZIONE"
dicitura.links(2) = "AgendaConfigura.asp"


dicitura.sezione = "Gestione impegni/appuntamenti - calendario"
dicitura.puls_new = "NUOVO IMPEGNO"
dicitura.link_new = "ImpegniNew.asp?FROM=calendario"
dicitura.scrivi_con_sottosez() 

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


%>
<!--#INCLUDE FILE="ImpegniFiltriRicerca.asp" -->
<%
sql_filtri_ricerca = Replace(sql,"*","imp_id")
sql_filtri_ricerca = Left(sql_filtri_ricerca,instr(sql_filtri_ricerca,"ORDER BY")-1)

sql = ""

dim profili_attivi

sql = "SELECT pro_id FROM mtb_profili"
if cString(GetValueList(conn, NULL, sql)) <> "" then
	profili_attivi = true
else
	profili_attivi = false
end if

%>

<style type="text/css"> 
	td.giorno {
	  background: #f4f4f2;
	  padding-left:3px;
	  padding-right:3px;
	  padding-top:2px;
	  border-top:1px solid white;
	  border-right:1px solid white;
	  vertical-align:top;
	}
	
	.free {
	  background:#f4f4f2;
	  padding:1px;
	  text-align:center;
	  border-right:1px solid white;
	  border-top:1px solid white;
	}
	
	.selected {
		border-top:1px solid white;
		border-right:1px solid white;
		padding-left:3px;
		padding-right:3px;
	}
	
	td.utente{
		background:#e7e8e2;
		padding-left:3px;
		border-top:1px solid white;
		border-right:1px solid white;
	}
	
	.no_utente{
		background:#e7e8e2;
		border-right:1px solid white;
		border-top:1px solid white;
	}
	
	td.profilo{
		background:#e7e8e2;
		padding-left:3px;
		border-top:1px solid white;
		border-right:1px solid white;
		font-weight:bold;
	}
	
	a.new_impegno{
		width:12px;
		line-height:10px;
		padding:0px !important;
		border:1px solid #999999;
		background-color:white;
		text-align:center;
		font-weight:bold;
		text-decoration:none !important;
	}
		
	a.new_impegno:hover{
		border:1px solid #e38000 !important;
	}
	
	a.cancella{
		width:12px;
		line-height:8px;
		padding-top:0px !important;
		padding-bottom:2px;
		border:1px solid #999999;
		background-color:white;
		text-align:center;
		font-weight:bold;
		text-decoration:none !important;
		font-size:10px;
	}
	
	a.cancella:hover{
		border:1px solid #e38000 !important;
	}
	
	a.modifica{
		padding-left:2px;
		padding-right:2px;
	}
	
	span.operazioni{
		padding-top:1px;
		padding-bottom:1px;
		font-size:1px !important;
		width:62px;
		text-align:right;
	}
	
	span.utente,
	span.profilo{
		width:140px;
		padding-right:3px;
	}
</style>



<div id="content">
    <table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
	    <caption class="border">Lista impegni/appuntamenti - elenco</caption>
        <tr>
            <td class="content">Visualizza gli impegni e gli appuntamenti come elenco </td>
            <td class="content_right">
                <a class="button" href="Impegni.asp" title="Apre la visualizzazione come elenco.">
				    VISUALIZZA COME ELENCO
                </a>
            </td>
        </tr>
    </table>
	
	<% 
	dim max_width, day_number, min_hour, max_hour, column_num, ora_first, ora_last, column_width, first_week_date, last_week_date, week_day, classe
	dim add_style, query, id_utente, chk_impegno_scritto, id_profilo, change_first_week_date, testo, colspan, min_hour_day, max_hour_day, last_id_mese
	dim controllo, i, sql, a, query_cond, ok_write, id_imp_per_mod, force_new_line, conta, ultimo_id_scritto, data_per_confronto, int_inf, int_sup, title
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
	
	max_width = 950

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
	column_width = cIntero(max_width/column_num+2)
	
	
	'riempio tutta la matrice che userò poi per il confronto con gli impegni
	for i=1 to 8
		for a=1 to 720
			calendario(i,a) = DateIta(DateAdd("d", i-1,first_week_date)) & " " & calendario(0,a)
		next
	next

	%>
	<table cellspacing="0" cellpadding="0" class="tabella_madre" style="width:<%=max_width+250%>px; ">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<%
					change_first_week_date = DateIta(DateAdd("d", -7,first_week_date))
					%>
					<td align="left">
						<a class="button" href="?FIRSTDATE=<%=change_first_week_date%>" title="vai alla settimana precedente">
							&lt;&lt; SETTIMANA PRECEDENTE
						</a>
					</td>
					<td class="caption" align="center">da <%=DataEstesa(first_week_date,LINGUA_ITALIANO)%> a <%=DataEstesa(last_week_date,LINGUA_ITALIANO)%></td>
					<%
					change_first_week_date = DateIta(DateAdd("d", 7,first_week_date))
					%>
					<td align="right">
						<a class="button" href="?FIRSTDATE=<%=change_first_week_date%>" title="vai alla settimana successiva">
							SETTIMANA SUCCESSIVA &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<!-- scrivo la riga con l'orario -->
		<tr>
			<th style="width:110px;">&nbsp;</td>
			<th style="width:140px;">&nbsp;</td>
			<% for i=1 to column_num %>
				<% if i MOD 2 then %>
					<th colspan="2" style="width:<%=IIF(i=column_num,column_width,column_width*2)%>px; text-align:left; font-size:9px !important;"><%=TimeIta(calendario(0,i))%></td>
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
					<td style="line-height:1px; border-bottom:1px solid #999999;" colspan="<%=column_num+2%>">
						&nbsp;
					</td>
				</tr>
			<% end if %>
			<% last_id_mese = Month(calendario(day_number,1)) %>
			<tr>
				<td class="giorno">
					<span style="float:left;"><%=Left(NomeGiorno(calendario(day_number,1),LINGUA_ITALIANO), 3)%></span>
					<span style="float:right; width:10px; margin-left:5px;">
						<a href="ImpegniNew.asp?DATA_INIZIO=<%=Server.UrlEncode(DateIta(calendario(day_number,1)))%>&RETURN_DATE=<%=Server.UrlEncode(first_week_date)%>" 
								class="new_impegno" title="aggiungi nuovo impegno il <%=DateIta(calendario(day_number,1))%>">+</a>
					</span>
					<span style="float:right;" class="note"><%=DateIta(calendario(day_number,1))%></span>
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
					'rs.Supports(adIndex)
					%>
						<% if conta=1 then %>
							<td class="utente"><span class="utente"><%=cString(GetNomeUtente(id_utente))%></span></td>
							<% ultimo_id_scritto = id_utente %>
						<% else %>
							<tr>
								<td class="giorno" style="border-top:0px;">&nbsp;</td>
								<% if ultimo_id_scritto = id_utente then %>
									<td class="utente" style="border-top:0px;">&nbsp;</td>
								<% else %>
									<td class="utente"><span class="utente"><%=cString(GetNomeUtente(id_utente))%></span></td>
									<% ultimo_id_scritto = id_utente %>
								<% end if %>
						<% end if %>
						
						<%
						for i=1 to column_num
							classe = "free"
							if TimeIta(calendario(day_number,i)) < TimeIta(min_hour_day) OR _
								TimeIta(calendario(day_number,i)) > TimeIta(max_hour_day) then
								add_style = "background:"&inactive_cell_color&";"
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
										testo = rs("imp_titolo_it")
										title = TimeIta(rs("imp_data_ora_inizio")) & " - " & TimeIta(rs("imp_data_ora_fine"))
										id_imp_per_mod = rs("imp_id")
									end if
									
									'se arrivo alla fine delle colonne e l'impegno non è ancora finito
									if i=column_num then
										testo = rs("imp_titolo_it")
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
								<td colspan="<%=IIF(colspan=0,1,colspan)%>" class="<%=classe%>" style="width:<%=column_width*colspan%>px; <%=add_style%>" title="<%=title%>">
									<span style="float:left;"><%=testo%></span>
									<% if colspan > 0 then %>
										<span style="float:right;" class="operazioni">
											<a class="cancella modifica" href="ImpegniMod.asp?ID=<%=id_imp_per_mod%>&RETURN_DATE=<%=Server.UrlEncode(first_week_date)%>" title="modifica impegno">modifica</a>
											<a class="cancella" href="javascript:void(0);" onclick="OpenDeleteWindow('IMPEGNI','<%=id_imp_per_mod%>');" title="cancella impegno">x</a>
										</span>
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
						add_style = "background:"&inactive_cell_color&";"
					else
						add_style = ""
					end if
					title = TimeIta(calendario(day_number,i))
					%>
					<td class="<%=classe%>" style="width:<%=column_width%>px; <%=add_style%>" title="<%=title%>">&nbsp;</td>
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
						<% if conta = 1 then %>
							<td class="profilo"><span class="profilo"><%=GetNomeProfilo(id_profilo, LINGUA_ITALIANO)%></span></td>
							<% ultimo_id_scritto = id_profilo %>
						<% else %>
							<tr>
								<td class="giorno" style="border-top:0px;">&nbsp;</td>
								<% if  ultimo_id_scritto = id_profilo then %>
									<td class="profilo" style="border-top:0px;">&nbsp;</td>
								<% else %>
									<td class="profilo"><span class="profilo"><%=GetNomeProfilo(id_profilo, LINGUA_ITALIANO)%></span></td>
									 <% ultimo_id_scritto = id_profilo %>
								<% end if %>
						<% end if %>
						
						<%
						for i=1 to column_num
							classe = "free"
							if TimeIta(calendario(day_number,i)) < TimeIta(min_hour_day) OR _
								TimeIta(calendario(day_number,i)) > TimeIta(max_hour_day) then
								add_style = "background:"&inactive_cell_color&";"
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
										testo = rsp("imp_titolo_it")
										title = TimeIta(rsp("imp_data_ora_inizio")) & " - " & TimeIta(rsp("imp_data_ora_fine"))
										id_imp_per_mod = rsp("imp_id")
									end if
									
									'se arrivo alla fine delle colonne e l'impegno non è ancora finito
									if i=column_num then
										testo = rsp("imp_titolo_it")
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
								<td colspan="<%=IIF(colspan=0,1,colspan)%>" class="<%=classe%>" style="width:<%=column_width*colspan%>px; <%=add_style%>" title="<%=title%>">
									<span style="float:left;"><%=testo%></span>
									<% if colspan > 0 then %>
										<span style="float:right;" class="operazioni">
											<a class="cancella modifica" href="ImpegniMod.asp?ID=<%=id_imp_per_mod%>&RETURN_DATE=<%=Server.UrlEncode(first_week_date)%>" title="modifica impegno">modifica</a>
											<a class="cancella" href="javascript:void(0);" onclick="OpenDeleteWindow('IMPEGNI','<%=id_imp_per_mod%>');" title="cancella impegno">x</a>
										</span>
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
	
	<br><br>

	<% CALL WriteBloccoRicerca(conn,"horizontal") %>
	

</div>
</body>
</html>