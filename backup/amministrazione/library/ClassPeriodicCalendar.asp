<%
'**************************************************************************************************************************************************************
'classe per la generazione di calendari per la selezione di periodi.
'**************************************************************************************************************************************************************



class PeriodicCalendar


'DEFINIZIONE VARIABILI PRIVATE E PUBBLICHE: *******************************************************************************************************************
    Public NumeroCalendari
    Public DataInizio
    Public SelectedDays
    Public MarkedDays
    Public MarkedDays_Message
    
    Private JsWrited
    
    Public IsPost
    Public IsPostBack
    
    Private SelectedDays_PrintedCount
    Private MarkedDays_PrintedCount
    
'DEFINIZIONE COSTRUTTORI: *************************************************************************************************************************************

    Private Sub Class_Initialize()
		
        'crea oggetto per collezione date selezionate
		set SelectedDays = Server.CreateObject("Scripting.Dictionary")
		SelectedDays.CompareMode = vbTextCompare
        
        'crea oggetto per collezione date evidenziate
		set MarkedDays = Server.CreateObject("Scripting.Dictionary")
		MarkedDays.CompareMode = vbTextCompare
		
        JsWrited = false
        
        'imposta dati di default per apertura calendario
        NumeroCalendari = 6
        DataInizio = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
        
        'indica se il form deriva da un post del form stesso (spostamento di mesi o submit)
        IsPost = (request.ServerVariables("REQUEST_METHOD") = "POST")
        IsPostBack = (request.form("prec")<>"" OR request.form("succ")<>"")
        
        'carica l'eventuale form
        CALL SelectedDays_LoadFromForm()
        
        'impostazione dei contatori
        SelectedDays_PrintedCount = 0
        MarkedDays_PrintedCount = 0
        
	End Sub
	
	Private Sub Class_Terminate()
		set SelectedDays = nothing
	End Sub


'...............................................................................
'funzioni per gestione date selezionate

    'procedura per aggiungere un giorno alla lista dei "giorni selezionati"
    public sub SelectedDays_ADD(DayDate)
        dim IsoDate
        IsoDate = EncodeDate(DayDate)
        if not SelectedDays.Exists(IsoDate) then
           CALL SelectedDays.Add(IsoDate, cString(DateSerial(Year(cDate(DayDate)), Month(cDate(DayDate)), Day(cDate(DayDate)))) )
        end if
    end sub
    
    
    'procedura per rimuovere un giorno dalla lista dei "giorni selezionati"
    public sub SelectedDays_REMOVE(DayDate)
        dim IsoDate
        IsoDate = EncodeDate(DayDate)
        if SelectedDays.Exists(IsoDate) then
            CALL SelectedDays.Remove(IsoDate)
        end if
    end sub
    
    
    'funzione che restituisce se un giorno e' selezionato o meno
    public function SelectedDays_IsSelected(DayDate)
        dim IsoDate
        IsoDate = EncodeDate(DayDate)
        SelectedDays_IsSelected = SelectedDays.Exists(IsoDate)
    end function
    
    
    'procedura che marchia i giorni selezionati gia' pubblicati
    private sub SelectedDays_MarkAsPrinted(DayDate)
        dim IsoDate
        IsoDate = EncodeDate(DayDate)
        if SelectedDays.Exists(IsoDate) then
            SelectedDays(IsoDate) = NULL
        end if
        
        MarkedDays_PrintedCount = MarkedDays_PrintedCount + 1
    end sub
    
    
    'funzione che verifica se il giorno selezionato e' gia' stato visualizzato
    private function SelectedDays_MarkedAsPrinted(FormattedDayDate)
        if SelectedDays.Exists(FormattedDayDate) then
            SelectedDays_MarkedAsPrinted = IsNull(SelectedDays(FormattedDayDate))
        else
            SelectedDays_MarkedAsPrinted = true
        end if
    end function
    
    
    'procedura che scorre il form recuperando i "giorni selezionati"
    public sub SelectedDays_LoadFromForm()
        dim var
        if IsPost then
            for each var in request.form
                if cIntero(request.form(var))=1 AND _
                   (left(var, 2) = "g_" OR left(var, 2) = "f_" ) then
                    SelectedDays_ADD(DecodeDate(right(var, len(var)-2)))
                end if
            next
        end if
    end sub


'...............................................................................
'funzioni per gestione date evidenziate

    'procedura per aggiungere un giorno alla lista dei "giorni evidenziati"
    public sub MarkedDays_ADD(DayDate, MarkStyle)
        dim IsoDate
        IsoDate = EncodeDate(DayDate)
        if not MarkedDays.Exists(IsoDate) then
           CALL MarkedDays.Add(IsoDate, MarkStyle)
        end if
    end sub
    
    
    'procedura per rimuovere un giorno dalla lista dei "giorni evidenziati"
    public sub MarkedDays_REMOVE(DayDate)
        dim IsoDate
        IsoDate = EncodeDate(DayDate)
        if MarkedDays.Exists(IsoDate) then
            CALL MarkedDays.Remove(IsoDate)
        end if
    end sub
    
    
    'funzione che restituisce se un giorno e' evidenziato o meno
    public function MarkedDays_IsMarked(DayDate)
        dim IsoDate
        IsoDate = EncodeDate(DayDate)
        MarkedDays_IsMarked = MarkedDays.Exists(IsoDate)
    end function
    
    
    'funzione che restituisce lo stile di evidenziazione
    public function MarkedDays_MarkedStyle(DayDate)
        dim IsoDate
        IsoDate = EncodeDate(DayDate)
        if MarkedDays.Exists(IsoDate) then
            MarkedDays_MarkedStyle = MarkedDays(IsoDate)
        end if
    end function
    
    
    'procedura che marchia i giorni evidenziati gia' pubblicati
    private sub MarkedDays_MarkAsPrinted(DayDate)
        dim IsoDate
        IsoDate = EncodeDate(DayDate)
        if MarkedDays.Exists(IsoDate) then
            MarkedDays(IsoDate) = NULL
        end if
        
        MarkedDays_PrintedCount = MarkedDays_PrintedCount + 1
    end sub
    
    
    'funzione che verifica se il giorno evidenziato e' gia' stato visualizzato
    private function MarkedDays_MarkedAsPrinted(FormattedDayDate)
        if MarkedDays.Exists(FormattedDayDate) then
            MarkedDays_MarkedAsPrinted = IsNull(MarkedDays(FormattedDayDate))
        else
            MarkedDays_MarkedAsPrinted = true
        end if
    end function
    
    
'...............................................................................
'funzioni per gestione periodi dalle liste di giorni
	
	'procedura che aggiunge i giorni compresi nel periodo come giorni selezionati
	public sub SelectedDays_AddPeriod(dateFrom, dateTo)
		if IsDate(dateFrom) AND IsDate(dateTo) then
			dim data
			for data = dateFrom to dateTo
    	        CALL SelectedDays_ADD(data)
        	next
		end if
	end sub
	
	
    'funzione che restituisce un array di periodi (data inizio - data fine) estratti dai giorni selezionati
    public function SelectedDays_ExtractPeriods()
        SelectedDays_ExtractPeriods = ExtractPeriods(SelectedDays)
    end function
	
	
	'procedura che aggiunge i giorni compresi nel periodo come giorni marchiati
    public sub MarkedDays_AddPeriod(dateFrom, dateTo)
		dim data
		for data = dateFrom to dateTo
            CALL MarkedDays_ADD(data)
        next
	end sub
	
    
    'funzione che restituisce un array di periodi (data inizio - data fine) estratti dai giorni marchiati
    public function MarkedDays_ExtractPeriods()
        MarkedDays_ExtractPeriods = ExtractPeriods(MarkedDays)
    end function
    
    
'DEFINIZIONE METODI PER GESTIONE INTERNA DATI: ****************************************************************************************************************
    
  

    
    
    'funzione che restituisce un array contenente i periodi (data inizio - data fine) 
    'ricavati dal dictionary di giorni (SelectedDays o MarkedDays)
    private function ExtractPeriods(Days)
        
		if Days.Count > 0 then
			dim data, cData, dateFrom, dateTo, count
	        dateFrom = null
	        dateTo = null
	        count = 0
			
	        'ordina il dizionario in base al formato iso delle date per garantire l'ordinamento cronologico delle date selezionate
	        CALL SortDictionary(Days, 1)
	
	        'scorre le date per individuare quelle contigue e raggrupparle in periodi
	        for each data in Days.Keys
	            cData = DecodeDate(data)
	            if IsNull(dateFrom) then
	                'primo ciclo: inizializza date
	                dateFrom = cData
	                dateTo = cData
	            else
	                if cData = dateAdd("d", 1, dateTo) then
	                    'data contigua alla precedente di chiusura del periodo: lo "allunga"
	                    dateTo = cData
	                else
	                    'data non contigua al periodo: termina periodo
	                    ReDim Preserve periods(1, count)
	                    periods(0, uBound(periods,2)) = dateFrom
	                    periods(1, uBound(periods,2)) = dateTo
	                    count = count + 1
	                    
	                    'e reninizializza periodo per ciclo successivo
	                    dateFrom = cData
	                    dateTo = cData
	                end if
	            end if
	        next
			
			'aggiunge l'ultimo periodo rimasto
	        if not IsNull(dateFrom) AND not IsNull(dateTo) then
	            ReDim Preserve periods(1, count)
	            periods(0, uBound(periods,2)) = dateFrom
	            periods(1, uBound(periods,2)) = dateTo
	        end if
	        
	        ExtractPeriods = periods 
        else
			ExtractPeriods =  NULL
		end if
    end function
	
	
    'funzione che codifica la data nel formato utilizzato per la gestione interna
    private function EncodeDate(Data)
        EncodeDate = replace(DateIso(Data), "-", "_")
    end function
    
    
    'funzione che decodifica la data dal formato utilizzato per la gestione interna
    private function DecodeDate(Data)
        dim DateParts
        DateParts = split(data, "_")
        DecodeDate = DateSerial(DateParts(0), DateParts(1), DateParts(2))
    end function
    
'DEFINIZIONE METODI PER GENERAZIONE HTML CALENDARI: ***********************************************************************************************************
    
    public sub WriteCalendars()
        CALL WriteJsTools()
        
        dim meseInizio, annoInizio, i
        
        'gestione spostamento mese
    	if IsPostBack AND request.form("curDate") <> "" then
    		DataInizio = request.form("curDate")
    		
    		if request.form("prec") <> "" then
    			DataInizio = DateAdd("m", - ((NumeroCalendari - 1)\2 + 1), DataInizio)
    		elseif request.form("succ") <> "" then
    			DataInizio = DateAdd("m", (NumeroCalendari - 1)\2 + 1, DataInizio)
    		end if
    	end if
        
        meseInizio = Month(DataInizio)
   		annoInizio = Year(DataInizio)
        %>
        
        <table cellpadding="0" cellspacing="0" width="100%">
		    <tr>
			    <td rowspan="2" class="calendario_opzioni">
				    <table cellpadding="0" cellspacing="0" width="100%" class="tabella_madre">
					    <caption>Filtri di selezione</caption>
					    <tr>
    						<td>
    							<table cellpadding="0" cellspacing="0" width="100%">
    								<tr><th class="l2" colspan="7">GIORNI</th></tr>
    								<tr>
    									<td class="content" colspan="7">
    										Seleziona tutti
    										<input type="checkbox" class="checkbox" id="d_tutti" name="d_tutti" value="1" <%= chk((IsPost AND request.form("d_tutti")<>"") OR not IsPost) %>
    											   onclick="this.checked = true; form1.d_1.checked=false;form1.d_2.checked=false;form1.d_3.checked=false;form1.d_4.checked=false;form1.d_5.checked=false;form1.d_6.checked=false;form1.d_0.checked=false;">
    									</td>
    								</tr>
    								<tr>
    								    <% for i = 2 to 8 %>
                                            <td class="content_center"><%= Left(NomeGiorno(IIF(i = 8, 1, i), LINGUA_ITALIANO), 1) %></td>
                                        <% next %>
                                    </tr>
    								<tr>
    								    <% for i = 1 to 7 %>
    									    <td class="content_center">
    										    <input type="checkbox" class="checkbox" id="d_<%= IIF(i = 7, 0, i) %>" name="d_<%= IIF(i = 7, 0, i) %>" <%= chk(request.form("d_" & IIF(i = 7, 0, i))<>"") %> value="1" onclick="form1.d_tutti.checked = false;">
    									    </td>
    								    <% next %>
    								</tr>
    							</table>
    						</td>
    					</tr>
    					<tr>
    						<td class="note">
    							Per selezionare solamente determinati giorni della settimana scegliere l'apposito checkbox.
    						</td>
    					</tr>
    					<tr>
    						<td>
    							<table cellpadding="0" cellspacing="0" class="calendario_opzioni_settimane" width="100%">
    								<tr><th class="l2" colspan="6">Settimane</th></tr>
    								<tr>
    									<td class="content" colspan="7">
    										Seleziona tutte
    										<input type="checkbox" class="checkbox" id="s_tutte" name="s_tutte" value="1" <%= chk((IsPost AND request.form("s_tutte")<>"") OR not IsPost) %>
    											   onclick="this.checked = true; form1.s_1.checked=false;form1.s_2.checked=false;form1.s_3.checked=false;form1.s_4.checked=false;form1.s_5.checked=false;form1.s_6.checked=false;">
    									</td>
    								</tr>
    								<tr>
    								    <% for i = 1 to 6 %>
    									    <td class="content_center"><%= i %>a</td>
    								    <% next %>
    								</tr>
    								<tr>
    								    <% for i = 1 to 6 %>
    									    <td class="content_center">
    										    <input type="checkbox" class="checkbox" id="s_<%= i %>" name="s_<%= i %>" value="1" <%= chk(request.form("s_" & i)<>"") %> onclick="form1.s_tutte.checked = false;">
    									    </td>
    								    <% next %>
    								</tr>
    							</table>
    						</td>
    					</tr>
    					<tr>
    						<td class="note">
    							Per selezionare solamente determinate settimane del mese scegliere i relativi checkbox.
    						</td>
    					</tr>
    				</table>
    			</td>
    			<td class="calendario_menu">
    				<table cellpadding="0" cellspacing="0" width="100%">
    					<tr>
    						<td><input type="submit" name="prec" value="&laquo; precedenti"></td>
    						<td style="text-align: right;">
    							<input type="submit" name="succ" value="successivi &raquo;">
    							<input type="hidden" name="curDate" value="<%= annoInizio &"/"& meseInizio &"/1" %>">
    						</td>
    					</tr>
    				</table>
    			</td>
    		</tr>
    		<tr>
    			<td class="calendari">
                    <%'visualizzazione calendari
    	            for i = meseInizio to meseInizio + NumeroCalendari-1
                        CALL WriteCalendar(((i-1) mod 12) + 1, annoInizio + (i-1) \ 12)
    	            next 
                    
                    'genera parte di form per "ricordare" i giorni selezionati all'esterno dei calendari visibili (giorni non MarkedAsPrinted)
                    for each i in SelectedDays.keys
                        if not SelectedDays_MarkedAsPrinted(i) then %>
                           <input type="hidden" name="f_<%= i %>" value="1"> 
                        <% end if
                    next %>
                </td>
            </tr>
    		<tr>
    			<td class="note" id="istruzioni" colspan="2">
    				1) Per selezionare un periodo scegliere la data di inizio e, tenendo premuto il tasto SHIFT, la data di fine.<br>
    				2) Una volta selezionato un filtro non &egrave; possibile scegliere una data che non vi appartiene.<br>
                    <% if MarkedDays_PrintedCount>0 then %>
                        <%= MarkedDays_Message %>
                    <% end if %>
    			</td>
    		</tr>
    	</table>
    <% end sub
    
    
    'procedura che genera la porzione di codice JavaScript (lo scrive solo la prima volta)
    private sub WriteJsTools()
        
        if not JsWrited then %>
            <script type="text/javascript">
        		var lastSel, inc, key
        		
				//funzione richiamata al click sul giorno
        		function DayOnClick(e, dayInput, dayCell) {
        			var p = new Array(), a = new Array()
        			var dp = new Date(), da = new Date()
					
        			if (window.event){ 				//explorer
        				key = event.shiftKey
                    }
        			else {							//firefox
        				key = e.shiftKey;
        				e.stopPropagation()
        			}
        			
        			//data arrivo
        			a = dayInput.id.split("_")
        			da.setFullYear(a[1], a[2]-1, a[3])
        			
        			if (key) {
        				if (lastSel == null)
        					alert("Scegli una data di inizio")
        				else {
        					// data partenza
        					p = lastSel.id.split("_")
        					dp.setFullYear(p[1], p[2]-1, p[3])
        					
							if (da.toString() != dp.toString()) {
	        					if (da > dp)
	        						inc = 1
	        					else
	        						inc = -1
	        					dp.setDate(dp.getDate() + inc)
	        					
	        					while (dp.toString() != da.toString()) {
	        						DaySelection(document.getElementById("g_"+ dp.getFullYear() +"_"+ FixString(dp.getMonth() + 1) +"_"+ FixString(dp.getDate())),
	        							  		 document.getElementById("t_"+ dp.getFullYear() +"_"+ FixString(dp.getMonth() + 1) +"_"+ FixString(dp.getDate())),
	        							  		 dp.getDay())
	        						
	        						dp.setDate(dp.getDate() + inc)
	        					}
	        					DaySelection(dayInput, dayCell, da.getDay())
							}
        				}
                        //termina selezione multipla
                       lastSel = null;
        			} else
        				DaySelection(dayInput, dayCell, da.getDay())
        			
        			return false
        		}
        		
				//funzione per selezione giorno
        		function DaySelection(dayInput, dayCell, day) {
        			if ((document.getElementById("d_"+ day).checked || form1.d_tutti.checked)
        				&& (document.getElementById(dayCell.name).checked || form1.s_tutte.checked))
        				if (dayInput.value == "0") {
        					dayInput.value = "1"
        					dayCell.className = "selected"
        					lastSel = dayInput
        				} else {
        					dayInput.value = "0"
        					dayCell.className = ""
        					lastSel = dayInput
        					//if (dayInput = lastSel)
        					//	lastSel = null
        				}
        		}
                
                function FixString(str){
                    str = "000" + str;
                    str = str.substring(str.length - 2, str.length);
                    return str;
                }
        	</script>
            <% JsWrited = true
        end if
        
    end sub
    
    
    'procedura che genera il calendario per il mese richiesto, selezionando i giorni presenti in SeletedDays
    public Sub WriteCalendar(mese, anno)
        CALL WriteJsTools()
    
	    dim i, giorno, sett
        dim scelto, evidenziato %>
	    <div class="calendario">
		    <table cellpadding="0" cellspacing="1" class="calendario">
			    <caption><%= NomeMese(mese, LINGUA_ITALIANO) &" "& anno %></caption>
			    <tr class="giorni">
                    <% 'INTESTAZIONE GIORNI
                    for i = 2 to 8 %>
                        <td><%= Left(NomeGiorno(IIF(i = 8, 1, i), LINGUA_ITALIANO), 1) %></td>
                    <% next %>
                </tr>
                <%
                'SPAZI VUOTI all'inizio del calendario
                if Weekday(anno &"/"& mese &"/1") <> vbMonday then
                    response.write "<tr>"
                end if
                for i = 1 to Weekday(anno &"/"& mese &"/1", vbMonday) - 1
                    response.write "<td>&nbsp;</td>"
                next
                
                'GIORNI del mese
                i = 1
                giorno = anno &"/"& mese &"/"& i
                sett = 1
	
                while IsDate(giorno)
                    if Weekday(giorno) = vbMonday then  'inizio settimana
                        response.write "<tr>"
                    end if
                    
                    scelto = SelectedDays_IsSelected(giorno)
                    if scelto then
                        CALL SelectedDays_MarkAsPrinted(giorno)
                    end if
                    evidenziato = MarkedDays_IsMarked(giorno)
                    if evidenziato then %>
                        <td class="<%= MarkedDays_MarkedStyle(giorno) %>">
                        <% CALL MarkedDays_MarkAsPrinted(giorno)
                    else %>
                        <td>
                    <% end if %>
					    <a <%= IIF(scelto, "class=""selected""", "") %>
					       href="javascript:void(0)" name="s_<%= sett %>" id="t_<%= EncodeDate(giorno) %>" onclick="return DayOnClick(event, form1.g_<%= EncodeDate(giorno) %>, this)">
                           <%= i %></a>
                        <input type="hidden" name="g_<%= EncodeDate(giorno) %>" id="g_<%= EncodeDate(giorno) %>" value="<%= IIF(scelto, "1", "0") %>">
                    </td>
                    <%if Weekday(giorno) = vbSunday then    'fine settimana
                        response.write "</tr>"
                        sett = sett + 1
                    end if
		            
                    i = i + 1
                    giorno = anno &"/"& mese &"/"& i
                wend
                
                'SPAZI VUOTI alla fine del mese
                for giorno = Weekday(anno &"/"& mese &"/"& i-1, vbMonday) + 1 to 7
                    response.write "<td>&nbsp;</td>"
                next
                if Weekday(anno &"/"& mese &"/"& i-1) <> vbSunday then
                    response.write "</tr>"
                end if %>
		    </table>
        </div>
    
    <% End Sub

end class

%>