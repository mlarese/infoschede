<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="TOOLS.asp" -->
<% 

'impostazione etichette in base alla lingua
dim LblAnnoPrec, LblAnnoSucc, LblMesePrec, LblMeseSucc, Title
Select case request("LINGUA")
	case LINGUA_INGLESE
		Title = "Choose the date..."
		LblAnnoPrec = "Previous year"
		LblAnnoSucc = "Next year"
		LblMesePrec = "Previous month"
		LblMeseSucc = "Next month"
	case LINGUA_FRANCESE
		Title = "Choisissez la date..."
		LblAnnoPrec = "ann&eacute;e pr&eacute;c&eacute;dent"
		LblAnnoSucc = "ann&eacute;e prochain"
		LblMesePrec = "mois pr&eacute;c&eacute;dent"
		LblMeseSucc = "mois prochain"
	case LINGUA_TEDESCO
		Title = "W&auml;hlen Sie das Datum"
		LblAnnoPrec = "vorhergehender jahr"
		LblAnnoSucc = "folgender jahr"
		LblMesePrec = "vorhergehender monat"
		LblMeseSucc = "folgender monat"
	case LINGUA_SPAGNOLO
		Title = "Elija la fecha..."
		LblAnnoPrec = "a&ntilde;o anterior"
		LblAnnoSucc = "a&ntilde;o pr&oacute;ximo"
		LblMesePrec = "mes anterior"
		LblMeseSucc = "mes pr&oacute;ximo"
	case else
		'LINGUA ITALIANO
		Title = "Scegli la data..."
		LblAnnoPrec = "Anno precedente"
		LblAnnoSucc = "Anno successivo"
		LblMesePrec = "Mese precedente"
		LblMeseSucc = "Mese successivo"
end select

%>
<!DOCTYPE html>
<html>
<head>
	<title><%= Title %></title>
	<link rel="stylesheet" type="text/css" href="<%= request("stili") %>">
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" >
</head>
<body class="PickerDate" onload="window.focus()" rightmargin="0" leftmargin="3" topmargin="5">
<script language="JavaScript" type="text/javascript">

function millisec( gg ) {
	var one_day = 1000*60*60*24;
	return (gg * one_day);
}

function tomorrow(s) {
	var data = new Date();
	var temp = s.split("/");
	data.setDate(temp[0]);
	data.setMonth(temp[1]-1);
	data.setYear(temp[2]);
	return new Date( millisec(1) + data.valueOf()  );
}

function printItalianDate( data ) {
	var gg,mm,aaaa;
	gg = data.getDate();
	mm = 1+data.getMonth();
	mm = mm.toString();
	if (mm.length == 1) mm = "0" + mm;
	aaaa = data.getFullYear();
	return (gg + "/" + mm + "/" + aaaa );
}


function select_data(data) {
	if (opener.document.<%= request("form") %>) {		// form asp
		opener.document.<%= request("form") %>.<%= request("input") %>.value = data;
	<% If request.querystring("campoAggiorna") <> "" then %>
		opener.document.<%= request("form") %>.<%= request.querystring("campoAggiorna") %>.value = printItalianDate(tomorrow(data));
	<% End If %>
	} else {											// form aspx
		opener.document.getElementById("<%= request("input") %>").value = data;
		<% If request.querystring("campoAggiorna") <> "" then %>
		opener.document.getElementById("<%= request.querystring("campoAggiorna") %>").value = printItalianDate(tomorrow(data));
		<% End If %>
	}
	<% If request.querystring("nameFunctionAfterClick") <> "" then %>
		opener.<%= request.querystring("nameFunctionAfterClick") %>;
	<% End If %>
	close();
}
	
	
</script>

<%
dim i, selected_date, cur_month, cur_year, first_day, last_day, start_displacement, dayCount, dayCurrent, cur_date
dim start_position, Row_Position, BASE_HREF, DayClass, InputDate

dim iv
iv = trim(request("inputvalue"))
'inizializza variabili
if request("data")<>"" AND isDate(request("data")) then
	selected_date = cDate(request("data"))
elseif iv<>"" then
	selected_date = cDate(DateEng(iv))
else
	selected_date = Date()
end if
if iv<>"" then
	InputDate = cDate(iv)
else
	InputDate = ""
end if

BASE_HREF = "PickerDate.asp?lingua=" & request("LINGUA") & "&form=" & request("form") & "&input=" & request("input") &_
			"&AllowPast=" & request("AllowPast")
If request.querystring("campoAggiorna") <> "" then
	BASE_HREF = BASE_HREF & "&campoAggiorna="& request.querystring("campoAggiorna")
end if
if request.querystring("nameFunctionAfterClick") <> "" then
	BASE_HREF = BASE_HREF & "&nameFunctionAfterClick="& request.querystring("nameFunctionAfterClick")
end if
BASE_HREF = BASE_HREF &"&inputvalue=" & iv & "&stili=" & Server.URLEncode(request("stili")) & "&data="


cur_year = Year(selected_Date)
cur_month = Month(selected_Date)

'calcola primo e ultimo giorno del mese
first_day = DateSerial(cur_year, cur_month, 1)
last_day = DateAdd("d", -1, DateAdd("m", 1, first_day))
dayCount = DateDiff("d", first_day, last_day) + 1

start_displacement = -1		'sposta il giorno di inizio per la settimana italiana
start_position = WeekDay(first_day) + start_displacement
%>
<table width="98%" border="0" cellspacing="1" cellpadding="0" align="center" class="pickerdate">
	<caption class="pickerdate">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<!-- SCORRIMENTO PER ANNO -->
			<tr>
				<td align="center">
					<a title="<%= LblAnnoPrec %>" <%= ACTIVE_STATUS %> href="<%= BASE_HREF %><%= DateIso(DateSerial(Cur_Year-1, cur_month, 1)) %>" class="pickerdate_previous">
						&lt;&lt;
					</a>
				</td>
				<td class="pickerdate_header"><%= cur_year %></td>
				<td align="center">
					<a title="<%= LblAnnoSucc %>" <%= ACTIVE_STATUS %> href="<%= BASE_HREF %><%= DateIso(DateSerial(Cur_Year+1, cur_month, 1)) %>" class="pickerdate_next">
						&gt;&gt;
					</a>
				</td>
			</tr>
<!-- SCORRIMENTO PER MESE -->
			<tr>
				<td align="center">
					<a title="<%= LblMesePrec %>" <%= ACTIVE_STATUS %> href="<%= BASE_HREF %><%= DateIso(DateAdd("m", -1, selected_date)) %>" class="pickerdate_previous">
						&lt;&lt;
					</a>
				</td>
				<td class="pickerdate_header"><%= NomeMese(cur_month, request("LINGUA")) %></td>
				<td align="center">
					<a title="<%= LblMeseSucc %>" <%= ACTIVE_STATUS %> href="<%= BASE_HREF %><%= DateIso(DateAdd("m", 1, selected_date)) %>" class="pickerdate_next">
						&gt;&gt;
					</a>
				</td>
			</tr>
		</table>
	</caption>
<!-- TESTATA GIORNI -->
	<tr>
		<th class="pickerdate"><%= Ucase(Left(NomeGiorno(vbMonday, request("LINGUA")), 1)) %></th>
		<th class="pickerdate"><%= Ucase(Left(NomeGiorno(vbTuesday, request("LINGUA")), 1)) %></th>
		<th class="pickerdate"><%= Ucase(Left(NomeGiorno(vbWednesday, request("LINGUA")), 1)) %></th>
		<th class="pickerdate"><%= Ucase(Left(NomeGiorno(vbThursday, request("LINGUA")), 1)) %></th>
		<th class="pickerdate"><%= Ucase(Left(NomeGiorno(vbFriday, request("LINGUA")), 1)) %></th>
		<th class="pickerdate"><%= Ucase(Left(NomeGiorno(vbSaturday, request("LINGUA")), 1)) %></th>
		<th class="pickerdate"><%= Ucase(Left(NomeGiorno(VbSunday, request("LINGUA")), 1)) %></th>
	</tr>
<!-- GENERAZIONE DATI -->
	<tr>
		<%'celle vuote in apertura
		if start_position = 0 then start_position = 7
		
		for i = 1 to start_position-1 %>
			<td class="pickerdate">&nbsp;</td>
		<% next 
		
		Row_Position = start_position
		dayCurrent = 1
		for dayCurrent = 1 to dayCount
			cur_date = DateSerial(cur_year, cur_month, dayCurrent)
			if Row_Position = 8 then%>
				</tr>
				<tr>
				<%Row_position = 1
			end if
			if cur_Date = InputDate then
				DayClass="pickerdate_selected"
			elseif cur_Date < Date then
				DayClass="pickerdate_yesterday"
			elseif cur_date = Date then
				DayClass="pickerdate_today"
			else
				DayClass="pickerdate_tomorrow"
			end if %>
			<td class="pickerdate" title="<%= DataEstesa(cur_date, request("LINGUA")) %>">
				<% if request("AllowPast")="" AND cur_Date < Date then %>
					<a class="<%= DayClass %>" >
						<%= dayCurrent %>
					</a>
				<% else %>
					<a class="<%= DayClass %>" href="javascript:void(0);" onclick="select_data('<%= DateIta(cur_date) %>')" title="<%= DataEstesa(cur_date, request("LINGUA")) %>" <%= ACTIVE_STATUS %>>
						<%= dayCurrent %>
					</a>
				<% end if %>
			</td>
			<%Row_Position = Row_Position + 1
		next
		
		'celle vuote in chiusura
		for i = Row_Position to 7 %>
			<td class="pickerdate">&nbsp;</td>
		<% next%>
	</tr>
</table>

</body>
</html>