<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%

dim conn, rs, sql, rsv, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsv = Server.CreateObject("ADODB.RecordSet")


if request("salva")<>"" AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	dim dal, al, data_fine
	Session("errore") = ""
	
	'controllo gli orari
	dal = DATE() & " " & request("orario_inizio") & ":00"
	al = DATE() & " " & request("orario_fine") & ":00"
	if DateDiff("n",dal,al) <= 0 then
		Session("errore") = "ATENZIONE! Orario di fine coincidente o precedente all'orario di inizio."
	end if
	
	
	dal = request("data_inizio") & " " & request("orario_inizio") & ":00"
	
	data_fine = cString(request("data_fine"))
	if data_fine = "" then data_fine = DATA_SENZA_FINE
	al = data_fine & " " & request("orario_fine") & ":00"
	
	'controllo le date
	if DateDiff("d",dal,al) < 0 then
		Session("errore") = "ATENZIONE! Data di fine precedente alla data di inizio."
	end if
	
	
	
	if cIntero(request("coi_id")) = 0 then
		sql = " INSERT INTO mtb_configurazione_impegni(coi_giorno, coi_dal, coi_al)" & _
			  " VALUES ('"&ParseSql(request("tft_coi_giorno"),adChar)&"', "&SQL_DateTime(conn, dal)&", "&SQL_DateTime(conn, al)&")"
	else
		sql = " UPDATE mtb_configurazione_impegni SET coi_giorno = '"&ParseSql(request("tft_coi_giorno"),adChar)&"', " & _
			  "		coi_dal = "&SQL_DateTime(conn, dal)&", coi_al = "&SQL_DateTime(conn, al) & _
			  " WHERE coi_id = " & cIntero(request("ID"))
	end if
	
	if Session("ERRORE") = "" then
		conn.Execute(sql)
		response.redirect "AgendaConfigura.asp"
	end if
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 

dicitura.iniz_sottosez(2)
dicitura.sottosezioni(1) = "IMPEGNI"
dicitura.links(1) = "Impegni.asp"
dicitura.sottosezioni(2) = "TIPOLOGIE"
dicitura.links(2) = "ImpegniTipologie.asp"

dicitura.sezione = "Gestione impegni/appuntamenti - configura"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Impegni.asp;"
dicitura.scrivi_con_sottosez() 

dim intervallo
'intervallo = cIntero(Session("AGENDA_INTERVALLO_CALENDARIO"))
'if intervallo = "" then
	intervallo = 5
'end if

sql = "SELECT * FROM mtb_configurazione_impegni ORDER BY coi_giorno, coi_dal, coi_al"
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Configurazione dell'agenda degli impegni/appuntamenti</td>
				</tr>
			</table>
		</caption>
		<tr>
			<th style="width:11%;">GIORNO</th>
			<th style="width:8%;">ORA INIZIO</th>
			<th style="width:8%;">ORA FINE</th>
			<th style="width:20%;">DATA INIZIO</th>
			<th>DATA FINE</th>
			<th style="text-align:center; width:22%;">OPERAZIONI</th>
		</tr>
		<% while not rs.eof %>
			<% if cIntero(rs("coi_id")) = cIntero(request("ID")) then %>
				<!-- IN MODIFICA -->
				<input type="hidden" name="coi_id" value="<%=cIntero(rs("coi_id"))%>">
				<tr>
					<td class="content">
						<% CALL WriteDropDownGiorno("tft_coi_giorno",rs("coi_giorno"),true) %>
					</td>
					<td class="content">
						<%CALL WriteDropDownOrario("orario_inizio",intervallo, TimeIta(Hour(rs("coi_dal"))&"."&Minute(rs("coi_dal"))),"","")%>
					</td>
					<td class="content">
						<%CALL WriteDropDownOrario("orario_fine",intervallo, TimeIta(Hour(rs("coi_al"))&"."&Minute(rs("coi_al"))),"","")%>
					</td>
					<td class="content">
						<% CALL WriteDataPicker_Input("form1", "data_inizio", DateIta(rs("coi_dal")), "", "/", false, true, LINGUA_ITALIANO) %>
					</td>
					<td class="content">
						<% CALL WriteDataPicker_Input("form1", "data_fine", IIF(Trim(DateIta(rs("coi_al")))<>Trim(DateIta(DATA_SENZA_FINE)),DateIta(rs("coi_al")),""), "", "/", true, true, LINGUA_ITALIANO) %>
					</td>
					<td class="content" style="text-align:center; font-size:1px;">
						<input type="submit" class="button" style="width:61px;" name="salva" value="SALVA">
						&nbsp;
						<a class="button" href="AgendaConfigura.asp">
							ANNULLA
						</a>
					</td>
				</tr>
			<% else %>
				<tr>
					<td class="content"><%=NomeGiorno(rs("coi_giorno"), LINGUA_ITALIANO)%></td>
					<td class="content"><%=TimeIta(Hour(rs("coi_dal"))&"."&Minute(rs("coi_dal")))%></td>
					<td class="content"><%=TimeIta(Hour(rs("coi_al"))&"."&Minute(rs("coi_al")))%></td>
					<td class="content"><%=DateIta(rs("coi_dal"))%></td>
					<td class="content">
						<% if DateISO(rs("coi_al")) = DateISO(DATA_SENZA_FINE) then
							response.write "&nbsp;"
						  else
							response.write DateIta(rs("coi_al"))
						  end if
						%>						
					</td>
					<td class="content" style="text-align:center; font-size:1px;">
						<a class="button" href="AgendaConfigura.asp?ID=<%= rs("coi_id") %>">
							MODIFICA
						</a>
						&nbsp;
						<a class="button" href="javascript:void(0);" onclick="OpenDeleteWindow('AGENDA_CONFIGURA','<%= rs("coi_id") %>');" >
							CANCELLA
						</a>
					</td>
				</tr>
			<% end if %>
			<% rs.moveNext %>
		<% wend %>
		
		<% if request("ID") = "" then %>
			<!-- AGGIUNTA -->
			<tr>
				<td class="content">
					<% CALL WriteDropDownGiorno("tft_coi_giorno",request("tft_coi_giorno"),true) %>
				</td>
				
				<td class="content">
					<%CALL WriteDropDownOrario("orario_inizio",intervallo,request("orario_inizio"),"","")%>
				</td>
				<td class="content">
					<%CALL WriteDropDownOrario("orario_fine",intervallo,request("orario_fine"),"","")%>
				</td>
			
				<td class="content">
					<% CALL WriteDataPicker_Input("form1", "data_inizio", IIF(request("data_inizio")<>"",request("data_inizio"),DateValue(Now())), "", "/", false, true, LINGUA_ITALIANO) %>
				</td>
				<td class="content">
					<% CALL WriteDataPicker_Input("form1", "data_fine", request("data_fine"), "", "/", true, true, LINGUA_ITALIANO) %>
				</td>
				<td class="content" style="text-align:center;">
					<input type="submit" class="button" style="width:140px;" name="salva" value="AGGIUNGI">
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="footer" colspan="6">
				&nbsp;
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>


<% rs.close
conn.close
set rs = nothing
set conn = nothing%>
