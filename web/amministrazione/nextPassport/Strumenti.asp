<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Strumenti di gestione del sito"
dicitura.puls_new = "INDIETRO A NEXT-PASSPORT"
dicitura.link_new = "Amministratori.asp"
dicitura.scrivi_con_sottosez()


%>
<div id="content">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption>Request Checker</caption>
		<tr>
			<th colspan="2">LOG DEGLI ERRORI DI REQUEST CHECKER</th>
		</tr>
		<tr>
			<td class="label_no_width" style="width:79%;">
				Visualizza il log degli errori rilevati dal modulo di request checker.
			</td>
			<td class="content_center">
				<a class="button_L2_block" href="RequestCheckerStorico.asp" title="Visualizzazione dello storico degli errori." <%= ACTIVE_STATUS %>>
					LOG
				</a>
			</td>
		</tr>		
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
		<caption>Log di sistema</caption>
		<tr>
			<th colspan="2">LOG GENERALE DEL FRAMEWORK</th>
		</tr>
		<tr>
			<td class="label_no_width" style="width:79%;">
				Visualizza il log delle operazioni effettuate all'interno del framework.
			</td>
			<td class="content_center">
				<a class="button_L2_block" href="LogSistemaStorico.asp" title="Visualizzazione dello storico dei log." <%= ACTIVE_STATUS %>>
					LOG
				</a>
			</td>
		</tr>		
	</table>	
	<% if IsNextAim() then %>
		<form action="" method="post" id="form1" name="form1">
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
				<caption>Conversione url</caption>
				<tr>
					<th colspan="2">TRASFORMA URL IN URL ENCODED</th>
				</tr>
				<tr>
					<td class="label_no_width" style="width:79%;">
						Trasforma la stringa indicata in url encoded:<br>
						<input type="text" size="107" class="text" name="url4Conversion" value="<%=request("url4Conversion")%>">
					</td>
					<td class="content_center"><br>
						<input type="submit" value="CONVERTI" id="converti" name="converti" class="button_l2">
					</td>
				</tr>
				<% if request("converti")<>"" and request("url4Conversion")<>"" then %>
					<tr>
						<td colspan="2" class="content">
							<%= Server.urlencode(request("url4Conversion")) %>
						</td>
					</tr>
				<% end if %>
			</table>
		</form>
	<% end if %>
</div>
</body>
</html>