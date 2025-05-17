
<%
dim i
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>Inserimento nuova tipologia</caption>
		<tr><th colspan="2">DATI DELLA TIPOLOGIA</th></tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 	if i = 0 then %>
			<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">nome:</td>
		<% 	end if %>
			<td class="content">
				<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
				<input type="text" class="text" name="tft_dot_nome_<%= Application("LINGUE")(i) %>" value="<%= request("tft_dot_nome_"& Application("LINGUE")(i)) %>" maxlength="250" size="75">
				<% 	if Application("LINGUE")(i) = LINGUA_ITALIANO then response.write "(*)" end if %>
			</td>
		</tr>
		<%next %>
        <tr>
			<td class="label" >codice:</td>
			<td class="content">
				<input type="text" class="text" name="tft_dot_codice" value="<%= request("tft_dot_codice") %>" maxlength="250" size="26">
			</td>
		</tr>
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
		<tr>
		<% 	if i = 0 then %>
			<td class="label" rowspan="<%= ubound(Application("LINGUE"))+1 %>">descrizione:</td>
		<% 	end if %>
			<td class="content">
				<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
				<input type="text" class="text" name="tft_dot_descrizione_<%= Application("LINGUE")(i) %>" value="<%= request("tft_descrizione_nome_"& Application("LINGUE")(i)) %>" maxlength="250" size="75">
			</td>
		</tr>
		<%next %>
		<tr><th colspan="2">INFORMAZIONI PER RIGA D'ORDINE</th></tr>
		<tr><td class="note" colspan="2">L'associazione delle informazioni alle tipologie di righe pu&ograve; essere fatta solo dopo aver salvato.</td></tr>
		<tr>
			<td class="footer" colspan="2">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA &gt;&gt;">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>