<% 	'Scrive le righe sul box ricerca per la gestione del salvataggio della query dei contatti salvata in sessione
	'sessionSQL:		nome della variabile in sessione che memorizza la query su tb_indirizzario
	'tabella:			nome della tabella esterna (verra inserita in syncroTable)
	Sub SalvaInRubrica(rs, sessionSQL, tabella) %>
	<tr><td style="font-size:4px;">&nbsp;</td></tr>
	<tr>
		<td>
			<% 	'visualizzo le rubriche salvate
				rs.open "SELECT * FROM tb_rubriche WHERE syncroTable='"& ParseSql(tabella, adChar) &"' ORDER BY nome_rubrica", conn, adOpenStatic, adLockReadOnly %>
			<table cellspacing="1" cellpadding="0" class="tabella_madre">
				<caption>Operazioni sui risultati</caption>
				<tr><th colspan="2">SALVATI COME RUBRICHE</th></tr>
				<tr>
					<td colspan="2" class="label" style="width:100%;">
						trovati n&deg; <%= rs.recordCount %> records
					</td>
				</tr>
				<tr>
					<td class="content_right" colspan="2">
						<a class="button_L2" href="javascript:void(0);" onclick="OpenAutoPositionedWindow('../NEXTcom/RubricheExport.asp?tabella=<%= tabella %>&sql=<%= sessionSQL %>', 'rubriche', 400, 100)">
							SALVA IN UNA NUOVA RUBRICA
						</a>
					</td>
				</tr>
				<tr>
					<th class="L2">NOME</th>
					<th class="l2_center" style="width:8%;">OPERAZIONI</th>
				</tr>
			<% 	while not rs.eof %>
				<tr>
					<td class="label"><%= rs("nome_rubrica") %></td>
					<td class="content_right">
						<a class="button_L2" href="javascript:void(0);" onclick="OpenAutoPositionedWindow('../NEXTcom/RubricheExport.asp?tabella=<%= tabella %>&sql=<%= sessionSQL %>&SALVA=<%= rs("id_rubrica") %>', 'rubriche', 400, 200)">
							AGGIORNA
						</a><br>
						<a class="button_L2" href="javascript:void(0);" onclick="OpenDeleteWindow('RUBRICHE','<%= rs("id_rubrica") %>');">
							CANCELLA
						</a>
					</td>
				</tr>
			<%		rs.movenext
				wend
				rs.close %>
				<tr><th colspan="2">ESPORTATI</th></tr>
				<tr>
					<td class="content_right" colspan="2">
						<a style="width:100%; text-align:center; line-height:12px;" class="button"
							title="Apre la palette di export dei dati" 
							onclick="OpenAutoPositionedScrollWindow('../NEXTcom/ContattiExport.asp?sql=<%= sessionSQL %>', 'export', 240, 142, true);" href="javascript:void(0);">
							ESPORTA ELENCO CLIENTI
						</a>
					</td>
				</tr>
			</table>
		</td>
	</tr>
<% 	End Sub %>