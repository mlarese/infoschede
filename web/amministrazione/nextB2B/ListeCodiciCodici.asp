<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% Server.ScriptTimeOut = 100000 %>
<% response.charset = "UTF-8" 

const LIMITE_ARTICOLI = 8000

dim conn, rs, rsc, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	Session("B2B_SQL_CODICI") = ""
	CALL GotoRecord(conn, rs, Session("B2B_LSTCOD_SQL"), "lstCod_id", "ListeCodiciCodici.asp")
end if


'salva modifiche alla lista codici
dim var, rel_id
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	if request("cerca")="" AND request("tutti")="" then
		'scorre la richiesta per cercare le righe dei codici
		conn.begintrans
		for each var in request.form
			if instr(1,var,"update_",vbTextCompare)>0 then
				rel_id = replace(var, "update_", "")
				sql = "SELECT * FROM gtb_codici WHERE cod_lista_id=" & cIntero(request("ID")) & " AND cod_variante_id=" & rel_id 
				rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
				if not rs.eof then
					if request("cod_codice_" & rel_id)="" then
						rs.delete
					else
						rs("cod_codice") = request("cod_codice_" & rel_id)
					end if
					rs.update
				else
					if request("cod_codice_" & rel_id)<>"" then
						rs.addNew
						rs("cod_codice") = request("cod_codice_" & rel_id)
						rs("cod_variante_id") = rel_id
						rs("cod_lista_id") = request("ID")
						rs.update
					end if
				end if
				rs.close	
			end if
		next
		conn.committrans
		response.redirect "ListeCodiciCodici.asp?ID=" & request("ID")
	end if
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione lista codici - elenco codici"
dicitura.puls_new = "INDIETRO;SCHEDA LISTA"
dicitura.link_new = "ListeCodici.asp;ListeCodiciMod.asp?ID=" & request("ID")
dicitura.scrivi_con_sottosez()  


'imposta ricerca
dim sql_where
if Request.ServerVariables("REQUEST_METHOD")="POST" and (request("cerca")<>"" or request("tutti")<>"") then
		session("B2B_CODICI_LISTE_SEARCHED") = true
		if request("tutti")<>"" then
			CALL SearchSession_Reset("codici_")
			session("B2B_CODICI_LISTE_SEARCHED_ALL") = true
		elseif request("cerca")<>"" then
			session("B2B_CODICI_LISTE_SEARCHED_ALL") = false
			CALL SearchSession_Reset("codici_")
			CALL SearchSession_Set("codici_")
		end if
		sql_where = ""
		'filtra per codice interno
		if Session("codici_codice_int")<>"" then
            sql_where = sql_where & " AND " & SQL_FullTextSearch(Session("codici_codice_int"), "rel_cod_int")
		end if
		'filtra per codice produttore
		if Session("codici_codice_pro")<>"" then
            sql_where = sql_where & " AND " & SQL_FullTextSearch(Session("codici_codice_pro"), "rel_cod_pro")
		end if
		'filtra per codice personalizzato
		if Session("codici_codice_per")<>"" then
            sql_where = sql_where & " AND " & SQL_FullTextSearch(Session("codici_codice_per"), "cod_codice")
		end if
		'filtra per nome
		if Session("codici_nome")<>"" then
			sql_where = sql_where &" AND "& sql_FullTextSearch(Session("codici_nome"), FieldLanguageList("art_nome_"))
		end if
		'filtra per categoria
		if Session("codici_categoria")<>"" then
			sql_where = sql_where & " AND art_tipologia_id IN (" & categorie.FoglieID(Session("codici_categoria")) & " ) "
		end if
		'filtra per marca
		if Session("codici_marchio")<>"" then
			sql_where = sql_where & " AND art_marca_id=" & Session("codici_marchio")
		end if
		
		'filtra per personalizzato
		if Session("codici_personalizzato") = "1" then
			sql_where = sql_where & " AND (ISNULL(cod_codice, '')<>'') "
		elseif Session("codici_personalizzato") = "0" then
			sql_where = sql_where & " AND (ISNULL(cod_codice, '')='') "
		end if
		
		'ricerca per stato a catalogo
		if Session("codici_stato_catalogo")<>"" then
			if not (instr(1, Session("codici_stato_catalogo"), "1", vbTextCompare)>0 AND _
				    instr(1, Session("codici_stato_catalogo"), "0", vbTextCompare)>0 ) then
				if sql_where <>"" then sql_where = sql_where & " AND "
				if instr(1, Session("codici_stato_catalogo"), "1", vbTextCompare)>0 then
					'articolo a catalogo
					sql_where = sql_where & " NOT (" & SQL_IsTrue(conn, "art_disabilitato") & ") "
				elseif instr(1, Session("codici_stato_catalogo"), "0", vbTextCompare)>0 then
					'articolo fuori catalogo
					sql_where = sql_where & SQL_IsTrue(conn, "art_disabilitato")
				end if
			end if
		end if
		
		'ricerca full-text
		if Session("codici_full_text")<>"" then
			sql_where = sql_where &" AND "& SQL_FullTextSearch(Session("codici_full_text"), FieldLanguageList("art_nome_;art_descr_"))
		end if
		
		Session("B2B_WHERE_SQL_CODICI") = sql_where
end if

%>

<script language="JavaScript" type="text/javascript">
	//variabile utilizzata per il controllo del submit nel form di ricerca
	var ClickVediTutti = true;
	
	function verifica_intenzioni(){
		if (ClickVediTutti){
			return window.confirm('ATTENZIONE: se il numero di articoli e\' elevato la pagina potrebbe impiegare alcuni minuti per essere visualizzata. \n' + 
								  'Visualizzare comunque TUTTI gli articoli?')
		}
		else
			return true;
	}
//-->
</script>
<div id="content">
	<table width="100%" cellspacing="0" cellpadding="0" border="0">
		<tr>
		<form action="" method="post" id="ricerca" name="ricerca" onsubmit="return verifica_intenzioni();">
	  		<td width="27%" valign="top">
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>Opzioni di ricerca</caption>
					<tr>
						<td class="footer" colspan="2">
							<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;" onclick="ClickVediTutti=false;">
							<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;" onclick="ClickVediTutti=true;">
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("codici_codice_int;codici_codice_pro;codici_codice_per") %>>CODICI</td></tr>
					<tr>
						<td class="label">interno:</td>
						<td class="content">
							<input type="text" name="search_codice_int" value="<%= TextEncode(Session("codici_codice_int")) %>" style="width:100%;">
						</td>
					</tr>
					<tr>
						<td class="label">produttore:</td>
						<td class="content">
							<input type="text" name="search_codice_pro" value="<%= TextEncode(Session("codici_codice_pro")) %>" style="width:100%;">
						</td>
					</tr>
					<tr>
						<td class="label">personalizzato:</td>
						<td class="content">
							<input type="text" name="search_codice_per" value="<%= TextEncode(Session("codici_codice_per")) %>" style="width:100%;">
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("codici_nome") %>>NOME</td></tr>
					<tr>
						<td class="content" colspan="2">
							<input type="text" name="search_nome" value="<%= TextEncode(Session("codici_nome")) %>" style="width:100%;">
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("codici_stato_catalogo") %>>STATO ARTICOLO A CATALOGO</td></tr>
					<tr>
						<td class="content" style="width:45%;">
							<input type="checkbox" class="checkbox" name="search_stato_catalogo" value="1" <%= chk(instr(1, session("codici_stato_catalogo"), "1", vbTextCompare)>0) %>>
							visibile
						</td>
						<td class="content">
							<input type="checkbox" class="checkbox" name="search_stato_catalogo" value="0" <%= chk(instr(1, Session("codici_stato_catalogo"), "0", vbTextCompare)>0) %>>
							non visibile
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("codici_personalizzato") %>>STATO ARTICOLO NELLA LISTA CODICI</td></tr>
					<tr>
						<td class="content" colspan="2">
							<input type="checkbox" class="checkbox" name="search_personalizzato" value="1" <%= chk(instr(1, session("codici_personalizzato"), "1", vbTextCompare)>0) %>>
							personalizzato
						</td>
					</tr>
					<tr>
						<td class="content" colspan="2">
							<input type="checkbox" class="checkbox" name="search_personalizzato" value="0" <%= chk(instr(1, Session("codici_personalizzato"), "0", vbTextCompare)>0) %>>
							non personalizzato
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("codici_categoria") %>>CATEGORIA</td></tr>
					<tr>
						<td class="content" colspan="2">
							<% CALL categorie.WritePicker("ricerca", "search_categoria", session("codici_categoria"), false, true, 32) %>
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("codici_marchio") %>>MARCHIO</td></tr>
					<tr>
						<td class="content" colspan="2">
							<%	sql = "SELECT * FROM gtb_marche ORDER BY mar_nome_it"
							CALL dropDown(conn, sql, "mar_id", "mar_nome_it", "search_marchio", Session("codici_marchio"), false, " style=""width:100%;""", LINGUA_ITALIANO) %>
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("codici_full_text") %>>FULL-TEXT (tutti i campi)</td></tr>
					<tr>
						<td class="content" colspan="2">
							<input type="text" name="search_full_text" value="<%= TextEncode(Session("codici_full_text")) %>" style="width:100%;">
						</td>
					</tr>
					<tr>
						<td class="footer" colspan="2">
							<input type="submit" name="cerca" value="CERCA" class="button" style="width: 49%;" onclick="ClickVediTutti=false;">
							<input type="submit" class="button" name="tutti" value="VEDI TUTTI" style="width: 49%;" onclick="ClickVediTutti=true;">
						</td>
					</tr>
				</table>
			</td>
			</form>	
			<td width="1%">&nbsp;</td>
			<td valign="top">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px">
					<caption>	
						<table align="right" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td align="right" style="font-size: 1px;">
									<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="lista codici precedente" <%= ACTIVE_STATUS %>>
										&lt;&lt; PRECEDENTE
									</a>
									&nbsp;
									<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="lista codici successiva" <%= ACTIVE_STATUS %>>
										SUCCESSIVA &gt;&gt;
									</a>
								</td>
							</tr>
						</table>
						<% sql = "SELECT LstCod_nome FROM gtb_lista_codici WHERE lstCod_id=" & cIntero(request("ID")) %>
						Modifica dati della lista codici "<%= GetValueList(conn, rs, sql) %>"
					</caption>
					<% if session("B2B_CODICI_LISTE_SEARCHED") = true then
					
						sql = " SELECT * FROM (gtb_articoli a "& _
							  " INNER JOIN grel_art_valori v ON a.art_id = v.rel_art_id) "& _
							  " LEFT JOIN gtb_codici c ON (v.rel_id = c.cod_variante_id AND cod_lista_id="& cIntero(request("ID")) &") "& _
							  " WHERE (1=1) "
						if cString(session("B2B_WHERE_SQL_CODICI")) <> "" then
							sql = sql & session("B2B_WHERE_SQL_CODICI")
						end if
						rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
						if rs.eof then	%>
							<tr>
								<td class="noRecords">Nessun articolo trovato</td>
							</tr>
						<% elseif rs.recordcount > LIMITE_ARTICOLI then %>
							<tr><th colspan="9">Elenco codici della lista</th></tr>
							<tr>
								<td class="label_no_width" colspan="5">
									Selezionati n&ordm; <%= rs.recordcount %> articoli
								</td>
							</tr>
							<tr>
								<td class="noRecords">Raggiunto limite massimo di articoli modificabili.<br>Selezionare un massimo di <%= LIMITE_ARTICOLI %> articoli.</td>
							</tr>
						<% else %>
							<tr><th colspan="9">Elenco codici della lista</th></tr>
							<tr>
								<td class="label" colspan="3">
									Selezionati n&ordm; <%= rs.recordcount %> articoli
								</td>
								<td class="content_right" colspan="2">
									<a class="button_L2" href="javascript:void(0);" onclick="form1.reset();"
									   title="annulla tutte le modifiche apportate ai dati ma non ancora salvate" <%= ACTIVE_STATUS %>>
										ANNULLA MODIFICHE
									</a>
								</td>
							</tr>
						<form action="" method="post" id="form1" name="form1">
							<tr>
								<th class="l2_center" rowspan="2" style="border-bottom:0px; width:36%;">Articolo</th>
								<th class="l2_center" colspan="3" style="border-bottom:0px;">codici</th>
								<th class="l2_center" rowspan="2" style="width: 10%;">SALVA</th>
							</tr>
							<tr>
								<th class="l2_center" style="width:17%;">INTERNO</th>
								<th class="l2_center" style="width:17%;">PRODUTTORE</th>
								<th class="l2_center" style="width:20%;">PERSONALIZZATO</th>
							</tr>
							<% while not rs.eof 
								if rs.absoluteposition mod 500 = 0 then
									response.flush
								end if%>
								<tr>
									<td class="content">
										<% CALL ArticoloLink(rs("rel_art_id"), rs("art_nome_it"), rs("rel_cod_int"))
										if rs("art_varianti") then %>
											<%= ListValoriVarianti(conn, rsc, rs("rel_id")) %>
										<% end if %>
									</td>
									<td class="content"><%=rs("rel_cod_int") %></td>
									<td class="content"><%= rs("rel_cod_pro") %></td>
									<td class="content">
										<input type="text" class="text" name="cod_codice_<%= rs("rel_id") %>" value="<%= rs("Cod_Codice") %>" size="15" maxlength="50">
									</td>
									<td class="Content_center">
										<input type="submit" class="button_L2" name="update_<%= rs("rel_id") %>" value="SALVA">
									</td>
								</tr>
								<%rs.movenext
							wend%>
							<tr>
								<td class="footer" colspan="9">
									<input type="Hidden" name="B2B_SQL_CODICI" value ="<%= Session("B2B_SQL_CODICI") %>">
									<input type="reset" class="button" name="annulla" value="ANNULLA MODIFICHE">
									<input type="submit" class="button" name="salva" value="SALVA TUTTI">
								</td>
							</tr>
						</form>
						<% end if
						rs.close
					else%>
						<tr>
							<td class="noRecords">Per visualizzare l'elenco degli articoli eseguire prima una ricerca.</td>
						</tr>
					<% end if %>
				</table>
			</td> 
		</tr>
	</table>	
</div>
</body>
</html>
<% 
conn.close 
set rs = nothing
set rsc = nothing
set conn = nothing
%>