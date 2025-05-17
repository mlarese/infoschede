<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<% 

CALL CheckAutentication(Session("LOGIN_4_LOG") <> "")


dim conn, rs, rsr, rsv, sql, rubriche_visibili, i, var, testo, sqlv

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

'recupera rubriche visibili all'utente
rubriche_visibili = GetList_Rubriche(conn, rs)

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	'azzera tutte le variabili di ricerca
	for each var in Session.Contents
		if instr(1, var, "search_", vbTextCompare)>0 then
			Session(var) = ""
		end if
	next
	
	'imposta variabili di ricerca (ricerca avanzata)
	i = 0
	for each var in request.form
		if instr(1, var, "ADV_search_", vbTextCompare)>0 AND cString(request(var))<>"" then
			Session(var) = request(var)
			i = i + 1
		end if
	next
	
	if i>0 then	'controlla se &egrave; stata impostata una ricerca
		'imposta criteri per ricerca avanzata
		testo = ""
		sql = " SELECT * FROM tb_indirizzario INNER JOIN tb_cnt_lingue ON tb_indirizzario.lingua=tb_cnt_lingue.lingua_codice WHERE " & _
              " (CntRel = 0 OR "  & SQL_IsNull(conn, "CntRel") & ") AND " & _
		      " IDElencoIndirizzi IN (SELECT id_indirizzo FROM rel_rub_ind WHERE id_rubrica IN ("
		'filtra sulle rubriche
		if Session("ADV_search_rubriche")<>"" then
			sql = sql & Session("ADV_search_rubriche") & ")) "
			if session("ADV_search_rubriche_tipo")<>"" then
				'addociato a tutte le rubriche
				sql = sql & " AND (SELECT COUNT(*) FROM rel_rub_ind WHERE id_indirizzo=tb_indirizzario.idElencoIndirizzi AND id_rubrica IN (" & _
							Session("ADV_search_rubriche") & "))=" & (1 + count(Session("ADV_search_rubriche"), ","))
				testo = testo & "<tr><td class=""label"">associati a tutte le rubriche:</td></tr>"
			else
				testo = testo & "<tr><td class=""label"">associati ad almeno una rubrica:</td></tr>"
			end if
			
			'recupera valori scelti da visualizzare
			sqlv = "SELECT nome_rubrica FROM tb_rubriche WHERE id_rubrica IN (" & _
					Session("ADV_search_rubriche") & ") ORDER BY nome_rubrica "
			testo = testo & "<tr><td class=""content_right"">" & GetValueList(conn,rs, sqlv) & "</td></tr>"
		else
			sql = sql & rubriche_visibili & ")) "
		end if
		
		if Session("ADV_search_iniziali")<>"" then
			sql = sql & " AND " & SQL_Ucase(conn) & "(LEFT(ModoRegistra, 1)) IN (" & Session("ADV_search_iniziali") & ")"
			testo = testo & "<tr><td class=""label"">che inizino per</td></tr><tr><td class=""content_right"">" & Session("ADV_search_iniziali") & "</td></tr>"
		end if

		
		if Session("ADV_search_denominazione")<>"" then
			sql = sql & " AND " + SQL_FullTextSearch_contatto_nominativo(conn, Session("ADV_search_denominazione")) & _
						" OR IDElencoIndirizzi IN (SELECT cntRel FROM  tb_indirizzario WHERE " + SQL_FullTextSearch_contatto_nominativo(conn, Session("ADV_search_denominazione")) &")"
			testo = testo & "<tr><td class=""label"">nella cui anagrafica compaia:</td></tr><tr><td class=""content_right"">" & Session("ADV_search_denominazione") & "</td></tr>"
		end if

		if Session("ADV_search_denominazione_tipo")<>"" AND _
			NOT(instr(1, Session("ADV_search_denominazione_tipo"), "0", vbTextCompare)>0 AND _
				instr(1, Session("ADV_search_denominazione_tipo"), "1", vbTextCompare)>0 )then
			if instr(1, Session("ADV_search_denominazione_tipo"), "1", vbTextCompare)>0 then
				'societa / ente / organizzazione
				sql = sql & " AND " & SQL_isTrue(conn, "isSocieta")
				testo = testo & "<tr><td class=""label"">salvati come:</td></tr><tr><td class=""content_right"">societ&agrave; / ente / organizzazione</td></tr>"
			elseif instr(1, Session("ADV_search_denominazione_tipo"), "0", vbTextCompare)>0 then
				'persona fisica
				sql = sql & " AND NOT(" & SQL_isTrue(conn, "isSocieta") & ") "
				testo = testo & "<tr><td class=""label"">salvati come:</td></tr><tr><td class=""content_right"">persona fisica</td></tr>"
			end if
		end if
		
		'ricerca per lingua
		if Session("ADV_search_lingua")<>"" then
			sql = sql & " AND lingua LIKE '" & Session("ADV_search_lingua") & "' "
			testo = testo & "<tr><td class=""label"">lingua:</td></tr><tr><td class=""content_right"">" & GetNomeLingua(Session("ADV_search_lingua")) & "</td></tr>"
		end if

		if Session("ADV_search_indirizzo")<>"" then
			sql = sql & " AND " + SQL_FullTextSearch_contatto_indirizzo(conn, Session("ADV_search_indirizzo"))
			testo = testo & "<tr><td class=""label"">nel cui indirizzo compaia:</td></tr><tr><td class=""content_right"">" & Session("ADV_search_indirizzo") & "</td></tr>"
		end if
		
		if Session("ADV_search_citta")<>"" then
			sql = sql & "AND " + SQL_FullTextSearch(Session("ADV_search_citta"), "CittaElencoIndirizzi")
			testo = testo & "<tr><td class=""label"">nella citt&agrave;:</td></tr><tr><td class=""content_right"">" & Session("ADV_search_citta") & "</td></tr>"
		end if
		
		if Session("ADV_search_indirizzo_completo")<>"" AND _
			NOT(instr(1, Session("ADV_search_indirizzo_completo"), "0", vbTextCompare)>0 AND _
				instr(1, Session("ADV_search_indirizzo_completo"), "1", vbTextCompare)>0 )then
			testo = testo & "<tr><td class=""label"">con indirizzo postale:</td></tr><tr>"
			if instr(1, Session("ADV_search_indirizzo_completo"), "1", vbTextCompare)>0 then
				'con indirizzo postale completo
				sql = sql & " AND (IndirizzoElencoIndirizzi<>'' AND CittaElencoIndirizzi<>'' AND CapElencoIndirizzi<>'') "
				testo = testo & "<td class=""content_right"">completo</td></tr>"
			elseif instr(1, Session("ADV_search_indirizzo_completo"), "0", vbTextCompare)>0 then
				'con indirizzo postale incompleto
				sql = sql & " AND (IndirizzoElencoIndirizzi='' OR CittaElencoIndirizzi='' OR CapElencoIndirizzi='') "
				testo = testo & "<td class=""content_right"">non completo</td></tr>"
			end if
		end if
		
		if session("ADV_search_telefono")<>"" then
			'cerca una stringa tra i valori del telefono
			sql = sql & " AND IDElencoIndirizzi IN (SELECT id_Indirizzario FROM tb_ValoriNumeri " &_
						" WHERE " + SQL_FullTextSearch(session("ADV_search_telefono"), "ValoreNumero") + " AND (id_TipoNumero<=4))"
			testo = testo & "<tr><td class=""label"">numero di telefono:</td></tr><tr><td class=""content_right"">" & Session("ADV_search_telefono") & "</td></tr>"
		elseif Session("ADV_search_telefono_tipo")<>"" AND _
			NOT(instr(1, Session("ADV_search_telefono_tipo"), "0", vbTextCompare)>0 AND _
				instr(1, Session("ADV_search_telefono_tipo"), "1", vbTextCompare)>0 )then
			if instr(1, Session("ADV_search_telefono_tipo"), "1", vbTextCompare)>0 then
				'con numero di telefono
				sql = sql & " AND IDElencoIndirizzi IN (SELECT id_Indirizzario FROM tb_ValoriNumeri WHERE id_TipoNumero<=4) "
				testo = testo & "<tr><td class=""label"">numero di telefono:</td></tr><tr><td class=""content_right"">con numero di telefono</td></tr>"
			elseif instr(1, Session("ADV_search_telefono_tipo"), "0", vbTextCompare)>0 then
				'senza numero di telelfono
				sql = sql & " AND IDElencoIndirizzi NOT IN (SELECT id_Indirizzario FROM tb_ValoriNumeri WHERE id_TipoNumero<=4) "
				testo = testo & "<tr><td class=""label"">numero di telefono:</td></tr><tr><td class=""content_right"">senza numero di telefono</td></tr>"
			end if
		end if
		
		if session("ADV_search_fax")<>"" then
			'cerca una stringa tra i valori del fax
			sql = sql & " AND IDElencoIndirizzi IN (SELECT id_Indirizzario FROM tb_ValoriNumeri " &_
						" WHERE " + SQL_FullTextSearch(session("ADV_search_fax"), "ValoreNumero") + " AND (id_TipoNumero<=5))"
			testo = testo & "<tr><td class=""label"">numero di fax:</td></tr><tr><td class=""content_right"">" & Session("ADV_search_fax") & "</td></tr>"
		elseif Session("ADV_search_fax_tipo")<>"" AND _
			NOT(instr(1, Session("ADV_search_fax_tipo"), "0", vbTextCompare)>0 AND _
				instr(1, Session("ADV_search_fax_tipo"), "1", vbTextCompare)>0 )then
			if instr(1, Session("ADV_search_fax_tipo"), "1", vbTextCompare)>0 then
				'con numero di fax
				sql = sql & " AND IDElencoIndirizzi IN (SELECT id_Indirizzario FROM tb_ValoriNumeri WHERE id_TipoNumero=5) "
				testo = testo & "<tr><td class=""label"">numero di fax:</td></tr><tr><td class=""content_right"">con numero di fax</td></tr>"
			elseif instr(1, Session("ADV_search_fax_tipo"), "0", vbTextCompare)>0 then
				'senza numero di fax
				sql = sql & " AND IDElencoIndirizzi NOT IN (SELECT id_Indirizzario FROM tb_ValoriNumeri WHERE id_TipoNumero=5) "
				testo = testo & "<tr><td class=""label"">numero di fax:</td></tr><tr><td class=""content_right"">senza numero di fax</td></tr>"
			end if
		end if
		
		if session("ADV_search_email")<>"" then
			'cerca una stringa tra le email
			sql = sql & " AND IDElencoIndirizzi IN (SELECT id_Indirizzario FROM tb_ValoriNumeri " &_
						" WHERE " + SQL_FullTextSearch(session("ADV_search_email"), "ValoreNumero") + " AND (id_TipoNumero<=6))"
			testo = testo & "<tr><td class=""label"">email:</td></tr><tr><td class=""content_right"">" & Session("ADV_search_email") & "</td></tr>"	
		elseif Session("ADV_search_email_tipo")<>"" AND _
			NOT(instr(1, Session("ADV_search_email_tipo"), "0", vbTextCompare)>0 AND _
				instr(1, Session("ADV_search_email_tipo"), "1", vbTextCompare)>0 )then
			if instr(1, Session("ADV_search_email_tipo"), "1", vbTextCompare)>0 then
				'con email
				sql = sql & " AND IDElencoIndirizzi IN (SELECT id_Indirizzario FROM tb_ValoriNumeri WHERE id_TipoNumero=6) "
				testo = testo & "<tr><td class=""label"">email:</td></tr><tr><td class=""content_right"">con email</td></tr>"
			elseif instr(1, Session("ADV_search_email_tipo"), "0", vbTextCompare)>0 then
				'senza email
				sql = sql & " AND IDElencoIndirizzi NOT IN (SELECT id_Indirizzario FROM tb_ValoriNumeri WHERE id_TipoNumero=6) "
				testo = testo & "<tr><td class=""label"">email:</td></tr><tr><td class=""content_right"">senza email</td></tr>"
			end if
		end if
		
		if session("ADV_search_web")<>"" then
			'cerca una stringa tra gli indirizzi web
			sql = sql & " AND IDElencoIndirizzi IN (SELECT id_Indirizzario FROM tb_ValoriNumeri " &_
						" WHERE " + SQL_FullTextSearch(session("ADV_search_web"), "ValoreNumero") + " AND (id_TipoNumero<=7))"
			testo = testo & "<tr><td class=""label"">indirizzo web:</td></tr><tr><td class=""content_right"">" & Session("ADV_search_web") & "</td></tr>"	
		elseif Session("ADV_search_web_tipo")<>"" AND _
			NOT(instr(1, Session("ADV_search_web_tipo"), "0", vbTextCompare)>0 AND _
				instr(1, Session("ADV_search_web_tipo"), "1", vbTextCompare)>0 )then
			if instr(1, Session("ADV_search_web_tipo"), "1", vbTextCompare)>0 then
				'con indirizzo web
				sql = sql & " AND IDElencoIndirizzi IN (SELECT id_Indirizzario FROM tb_ValoriNumeri WHERE id_TipoNumero=7) "
				testo = testo & "<tr><td class=""label"">indirizzo web:</td></tr><tr><td class=""content_right"">con indirizzo web</td></tr>"
			elseif instr(1, Session("ADV_search_web_tipo"), "0", vbTextCompare)>0 then
				'senza indirizzo web
				sql = sql & " AND IDElencoIndirizzi NOT IN (SELECT id_Indirizzario FROM tb_ValoriNumeri WHERE id_TipoNumero=7) "
				testo = testo & "<tr><td class=""label"">indirizzo web:</td></tr><tr><td class=""content_right"">senza indirizzo web</td></tr>"
			end if
		end if
		
		if Session("ADV_search_codFisc")<>"" then
			'ricerca per codice fiscale
			sql = sql & " AND " + SQL_FullTextSearch(session("ADV_search_codFisc"), "CF")
			testo = testo & "<tr><td class=""label"">Partita I.V.A. / Cod. Fiscale:</td></tr><tr><td class=""content_right"">" & Session("ADV_search_codFisc") & "</td></tr>"	
		elseif Session("ADV_search_codFisc_tipo")<>"" AND _
			NOT(instr(1, Session("ADV_search_codFisc_tipo"), "0", vbTextCompare)>0 AND _
				instr(1, Session("ADV_search_codFisc_tipo"), "1", vbTextCompare)>0 )then
			if instr(1, Session("ADV_search_codFisc_tipo"), "1", vbTextCompare)>0 then
				'con codice fiscale
				sql = sql & " AND (NOT " & SQL_IsNULL(conn, "CF") & " AND CF<>'')" 
				testo = testo & "<tr><td class=""label"">Partita I.V.A. / Cod. Fiscale:</td></tr><tr><td class=""content_right"">con P. I.V.A. o Cod. Fiscale</td></tr>"
			elseif instr(1, Session("ADV_search_codFisc_tipo"), "0", vbTextCompare)>0 then
				'senza codice fiscale
				sql = sql & " AND (" & SQL_IsNULL(conn, "CF") & " OR CF='')" 
				testo = testo & "<tr><td class=""label"">Partita I.V.A. / Cod. Fiscale:</td></tr><tr><td class=""content_right"">con P. I.V.A. o Cod. Fiscale</td></tr>"
			end if
		end if
		
		if Session("adv_search_dataiscrizione_to")<>"" AND IsDate(Session("adv_search_dataiscrizione_to")) AND _
		   Session("adv_search_dataiscrizione_from")<>"" AND IsDate(Session("adv_search_dataiscrizione_from")) then
		   	'ricerca per data di iscrizione in un intervallo
			sql = sql & " AND " & SQL_BetweenDate(conn, "DataIscrizione", Session("adv_search_dataiscrizione_from"), Session("adv_search_dataiscrizione_to"))
			testo = testo & "<tr><td class=""label"">Data di iscrizione:</td></tr><tr><td class=""content_right"">dal " & Session("adv_search_dataiscrizione_from") & " al " & Session("adv_search_dataiscrizione_to") & "</td></tr>"	
		elseif Session("adv_search_dataiscrizione_from")<>"" AND IsDate(Session("adv_search_dataiscrizione_from")) then
			'ricerca per data do iscrizione maggiore di una data richiesta
			sql = sql & " AND " & SQL_CompareDateTime(conn, "DataIscrizione", adCompareGreaterThan, Session("adv_search_dataiscrizione_from"))
			testo = testo & "<tr><td class=""label"">Data di iscrizione:</td></tr><tr><td class=""content_right"">a partire dal " & Session("adv_search_dataiscrizione_from") & "</td></tr>"	
		elseif Session("adv_search_dataiscrizione_to")<>"" AND IsDate(Session("adv_search_dataiscrizione_to")) then
			'ricerca per data do iscrizione minore di una data richiesta
			sql = sql & " AND " & SQL_CompareDateTime(conn, "DataIscrizione", adCompareLessThan, Session("adv_search_dataiscrizione_to"))
			testo = testo & "<tr><td class=""label"">Data di iscrizione:</td></tr><tr><td class=""content_right"">fino al " & Session("adv_search_dataiscrizione_to") & "</td></tr>"	
		end if
		
		'ricerca in tutti i campi dell'anagrafica, dell'indirizzo e nelle note
		if Session("ADV_search_full")<>"" then
			sql = sql & " AND (( " + SQL_FullTextSearch_contatto_nominativo(conn, Session("ADV_search_full")) + " OR " + _
                                    SQL_FullTextSearch_contatto_indirizzo(conn, Session("ADV_search_full")) + " OR " + _
                                    SQL_FullTextSearch(session("ADV_search_full"), "TitoloElencoIndirizzi;QualificaElencoIndirizzi;NoteElencoIndirizzi") + ")" + _
								" OR (IDElencoIndirizzi IN (SELECT cntRel FROM tb_indirizzario WHERE " + _
									SQL_FullTextSearch_contatto_nominativo(conn, Session("ADV_search_full")) + " OR " + _
                                    SQL_FullTextSearch_contatto_indirizzo(conn, Session("ADV_search_full")) + " OR " + _
                                    SQL_FullTextSearch(session("ADV_search_full"), "TitoloElencoIndirizzi;QualificaElencoIndirizzi;NoteElencoIndirizzi") + ")))"
			testo = testo & "<tr><td class=""label"">full-text</td></tr><tr><td class=""content_right"">" & Session("ADV_search_full") & "</td></tr>"
		end if
		
		sql = sql & " ORDER BY ModoRegistra"
	else
		sql = ""
	end if
	
	'imposta variabili query di ricerca avanzata
	Session("ADV_search_TXT") = testo
	Session("ADV_search_SQL") = sql
	if sql<>"" then%>
		<SCRIPT LANGUAGE="javascript" type="text/javascript">
			opener.document.location = "Contatti.asp"
			window.close();
		</SCRIPT>
	<%end if
end if


%>
<html>
	<head>
		<title>Opzioni di ricerca avanzata</title>
		<link rel="stylesheet" type="text/css" href="../library/stili.css">
		<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
		<META NAME="copyright" CONTENT="Copyright &copy;2003 - next-aim.com">
		<meta name="robots" content="noindex,nofollow" />
		<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
		<SCRIPT LANGUAGE="javascript" src="../library/utils.js" type="text/javascript"></SCRIPT>
		<SCRIPT LANGUAGE="javascript" type="text/javascript">
			function VediTutti_Click(){
				form1.action = "Contatti.asp";
				form1.target = opener.name;
				document.location = "ContattiRicercaAvanzata.asp"
				window.close();
			}
			
			function Cerca_Click(){
				form1.action = "";
				form1.target = "";
			}
		</SCRIPT>
	</head>
<body topmargin="9" onload="window.focus()">
<form action="" method="post" name="form1">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td>
				<table cellspacing="1" cellpadding="0" class="tabella_madre">
					<caption>
						<table width="100%" border="0" cellspacing="0">
							<tr>
								<td class="caption">Opzioni di ricerca avanzata</td>
								<td align="right" style="padding-right:5px;">
									<input type="submit" name="cerca" value="CERCA" class="button" onclick="Cerca_Click()">
									<input type="submit" class="button" name="tutti" value="VEDI TUTTI" onclick="VediTutti_Click()">
									<input type="button" onclick="window.close();" class="button" name="chiudi" value="CHIUDI">
								</td>
							</tr>
						</table>
					</caption>
					<tr><th colspan="2" <%= Search_Bg("ADV_search_rubriche") %>>RUBRICHE</th></tr>
					<tr>
						<td width="69%" class="content">
							<script language="JavaScript" type="text/javascript">
								function ShowName(obj){
									var value = obj.options(obj.selectedIndex).text;
									if (value.length>42)
										alert(obj.options(obj.selectedIndex).text);
								}
							</script>
							<% sql = " SELECT (' ' " & SQL_concat(conn) & " id_rubrica " & SQL_concat(conn) & " ' ') AS id_rubrica, nome_rubrica FROM tb_rubriche " &_
									 " WHERE id_rubrica IN (" & rubriche_visibili & ") " &_
									 " ORDER BY rubrica_esterna, nome_rubrica"
							CALL dropDown(conn, sql, "id_rubrica", "nome_rubrica", "ADV_search_rubriche", Session("ADV_search_rubriche"), true, _
										  "multiple size=""11"" style=""width:100%;"" onDblClick=""ShowName(this);""", LINGUA_ITALIANO)%>
						</td>
						<td>
							<table width="100%" border="0" cellspacing="1">
								<tr>
									<td class="content_center">
										<input class="checkbox" type="Radio" name="ADV_search_rubriche_tipo" value="" <% if session("ADV_search_rubriche_tipo")="" then %> checked <%end if %>>
									</td>
									<td class="content">
										associati ad almeno una rubrica selezionata
									</td>
								</tr>
								<tr>
									<td class="content_center">
										<input class="checkbox" type="Radio" name="ADV_search_rubriche_tipo" value="AND" <% if session("ADV_search_rubriche_tipo")="AND" then %> checked <%end if %>>
									</td>
									<td class="content">
										associati a tutte le rubriche selezionate
									</td>
								</tr>
								<tr>
									<td colspan="2" class="note">
										Ctrl + Click per selezioni multiple.<br>
										Doppio click per visualizzare il nome.
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("ADV_search_iniziali") %>>INIZIALI</th></tr>
					<tr>
						<td colspan="2">
							<table width="100%" border="0" cellspacing="1">
								<tr>
									<%for i=asc("A") to asc("Z")%>
	    								<TD class="content">
											<INPUT class="checkbox" type="checkbox" name="ADV_search_iniziali" value="'<%=chr(i)%>'" <%if instr(1, Session("ADV_search_iniziali"), chr(i), vbTextCompare)>0 then %> checked <% end if %>>
											<%=chr(i)%>
										</TD>
	    								<%if (i - Asc("A")) mod 9 = 8 then%>
											</tr>
											<tr>
										<%end if
									next %>
									<TD class="content">&nbsp;</TD>
								</tr>
							</table>
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("ADV_search_denominazione;ADV_search_denominazione_tipo;ADV_search_lingua") %>>NOME / DENOMINAZIONE</th></tr>
					<tr>
						<td class="content" colspan="2">
							<input type="text" name="ADV_search_denominazione" value="<%= replace(session("ADV_search_denominazione"), """", "&quot;") %>" style="width:100%;">
						</td>
					</tr>
					<tr>
						<td colspan="2">
							<table width="100%" border="0" cellspacing="1">
								<tr>
									<td class="label" style="width:23%;">salvato come:</td>
									<td class="content" style="width:28%;">
										<input class="checkbox" type="checkbox" name="ADV_search_denominazione_tipo" value="0" 
											<% if instr(1, session("ADV_search_denominazione_tipo"), "0", vbTextCompare)>0 then %> checked <%end if %>>
										persona fisica
									</td>
									<td class="content">
										<input class="checkbox" type="checkbox" name="ADV_search_denominazione_tipo" value="1" 
											<% if instr(1, session("ADV_search_denominazione_tipo"), "1", vbTextCompare)>0 then %> checked <%end if %>>
										ente / società / organizzazione
									</td>
								</tr>
								<tr>
									<td class="label">lingua:</td>
									<td colspan="2" class="content"><% CALL DropLingue(conn, NULL, "ADV_search_lingua", session("ADV_search_lingua"), true, true, "width:100px;") %></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<th <%= Search_Bg("ADV_search_indirizzo;ADV_search_indirizzo_completo") %>>INDIRIZZO</th>
						<th <%= Search_Bg("ADV_search_citta;ADV_search_indirizzo_completo") %>>CITT&Agrave;</th>
					</tr>
					<tr>
						<td class="content">
							<input type="text" name="ADV_search_indirizzo" value="<%= replace(session("ADV_search_indirizzo"), """", "&quot;") %>" style="width:100%;">
						</td>
						<td class="content">
							<input type="text" name="ADV_search_citta" value="<%= replace(session("ADV_search_citta"), """", "&quot;") %>" style="width:100%;">
						</td>
					</tr>
					<tr>
						<td colspan="2">
							<table width="100%" border="0" cellspacing="1">
								<tr>
									<td class="label" style="width:25%;">indirizzo postale:</td>
									<td class="content" style="width:26%;">
										<input class="checkbox" type="checkbox" name="ADV_search_indirizzo_completo" value="0" 
											<% if instr(1, session("ADV_search_indirizzo_completo"), "0", vbTextCompare)>0 then %> checked <%end if %>>
										incompleto
									</td>
									<td class="content">
										<input class="checkbox" type="checkbox" name="ADV_search_indirizzo_completo" value="1" 
											<% if instr(1, session("ADV_search_indirizzo_completo"), "1", vbTextCompare)>0 then %> checked <%end if %>>
										completo <span class="note">(indirizzo, cap, citt&agrave;)</span>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("ADV_search_telefono;ADV_search_telefono_tipo") %>>NUMERO TELEFONO</th></tr>
					<tr>
						<td colspan="2">
							<table width="100%" cellspacing="1">
								<tr>
									<td style="width:54%;" class="content" colspan="2">
										<input type="text" name="ADV_search_telefono" value="<%= replace(session("ADV_search_telefono"), """", "&quot;") %>" style="width:100%;">
									</td>
									<td class="content">
										<input class="checkbox" type="checkbox" name="ADV_search_telefono_tipo" value="1" 
											<% if instr(1, session("ADV_search_telefono_tipo"), "1", vbTextCompare)>0 then %> checked <%end if %>>
										con tel.
									</td>
									<td class="content">
										<input class="checkbox" type="checkbox" name="ADV_search_telefono_tipo" value="0" 
											<% if instr(1, session("ADV_search_telefono_tipo"), "0", vbTextCompare)>0 then %> checked <%end if %>>
										senza tel.
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("ADV_search_fax;ADV_search_fax_tipo") %>>NUMERO FAX</th></tr>
					<tr>
						<td colspan="2">
							<table width="100%" cellspacing="1">
								<tr>
									<td style="width:54%;" class="content" colspan="2">
										<input type="text" name="ADV_search_fax" value="<%= replace(session("ADV_search_fax"), """", "&quot;") %>" style="width:100%;">
									</td>
									<td class="content">
										<input class="checkbox" type="checkbox" name="ADV_search_fax_tipo" value="1" 
											<% if instr(1, session("ADV_search_fax_tipo"), "1", vbTextCompare)>0 then %> checked <%end if %>>
										con fax
									</td>
									<td class="content">
										<input class="checkbox" type="checkbox" name="ADV_search_fax_tipo" value="0" 
											<% if instr(1, session("ADV_search_fax_tipo"), "0", vbTextCompare)>0 then %> checked <%end if %>>
										senza fax
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><th colspan="2"  <%= Search_Bg("ADV_search_email;ADV_search_email_tipo") %>>INDIRIZZO EMAIL</th></tr>
					<tr>
						<td colspan="2">
							<table width="100%" cellspacing="1">
								<tr>
									<td style="width:54%;" class="content" colspan="2">
										<input type="text" name="ADV_search_email" value="<%= replace(session("ADV_search_email"), """", "&quot;") %>" style="width:100%;">
									</td>
									<td class="content">
										<input class="checkbox" type="checkbox" name="ADV_search_email_tipo" value="1" 
											<% if instr(1, session("ADV_search_email_tipo"), "1", vbTextCompare)>0 then %> checked <%end if %>>
										con email
									</td>
									<td class="content">
										<input class="checkbox" type="checkbox" name="ADV_search_email_tipo" value="0" 
											<% if instr(1, session("ADV_search_email_tipo"), "0", vbTextCompare)>0 then %> checked <%end if %>>
										senza email
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("ADV_search_web;ADV_search_web_tipo") %>>INDIRIZZO WEB</th></tr>
					<tr>
						<td colspan="2">
							<table width="100%" cellspacing="1">
								<tr>
									<td style="width:54%;" class="content" colspan="2">
										<input type="text" name="ADV_search_web" value="<%= replace(session("ADV_search_web"), """", "&quot;") %>" style="width:100%;">
									</td>
									<td class="content">
										<input class="checkbox" type="checkbox" name="ADV_search_web_tipo" value="1" 
											<% if instr(1, session("ADV_search_web_tipo"), "1", vbTextCompare)>0 then %> checked <%end if %>>
										con sito
									</td>
									<td class="content">
										<input class="checkbox" type="checkbox" name="ADV_search_web_tipo" value="0" 
											<% if instr(1, session("ADV_search_web_tipo"), "0", vbTextCompare)>0 then %> checked <%end if %>>
										senza sito
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("ADV_search_codfisc;ADV_search_codfisc_tipo") %>>CODICE FISCALE / PARTITA I.V.A.</th></tr>
					<tr>
						<td colspan="2">
							<table width="100%" cellspacing="1">
								<tr>
									<td style="width:54%;" class="content" colspan="2">
										<input type="text" name="ADV_search_codfisc" value="<%= replace(session("ADV_search_codfisc"), """", "&quot;") %>" style="width:100%;">
									</td>
									<td class="content" nowrap>
										<input class="checkbox" type="checkbox" name="ADV_search_codfisc_tipo" value="1" 
											<% if instr(1, session("ADV_search_codfisc_tipo"), "1", vbTextCompare)>0 then %> checked <%end if %>>
										con C.F.
									</td>
									<td class="content" nowrap>
										<input class="checkbox" type="checkbox" name="ADV_search_codfisc_tipo" value="0" 
											<% if instr(1, session("ADV_search_codfisc_tipo"), "0", vbTextCompare)>0 then %> checked <%end if %>>
										senza C.F.
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("adv_search_dataiscrizione_from;adv_search_dataiscrizione_to") %>>DATA DI ISCRIZIONE</th></tr>
					<tr>
						<td colspan="2">
							<table width="100%" cellspacing="1">
								<tr>
									<td class="label" style="width:25%;">a partire dal:</td>
									<td class="content">
										<% CALL WriteDataPicker_Input("form1", "adv_search_dataiscrizione_from", Session("adv_search_dataiscrizione_from"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
								<tr>
									<td class="label">fino al:</td>
									<td class="content">
										<% CALL WriteDataPicker_Input("form1", "adv_search_dataiscrizione_to", Session("adv_search_dataiscrizione_to"), "", "/", true, true, LINGUA_ITALIANO) %>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><th colspan="2" <%= Search_Bg("ADV_search_full") %>>FULL-TEXT</th></tr>
					<tr>
						<td class="content" colspan="2">
							<input type="text" name="ADV_search_full" value="<%= replace(session("ADV_search_full"), """", "&quot;") %>" style="width:100%;">
							<div class="note">Cerca in tutti i campi dell'anagrafica, dell'indirizzo e nelle note.</div>
						</td>
					</tr>
					<tr>
						<td colspan="2" class="footer">
							<input type="submit" name="cerca" value="CERCA" class="button" onclick="Cerca_Click()">
							<input type="submit" class="button" name="tutti" value="VEDI TUTTI" onclick="VediTutti_Click()">
							<input type="button" onclick="window.close();" class="button" name="chiudi" value="CHIUDI">
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</form>
</body>
</html>
<% 
conn.close 
set rs = nothing
set conn = nothing
%>