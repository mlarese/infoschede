<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="Tools_Contatti.asp" -->
<!--#INCLUDE FILE="Tools_DocumentiFiles.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<% 
CALL CheckAutentication(Session("LOGIN_4_LOG") <> "")


dim conn, rs, sql, i, var, testo, prefix, cambioTipo, id

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")

Prefix = Session("DOC_PREFIX")

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	cambioTipo = (request.form("adv_doc_tipologia") <> Session(Prefix & "adv_doc_tipologia"))
	
	'azzera tutte le variabili di ricerca
	for each var in Session.Contents
		if left(var, len(Prefix &"adv_doc_")) = Prefix &"adv_doc_" then
			Session(var) = ""
		end if
	next
	
	'imposta variabili di ricerca (ricerca avanzata)
	i = 0
	for each var in request.form
		if instr(1, var, "adv_doc_", vbTextCompare)>0 AND cString(request(var))<>"" then
			Session(Prefix & var) = request.form(var)
			i = i + 1
		end if
	next
	
	if i>0 AND request.form("cerca") <> "" then	'controlla se è stata impostata una ricerca
		'imposta criteri per ricerca avanzata
		testo = ""
		sql = " SELECT * FROM (tb_documenti INNER JOIN tb_tipologie ON tb_documenti.doc_tipologia_id=tb_tipologie.tipo_id) "& _
			  " INNER JOIN tb_admin ON tb_documenti.doc_creatore_id=tb_admin.id_admin "& _
			  " WHERE (doc_creatore_id="& Session("ID_ADMIN") &" OR "& AL_query(conn, AL_DOCUMENTI) &")"
		
		if Session("DOC_PRA_ID")<>"" then
			'filtra su id pratica (mostra solo documenti della pratica)
			sql = sql & " AND doc_pratica_id=" & Session("DOC_PRA_ID")
		
			'recupera nome pratica e nome contatto
			if Session("DOC_PRA_NOME") = "" OR Session("DOC_PRA_CNT_NOME") = "" then
				sql = "SELECT pra_nome, isSocieta, NomeElencoIndirizzi, CognomeElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi " & _
					   " FROM tb_pratiche INNER JOIN tb_indirizzario ON tb_pratiche.pra_cliente_id = tb_Indirizzario.IDElencoIndirizzi " & _
					   " WHERE pra_id=" & Session("DOC_PRA_ID")
				set rs = conn.execute(sql)
				Session("DOC_PRA_NOME") = rs("pra_nome")
				Session("DOC_PRA_CNT_NOME") = ContactName(rs)
			end if
			
		elseif Session("DOC_ATT_ID")<>"" then
			'filtra su id documento: mostra solo documenti allegati alla pratica
			sql = sql & " AND doc_id IN (SELECT all_documento_id FROM tb_allegati WHERE all_attivita_id=" & Session("DOC_ATT_ID") & ") "
			
			'recupera oggetto attivita
			if Session("DOC_ATT_OGGETTO")="" then
				sql = "SELECT att_oggetto FROM tb_attivita WHERE att_id=" & Session("DOC_ATT_ID")
				set rs = conn.execute(sql)
				Session("DOC_ATT_OGGETTO") = rs("att_oggetto")
			end if
		end if
		
		'filtra per nome documento
		if Session(Prefix & "adv_doc_nome")<>"" then
			sql = sql & " AND " + SQL_FullTextSearch(Session(Prefix & "adv_doc_nome"), "doc_nome")
			testo = testo & "<tr><td class=""label"">nel cui nome compaia:</td></tr><tr><td class=""content_right"">"& Session(Prefix & "adv_doc_nome") &"</td></tr>"
		end if
		
		'filtra per full-text su nome e note
		if Session(Prefix & "adv_doc_full_text") <> "" then
			sql = sql & " AND " + SQL_FullTextSearch(Session(Prefix & "adv_doc_full_text"), "doc_nome;doc_note")
			testo = testo & "<tr><td class=""label"">full-text:</td></tr><tr><td class=""content_right"">"& Session(Prefix & "adv_doc_full_text") &"</td></tr>"
		end if
		
		'filtra per nome file contenuto
		if Session(Prefix & "adv_doc_file")<>"" then
			sql = sql & " AND doc_id IN (SELECT rel_documento_id FROM rel_documenti_files " & _
						" INNER JOIN tb_files ON rel_documenti_files.rel_files_id = tb_files.f_id " & _
						" WHERE " + SQL_FullTextSearch(Session(Prefix & "adv_doc_file"), "f_original_name") + ")"
			testo = testo & "<tr><td class=""label"">i nome dei file associati:</td></tr><tr><td class=""content_right"">"& Session(Prefix & "adv_doc_file") &"</td></tr>"
		end if
		
		'filtra per pratica e per contatto 
		if Session(Prefix & "adv_doc_pratica")<>"" OR Session(Prefix & "adv_doc_contatto")<>"" then
			sql = sql & " AND doc_pratica_id IN (SELECT pra_id FROM tb_pratiche "
			if Session(Prefix & "adv_doc_contatto")<>"" then
				sql = sql & " INNER JOIN tb_indirizzario ON tb_pratiche.pra_cliente_id=tb_indirizzario.idElencoIndirizzi " & _
							" WHERE " + SQL_FullTextSearch_Contatto_Nominativo(conn, Session(Prefix & "adv_doc_contatto"))
				if Session(Prefix & "adv_doc_pratica")<>"" then
					sql = sql & " AND "
				end if
				testo = testo & "<tr><td class=""label"">appartente ad una pratica di:</td></tr><tr><td class=""content_right"">"& Session(Prefix & "adv_doc_contatto") &"</td></tr>"
			else
				sql = sql & " WHERE "
			end if
			if Session(Prefix & "adv_doc_pratica")<>"" then
				sql = sql & SQL_FullTextSearch(Session(Prefix & "adv_doc_pratica"), "pra_nome")
				testo = testo & "<tr><td class=""label"">appartenente alla pratica:</td></tr><tr><td class=""content_right"">"& Session(Prefix & "adv_doc_pratica") &"</td></tr>"
			end if
			sql = sql & " )"
		end if
		
		'filtra per data di creazione
		if isDate(Session(Prefix & "adv_doc_data_creazione_from")) then
			sql = sql & " AND " & SQL_CompareDateTime(conn, "doc_dataC", adCompareGreaterThan, Session(Prefix & "adv_doc_data_creazione_from")) & " "
			testo = testo & "<tr><td class=""label"">creato dal:</td></tr><tr><td class=""content_right"">"& DateTimeIta(Session(Prefix & "adv_doc_data_creazione_from")) &"</td></tr>"
		end if
		if isDate(Session(Prefix & "adv_doc_data_creazione_to")) then
			sql = sql & " AND " & SQL_CompareDateTime(conn, "doc_dataC", adCompareLessThan, Session(Prefix & "adv_doc_data_creazione_to")) & " "
			testo = testo & "<tr><td class=""label"">creato prima del:</td></tr><tr><td class=""content_right"">"& DateTimeIta(Session(Prefix & "adv_doc_data_creazione_to")) &"</td></tr>"
		end if
		
		'filtra per tipologia
		if Session(Prefix & "adv_doc_tipologia")<>"" then
			sql = sql & " AND doc_tipologia_id = " & ParseSQL(Session(Prefix & "adv_doc_tipologia"), adChar)
			testo = testo & "<tr><td class=""label"">di tipo:</td></tr><tr><td class=""content_right"">"& GetValueList(conn, NULL, "SELECT tipo_nome FROM tb_tipologie WHERE tipo_id="& Session(Prefix & "adv_doc_tipologia")) &"</td></tr>"
		end if
		
		DesRicercaQuery sql, testo, "tb_descrittori", "descr_id", "descr_nome", "", "doc_id", "rel_documenti_descrittori", "rdd_descrittore_id", "rdd_documento_id", "rdd_valore", prefix & "adv_doc"
		
		sql = sql & " ORDER BY doc_nome"
	else
		sql = ""
	end if
	
	'imposta variabili query di ricerca avanzata
	Session(Prefix & "ADV_doc_TXT") = testo
	Session(Prefix & "ADV_doc_SQL") = sql
	
	if sql<>"" AND NOT cambioTipo then%>
		<SCRIPT LANGUAGE="javascript" type="text/javascript">
			opener.document.location = "Documenti.asp"
			window.close();
		</SCRIPT>
<%	end if
end if

if Session(Prefix & "adv_doc_tipologia") <> "" then %>
		<SCRIPT LANGUAGE="javascript" type="text/javascript">
			window.resizeTo(422, 550)
		</SCRIPT>
<% end if %>
<html>
	<head>
		<title>Opzioni di ricerca avanzata</title>
		<link rel="stylesheet" type="text/css" href="../library/stili.css">
		<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
		<meta name="robots" content="noindex,nofollow" />
		<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
		<SCRIPT LANGUAGE="javascript" src="../library/utils.js" type="text/javascript"></SCRIPT>
		<SCRIPT LANGUAGE="javascript" type="text/javascript">
			function VediTutti_Click(){
				form1.action = "Documenti.asp";
				form1.target = opener.name;
				document.location = "DocumentiRicercaAvanzata.asp"
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
					<tr><th <%= Search_Bg(prefix & "adv_doc_nome") %>>NOME DOCUMENTO</th></tr>
					<tr>
						<td class="content">
							<input type="text" name="adv_doc_nome" value="<%= replace(session(prefix & "adv_doc_nome"), """", "&quot;") %>" style="width:100%;">
						</td>
					</tr>
					<tr><th <%= Search_Bg(prefix & "adv_doc_file") %>>NOME FILE</th></tr>
					<tr>
						<td class="content">
							<input type="text" name="adv_doc_file" value="<%= replace(session(prefix & "adv_doc_file"), """", "&quot;") %>" style="width:100%;">
						</td>
					</tr>
					<% If Session("DOC_PRA_ID") = "" AND Session("DOC_ATT_ID")="" then %>
						<tr><th <%= Search_Bg(Prefix & "adv_doc_pratica") %>>PRATICA</th></tr>
						<tr>
							<td class="content">
								<input type="text" name="adv_doc_pratica" value="<%= replace(Session(Prefix & "adv_doc_pratica"), """", "&quot;") %>" style="width:100%;">
							</td>
						</tr>
						<tr><th <%= Search_Bg(Prefix & "adv_doc_contatto") %>>CONTATTO</th></tr>
						<tr>
							<td class="content">
								<input type="text" name="adv_doc_contatto" value="<%= replace(Session(Prefix & "adv_doc_contatto"), """", "&quot;") %>" style="width:100%;">
							</td>
						</tr>
					<% End If %>
					<tr><th <%= Search_Bg(Prefix & "adv_doc_data_creazione_from;" & Prefix & "adv_doc_data_creazione_to") %>>DATA CREAZIONE</td></tr>
					<tr><td class="label">a partire dal:</td></tr>
					<tr>
						<td class="content">
							<% CALL WriteDataPicker_Input("form1", "adv_doc_data_creazione_from", Session(Prefix & "adv_doc_data_creazione_from"), "", "/", true, true, LINGUA_ITALIANO) %>
						</td>
					</tr>
					<tr><td class="label">fino al:</td></tr>
					<tr>
						<td class="content">
							<% CALL WriteDataPicker_Input("form1", "adv_doc_data_creazione_to", Session(Prefix & "adv_doc_data_creazione_to"), "", "/", true, true, LINGUA_ITALIANO) %>
						</td>
					</tr>
					<tr><th <%= Search_Bg(Prefix & "adv_doc_full_text") %>>FULL-TEXT</th></tr>
					<tr>
						<td class="content">
							<input type="text" name="adv_doc_full_text" value="<%= replace(session(Prefix & "adv_doc_full_text"), """", "&quot;") %>" style="width:100%;">
						</td>
					</tr>
					<tr><th <%= Search_Bg(Prefix & "adv_doc_tipologia") %>>TIPOLOGIA</th></tr>
					<tr>
						<td class="content">
							<% sql = "SELECT * FROM tb_tipologie ORDER BY tipo_nome"
							CALL dropDown(conn, sql, "tipo_id", "tipo_nome", "adv_doc_tipologia", Session(Prefix & "adv_doc_tipologia"), false, "onchange=""form1.submit()"" style=""width:100%;""", LINGUA_ITALIANO) %>
						</td>
					</tr>
					<% 	If Session(Prefix & "adv_doc_tipologia") <> "" then %>
					<tr><th class="L2">DESCRITTORI DELLA TIPOLOGIA</th></tr>
					<%		sql = "SELECT * FROM tb_descrittori d INNER JOIN rel_tipologie_descrittori r " & _
								  "ON d.descr_id = r.rtd_descrittore_id WHERE rtd_tipologia_id="& Session(Prefix & "adv_doc_tipologia")
							DesRicerca conn, sql, "tb_descrittori", "descr_id", "descr_nome", "descr_tipo", "", "adv_doc"
					 	End If 
					%>
					<tr>
						<td class="footer">
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

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>