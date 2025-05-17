<!--#INCLUDE FILE="../../tools.asp" -->
<%
dim conn
set conn = server.createobject("adodb.connection")
conn.open Application("L_conn_ConnectionString")

response.write "Inizio Aggiornamento<br>"
conn.BeginTrans
ObjToPlugins(conn)
conn.CommitTrans
response.write "AGGIORNAMENTO ESEGUITO"

conn.close
set conn = nothing

'*****************************************************************************************************************
'FUNZIONI PER CONVERSIONE OGGETTO COM IN PLUGINS (eventi, luoghi, strutture, NEXTevents, ...)
'*****************************************************************************************************************

'conn:		connessione gia aperta su dbLayers
Sub ObjToPlugins(conn)
	dim sql, CodiceAptPart, CodiceTipoLuoghi, CodiceTipoSpiagge, CodiceTipoLocali, CodiceTipoNotizie, CodiceTipoStrutture
    'costruisce i codici di tipo per ogni APT
    CodiceTipoStrutture = "ASS00R00000"
    CodiceAptPart = replace(Application("APT"), "000000", "")
    CodiceTipoLuoghi = CodiceAptPart + "L00000"
    CodiceTipoSpiagge = CodiceAptPart + "S00000"
    CodiceTipoLocali = CodiceAptPart + "T00000"
    CodiceTipoNotizie = CodiceAptPart + "N00000"
    
	'CARNET
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/CarnetBtn.asp'"+ _
		  " WHERE identif_objects = 'nextApt.Carnet' AND img_objects LIKE '%Carnet_ADD%'"
	conn.Execute(sql)		'pulsante aggiungi
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/Carnet.asp'"+ _
		  " WHERE identif_objects = 'nextApt.Carnet' AND img_objects LIKE '%Carnet.gif'"
	conn.Execute(sql)		'scheda carnet
	
	'EVENTI
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/EventiContemporanea.asp'"+ _
		  " WHERE identif_objects = 'nextApt.Eventi' AND img_objects LIKE '%Eventi_contemp_report.gif'"
	conn.Execute(sql)		'form scelta data contemporanea
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/EventiContemporaneaBtn.asp'"+ _
		  " WHERE identif_objects = 'nextApt.Eventi' AND img_objects LIKE '%Eventi_Contemporanea%'"
	conn.Execute(sql)		'pulsante vedi eventi contemporanea
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/EventiReport.asp'"+ _
		  " WHERE identif_objects = 'nextApt.Eventi' AND img_objects LIKE '%Eventi_report.gif'"
	conn.Execute(sql)		'report
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/EventiRicerca.asp'"+ _
		  " WHERE identif_objects = 'nextApt.Eventi' AND img_objects LIKE '%Eventi_ricerca.gif'"
	conn.Execute(sql)		'ricerca
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/EventiRicercaMin.asp'"+ _
		  " WHERE identif_objects = 'nextApt.Eventi' AND img_objects LIKE '%Eventi_ricerca_ridotta.gif'"
	conn.Execute(sql)		'ricerca ridotta
	
	dim rs, PaginaLuoghi
	sql = "SELECT * FROM tb_webs"
	set rs = conn.execute(sql)
	while not rs.eof
		sql = " SELECT TOP 1 id_pagineSito FROM tb_pagineSito " + _
			  " INNER JOIN tb_layers ON tb_pagineSito.id_pagDyn_IT = tb_layers.id_pag " + _
			  " WHERE id_tipo=4 AND nome LIKE '%Luoghi_scheda.gif' AND id_web=" & rs("id_webs")
		PaginaLuoghi = cIntero(GetValueList(Conn, NULL, sql))
		sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/EventiScheda.asp'"+ _
			  ", param_list = param_list & ';"& vbCrLf &"LuoghiLinkPage=" & PaginaLuoghi & ";' " + _
			  " WHERE identif_objects = 'nextApt.Eventi' AND img_objects LIKE '%Eventi_scheda.gif'"
		conn.Execute(sql)		'scheda
		rs.movenext
	wend
	
	
	
	'LUOGHI
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheReport.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoLuoghi + ";' " + _
		  " WHERE identif_objects = 'nextApt.Luoghi' AND img_objects LIKE '%Luoghi_report.gif'"
	conn.Execute(sql)		'report
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheRicerca.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoLuoghi + ";' " + _
		  " WHERE identif_objects = 'nextApt.Luoghi' AND img_objects LIKE '%Luoghi_ricerca.gif'"
	conn.Execute(sql)		'ricerca
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheRicercaMin.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoLuoghi + ";' " + _
		  " WHERE identif_objects = 'nextApt.Luoghi' AND img_objects LIKE '%Luoghi_ricerca_ridotta.gif'"
	conn.Execute(sql)		'ricerca ridotta
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheScheda.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoLuoghi + ";' " + _
		  " WHERE identif_objects = 'nextApt.Luoghi' AND img_objects LIKE '%Luoghi_scheda.gif'"
	conn.Execute(sql)		'scheda
	
	'NOTIZIE
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheReport.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoNotizie + ";' " + _
		  " WHERE identif_objects = 'nextApt.NotizieUtili' AND img_objects LIKE '%Notizie_report.gif'"
	conn.Execute(sql)		'report
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheRicerca.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoNotizie + ";' " + _
		  " WHERE identif_objects = 'nextApt.NotizieUtili' AND img_objects LIKE '%Notizie_Ricerca.gif'"
	conn.Execute(sql)		'ricerca
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheRicercaMin.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoNotizie + ";' " + _
		  " WHERE identif_objects = 'nextApt.NotizieUtili' AND img_objects LIKE '%Notizie_ricerca_ridotta.gif'"
	conn.Execute(sql)		'ricerca ridotta
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheScheda.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoNotizie + ";' " + _
		  " WHERE identif_objects = 'nextApt.NotizieUtili' AND img_objects LIKE '%Notizie_scheda.gif'"
	conn.Execute(sql)		'scheda
	
	'LOCALI E SERVIZI
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheReport.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoLocali + ";' " + _
		  " WHERE identif_objects = 'nextApt.Locali' AND img_objects LIKE '%LS_rep_tr.gif'"
	conn.Execute(sql)		'report
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheRicerca.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoLocali + ";' " + _
		  " WHERE identif_objects = 'nextApt.Locali' AND img_objects LIKE '%ls_form_tr.gif'"
	conn.Execute(sql)		'ricerca
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheScheda.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoLocali + ";' " + _
		  " WHERE identif_objects = 'nextApt.Locali' AND img_objects LIKE '%LS_sche_tr.gif'"
	conn.Execute(sql)		'scheda
	
	'SPIAGGE
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheReport.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoSpiagge + ";' " + _
		  " WHERE identif_objects = 'nextApt.Spiagge' AND img_objects LIKE '%Spiagge_report.gif'"
	conn.Execute(sql)		'report
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheScheda.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoSpiagge + ";' " + _
		  " WHERE identif_objects = 'nextApt.Spiagge' AND img_objects LIKE '%Spiagge_scheda.gif'"
	conn.Execute(sql)		'scheda
	
	'STRUTTURE
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheReport.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoStrutture + ";' " + _
		  " WHERE identif_objects = 'nextApt.StruttureRicettive' AND img_objects LIKE '%Ricettivita_report.gif'"
	conn.Execute(sql)		'report
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheRicerca.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoStrutture + ";' " + _
		  " WHERE identif_objects = 'nextApt.StruttureRicettive' AND img_objects LIKE '%Ricettivita_ricerca%'"
	conn.Execute(sql)		'ricerca
	sql = " UPDATE tb_objects SET " + _
		  		 " identif_objects = 'amministrazione/PluginsAPT/AnagraficheScheda.asp', "+ _
				 " param_list = param_list & ';"& vbCrLf &"tipo=" + CodiceTipoStrutture + ";' " + _
		  " WHERE identif_objects = 'nextApt.StruttureRicettive' AND (img_objects LIKE '%Ricettivita_scheda.gif' OR img_objects LIKE 'Ricettivita_scheda_venicesystem.gif')"
	conn.Execute(sql)		'scheda
	
	'STAMPA
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/obj_stampa.asp'"+ _
		  " WHERE identif_objects = 'nextApt.Stampa'"
	conn.Execute(sql)		'pulsante
	
	'BANNER
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/banner.asp'"+ _
		  " WHERE identif_objects = 'nextApt.Banner'"
	conn.Execute(sql)
	
	'CALENDARIO
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/calendario.asp'"+ _
		  " WHERE identif_objects = 'nextApt.calendar'"
	conn.Execute(sql)
	
	'FRAME
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/frame.asp'"+ _
		  " WHERE identif_objects = 'nextApt.Frame'"
	conn.Execute(sql)
	
	'MAPPA
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/obj_mappa.asp'"+ _
		  " WHERE identif_objects = 'nextApt.map'"
	conn.Execute(sql)
	
	'IN PROSSIMITA
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/pluginsAPT/obj_prossimita_report.asp'"+ _
		  " WHERE identif_objects = 'nextApt.Luoghi' AND img_objects LIKE '%Luoghi_Prossimita_report.gif'"
	conn.Execute(sql)		'report
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/pluginsAPT/obj_prossimita.asp'"+ _
		  " WHERE identif_objects = 'nextApt.Luoghi' AND img_objects LIKE '%Luoghi_Prossimita_IT.gif'"
	conn.Execute(sql)		'pulsante
	
	'MENU
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/menu.asp'"+ _
		  " WHERE identif_objects = 'nextApt.Menu'"
	conn.Execute(sql)
	
	'TOPMENU
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/TopMenu.asp'"+ _
		  " WHERE identif_objects = 'nextApt.TopMenu'"
	conn.Execute(sql)
	
	'EVENTO RANDOM
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/EventiReportRandom.asp'"+ _
		  " WHERE identif_objects LIKE '%obj_evento_random.asp%'"
	conn.Execute(sql)
	
	'EVENTO SCROLLER
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/PluginsAPT/EventiReportScroller.asp'"+ _
		  " WHERE identif_objects LIKE '%obj_evento_scroller.asp%'"
	conn.Execute(sql)
	
	'CREDITS
	sql = " UPDATE tb_objects SET identif_objects = 'amministrazione/library/plugins/obj_credits.asp'"+ _
		  " WHERE identif_objects LIKE '%credits%'"
	conn.Execute(sql)
	
	'AGGIORNAMENTO LAYERS OGGETTO
	sql = " UPDATE tb_layers INNER JOIN tb_objects ON tb_layers.id_objects = tb_objects.id_objects"+ _
		  " SET aspcode = identif_objects, testo = param_list"
	conn.Execute(sql)
	
	'AGGIORNA VALORI SUI MENU: rimuove indirizzi assoluti per le homepage ed i cambi lingua
	sql = "UPDATE tb_menuitem SET link_menuItem_IT='default.asp', link_menuItem_EN='', link_menuItem_DE='', link_menuItem_FR='', link_menuItem_ES='' WHERE titolo_menuItem_it LIKE '%home%' AND id_pagina=0"
	conn.Execute(sql)
	
	'aggiorna percorsi per lingua inglese e ritorno ad italiana
	sql = "UPDATE tb_menuitem SET link_menuItem_IT='eng/default.asp?lingua=en', " + _
								" link_menuItem_EN='../default.asp?lingua=it', " + _
								" link_menuItem_DE='../eng/default.asp?lingua=en' " + _
		  " WHERE titolo_menuItem_IT LIKE '%english%' AND id_pagina=0 "
	conn.Execute(sql)
	
	'aggiorna percorsi per lingua tedesca e ritorno ad italiana
	sql = "UPDATE tb_menuitem SET link_menuItem_IT='deu/default.asp?lingua=de', " + _
								" link_menuItem_EN='../deu/default.asp?lingua=de', " + _
								" link_menuItem_DE='../default.asp?lingua=it' " + _
		  " WHERE titolo_menuItem_IT LIKE '%deutsch%' AND id_pagina=0 "
	conn.Execute(sql)
	
	'aggiorna percorso amministrazione
	sql = "UPDATE tb_menuitem SET link_menuItem_IT='amministrazione/', " + _
								" link_menuItem_EN='../amministrazione/', " + _
								" link_menuItem_DE='../amministrazione/' " + _
		  " WHERE link_menuItem_IT LIKE '%amministrazione%' AND id_pagina=0 "
	conn.Execute(sql)
	
End Sub
%>