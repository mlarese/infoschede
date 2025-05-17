<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="../TOOLS.ASP"-->
<!--#INCLUDE FILE="../TOOLS4PlugIn.ASP"-->
<!--#INCLUDE FILE="../CLASSCONFIGURATION.ASP"-->
<!--#INCLUDE FILE="../ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="../ClassPageNavigator.asp" -->
<% 
dim config
set config = new Configuration
'impostazione delle proprieta' di default
Config.AddDefault "Categorie", ""
Config.AddDefault "page_scheda", ""
Config.AddDefault "gallery_per_pagina", "4"
Config.AddDefault "lunghezza_testo", ""

Config.AddDefault "NoRecords_IT", "Nessun record disponibile."
Config.AddDefault "NoRecords_EN", "No record available."
Config.AddDefault "NoRecords_FR", "Aucuns record disponibles."
Config.AddDefault "NoRecords_DE", "Keine Record vorhanden."
Config.AddDefault "NoRecords_ES", "Ningunos record disponibles."

dim pager
set Pager = new PageNavigator
Pager.ParentConfiguration = Config

dim conn, rs, rsf, sql
Set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.recordset")
set rsf = Server.CreateObject("ADODB.recordset")

sql = " SELECT * FROM ptb_gallery LEFT JOIN ptb_CategorieGallery " + _
	  " ON (ptb_gallery.gallery_idcategoria = ptb_CategorieGallery.catC_id AND catC_albero_visibile AND catC_visibile) " + _
	  " WHERE gallery_visibile"
if Config("Categorie")<>"" then
	sql = sql & " AND catC_id IN (" & cIntero(Config("Categorie")) & ") "
end if
sql = sql + " ORDER BY catC_ordine_assoluto, gallery_ordine "

CALL Pager.OpenSmartRecordset(conn, rs, sql, cInt(config("gallery_per_pagina"))) 

if not rs.eof then
	rs.AbsolutePage = Pager.PageNo
	while not rs.eof and rs.AbsolutePage = Pager.PageNo %>
		<div class="gallery_elenco">
			<h1 class="gallery_elenco"><%= CBL(rs, "gallery_name") %></h1>
			
			<% 	'sottotitolo
				sql = " SELECT * FROM prel_descrittori_gallery INNER JOIN ptb_descrittori " + _
						 " ON prel_descrittori_gallery.rdi_descrittore_id = ptb_descrittori.des_id " + _
						 " WHERE rdi_gallery_id=" & cIntero(rs("gallery_id")) &" AND des_tipo="
				rsf.open sql & adVarChar, conn, adOpenStatic, adLockOptimistic, adCmdText
				if not rsf.eof then %>
			<h2 class="gallery_elenco">
				<%= DesFormatValue(rsf("des_tipo"), CBL(rsf, "rdi_valore"), CBL(rsf, "rdi_memo"), "", 0, "") %>
			</h2>
			<% 	end if
				rsf.close %>
			
			<% 	'immagine
				sql = "SELECT * FROM ptb_Immagini WHERE I_Gallery_id=" & cIntero(rs("gallery_id")) & " ORDER BY I_ordine"
				rsf.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText 
				if not rsf.eof then
					CALL WriteThumbnail(Config.EncodePage(Config("page_scheda")), _
								    	Config.ImageURL & rsf("I_thumb"), _
										rsf("I_id"), rs("gallery_id"), _
										CBL(rsf, "I_didascalia"), true)
					rsf.movenext
				end if 
				rsf.close %>
					
			<% 	'descrizione
				sql = " SELECT * FROM prel_descrittori_gallery INNER JOIN ptb_descrittori " + _
						 " ON prel_descrittori_gallery.rdi_descrittore_id = ptb_descrittori.des_id " + _
						 " WHERE rdi_gallery_id=" & cIntero(rs("gallery_id")) &" AND des_tipo="
				rsf.open sql & adLongVarChar, conn, adOpenStatic, adLockOptimistic, adCmdText 
				if not rsf.eof then %>
			<p class="gallery_elenco">
			<%		if CIntero(config("lunghezza_testo")) > 0 then
						response.write Sintesi(DesFormatValue(rsf("des_tipo"), CBL(rsf, "rdi_valore"), CBL(rsf, "rdi_memo"), "", 0, ""), config("lunghezza_testo"), "...")
					else
						response.write DesFormatValue(rsf("des_tipo"), CBL(rsf, "rdi_valore"), CBL(rsf, "rdi_memo"), "", 0, "")
					end if %>
			</p>
			<% 	end if
				rsf.close %>
			
			<%	'link
				rsf.open sql & adGUID, conn, adOpenStatic, adLockOptimistic, adCmdText
				if not rsf.eof then %>
			<p class="gallery_elenco_link">qqq
				<a href="dynalay.asp?pagina=<%= Config.EncodePage(rsf("rdi_valore_it")) %>" title="<%= CBL(rs, "gallery_name") %>"><%= CBL(Config, "label_link") %></a>
			</p>
			<% 	end if
				rsf.close %>
		</div>
		<% rs.movenext
	wend
	
	'paginazione
	if Pager.PageCount>1 then%>
		<div class="pager">
			<%CALL Pager.Render_PageNavigator(Pager.PageCount, "pager", "pager", "pager_current")%>	
		</div>
<% 	end if
else %>
	<div class="NoRecords"><%= CBL(Config, "NoRecords") %></div>
<%
end if
set rs = nothing
set rsf = nothing
conn.close
set conn = nothing


'scrive il codice per fare l'immagine ridotta
sub WriteThumbnail(PageScheda, ImageSrc, ImageId, gallery_id, didascalia, NewPage)%>
	<div class="thumbnail">
		<img src="<%= ImageSrc %>" alt="<%= didascalia %>" border="0">
		<div class="didascalia"><%= didascalia %></div>
	</div>
<%
end sub
%>