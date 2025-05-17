<%@ Language=VBScript CODEPAGE=65001%>
<% response.charset = "UTF-8" %>
<% Imposta_Proprieta_Sito("ID") %>
<!--#INCLUDE FILE="intestazione.asp" -->

<% 	dim dicitura
	set dicitura = New testata 
	dicitura.sezione = "Gestione files - elenco"
	dicitura.puls_new = "INDIETRO A SITI"
	dicitura.link_new = "Siti.asp"
	if index.ChkPrm(prm_immaginiFormati_accesso, 0) then
		dicitura.iniz_sottosez(1)
		dicitura.sottosezioni(1) = "FORMATI IMMAGINI"
		dicitura.links(1) = "SitoImmaginiFormati.asp"
	else
		dicitura.iniz_sottosez(0)
	end if
	dicitura.scrivi_con_sottosez() %>

<iframe src="../library/filemanager.asp?FILEMAN_AZ_ID=<%= Session("AZ_ID") %>" frameborder="0" scrolling="No" 
		id="content_liquid" style="width:900px; height:450px;">
</iframe>

</html>

