<%@ Language=VBScript CODEPAGE=65001%>
<% option explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_categorie.asp" -->
<!--#INCLUDE FILE="../library/ClassUpload.asp" -->
<!--#INCLUDE FILE="../library/class_testata.asp" -->
<html>
<head>
	<title>Gestione documenti - file</title>
	<link rel="stylesheet" type="text/css" href="../library/stili.css">
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" onload="window.focus()">
<% 
	dim intestazione
	'disegna intestazione
	set intestazione= New testata
	intestazione.sezione = "Gestione documenti - file"
	intestazione.scrivi_ridotta()
%>

	<div id="content_ridotto">
		<%dim classe
		Set Classe = New UpLoad
		Classe.Connection_String = Application("DATA_ConnectionString")
		Classe.Table_Name		="mtb_documenti"								'nome della tabella
		Classe.ID_Field			="doc_id"										'campo Identity della tabella
		Classe.SQL_Nominativo	=" 'documento: ' + doc_titolo_it "			'Stringa sql per comporre il nome del record
		Classe.File_Field		="doc_FILE_it"									'Nome campo contenente il file
		Classe.File_Path		=Application("IMAGE_PATH") & "/docs/" 
					
		Classe.Border_color		="#919191"					'colore bordi tabelle
		Classe.Bg_testata		="#E6E6E6"					'colore sfondo bordi
		
		Classe.Stile_Input		=""			'stile input di testo
		Classe.Stile_Submit		="class=""button"""			'stile pulsanti submit
		Classe.Stile_Titoli		="class=""content_b"""		'stile testo titoli
		Classe.Stile_Testata	="class=""caption"""		'stile testo testata
		Classe.Stile_testo		=""							'stile testo normale
					
		Classe.Gestione_Completa_Record()
		if Classe.OperationOK then%>
				<script language="JavaScript">
						opener.location.reload(true);
				</script>
		<% end if%>
	</div>
</body>
</html>