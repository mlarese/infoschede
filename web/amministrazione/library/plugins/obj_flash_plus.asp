<!--#INCLUDE FILE="../Tools.asp"-->
<!--#INCLUDE FILE="../Tools4PlugIn.asp"-->
<!--#INCLUDE FILE="../ClassConfiguration.asp"-->
<!--#INCLUDE FILE="flash_scripts.asp"-->
<%
dim Config
set Config = new Configuration
'impostazione delle proprieta' di default
Config.AddDefault "movie", ""
Config.AddDefault "id", ""
Config.AddDefault "play", "true"
Config.AddDefault "loop", "true"
Config.AddDefault "quality", "high"
Config.AddDefault "scale", "showall"
Config.AddDefault "devicefont", "false"
Config.AddDefault "bgcolor", "#FFF"
Config.AddDefault "allowScriptAccess", ""
Config.AddDefault "width", ""
Config.AddDefault "height", ""
Config.AddDefault "align", ""
'caricamento proprieta' specifiche
Config.SetConfigurationString(Session("CONFSTR"))

%>
<!--web studio & creazioni multimediali next-Aim   -->
<script language="JavaScript" type="text/javascript">
<%
	dim movie
	if instr(1, Config("movie"), "http", vbTextCompare)>0 then
		movie = Config("movie")
	else
		movie = "upload/" & Application("AZ_ID") & "/images/" & Config("movie")
	end if
	if right(movie,4) = ".swf" then
		movie = left( movie, len(movie)-4)
	end if
	dim width,height
	if Config("width")<>"" then
		width = Config("width")
	else
		width = SESSION("LAYER_WIDTH") 
	end if
	if Config("height")<>"" then
		height = Config("height")
	else
		height = SESSION("LAYER_HEIGHT") 
	end if
%>
<!-- 
var hasRightVersion = DetectFlashVer(requiredMajorVersion, requiredMinorVersion, requiredRevision);
if(hasRightVersion) {  // se è stata rilevata una versione accettabile
	if (AC_FL_RunContent == 0) {
		alert("Questa pagina richiede AC_RunActiveContent.js. In Flash, selezionare \"Applica Aggiornamento per contenuto attivo\" nel menu Comandi per copiare AC_RunActiveContent.js nella cartella di output HTML.");
	} else {
		// incorpora il filmato Flash
		AC_FL_RunContent(
			'codebase', 'https://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,2,0',
			'width', '<%= width %>',
			'height', '<%= height %>',
			'src', '<%= movie %>',
			'quality', '<%= Config("quality") %>',
			'pluginspage', 'https://www.macromedia.com/go/getflashplayer',
			'align', '<%= Config("align") %>',
			'play', '<%= Config("play") %>',
			'loop', '<%= Config("loop") %>',
			'scale', '<%= Config("scale") %>',
			'wmode', 'window',
			'devicefont', '<%= Config("devicefont") %>',
			'id', '<%= Config("id") %>',
			'bgcolor', '<%= Config("bgcolor") %>',
			'name', '<%= Config("id") %>',
			'menu', 'true',
			'allowScriptAccess','sameDomain',
			'movie', '<%= movie %>',
			'salign', ''
			); //end AC code
	}
  } else {  // la versione di Flash è troppo vecchia o non è possibile rilevare il plug-in
    var alternateContent = 'Il contenuto HTML alternativo deve essere posizionato qui.'
  	+ 'Questo contenuto richiede Macromedia Flash Player.'
   	+ '<a href=https://www.macromedia.com/go/getflash/>Ottieni Flash</a>';
    document.write(alternateContent);  // Inserisci contenuto non Flash
  }
// -->
</script>
<noscript>
	// Fornisci contenuto alternativo per i browser che non supportano la creazione di script
	// o in cui la funzione di creazione di script è disabilitata.
  	Il contenuto HTML alternativo deve essere posizionato qui. Questo contenuto richiede Macromedia Flash Player.
  	<a href="https://www.macromedia.com/go/getflash/">Ottieni Flash</a>  	
</noscript>

