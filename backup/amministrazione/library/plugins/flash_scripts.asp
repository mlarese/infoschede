<script language="javascript"> AC_FL_RunContent = 0; </script>
<script src="<%= GetLibraryPath() %>plugins/AC_RunActiveContent.js" language="javascript"></script>
<script language="JavaScript" type="text/javascript">
<!--
// -----------------------------------------------------------------------------
// Globali
// È richiesta la versione principale di Flash
var requiredMajorVersion = 6;
// È richiesta la versione minore di Flash
var requiredMinorVersion = 0;
// Versione di Flash richiesta
var requiredRevision = 2;
// -----------------------------------------------------------------------------
// -->
</script>
<script language="VBScript" type="text/vbscript">
<!-- // Helper di Visual Basic richiesto per rilevare le informazioni sulla versione dei controlli ActiveX di Flash Player
Function VBGetSwfVer(i)
  on error resume next
  Dim swControl, swVersion
  swVersion = 0
  
  set swControl = CreateObject("ShockwaveFlash.ShockwaveFlash." + CStr(i))
  if (IsObject(swControl)) then
    swVersion = swControl.GetVariable("$version")
  end if
  VBGetSwfVer = swVersion
End Function
// -->
</script>
<script language="JavaScript1.1" type="text/javascript">
<!-- 
// Rileva tipo di browser client
var isIE  = (navigator.appVersion.indexOf("MSIE") != -1) ? true : false;
var isWin = (navigator.appVersion.toLowerCase().indexOf("win") != -1) ? true : false;
var isOpera = (navigator.userAgent.indexOf("Opera") != -1) ? true : false;
// Helper di JavaScript richiesto per rilevare le informazioni sulla versione del plug-in Flash Player
function JSGetSwfVer(i){
	// Le versioni di NS/Opera dalla 3 in poi verificano la presenza del plug-in Flash nell'array dei plug-in
	var flashVer = -1;
	if (navigator.plugins != null && navigator.plugins.length > 0) {
		if (navigator.plugins["Shockwave Flash 2.0"] || navigator.plugins["Shockwave Flash"]) {
			var swVer2 = navigator.plugins["Shockwave Flash 2.0"] ? " 2.0" : "";
      		var flashDescription = navigator.plugins["Shockwave Flash" + swVer2].description;
			var descArray = flashDescription.split(" ");
			var tempArrayMajor = descArray[2].split(".");
			var versionMajor = tempArrayMajor[0];
			var versionMinor = tempArrayMajor[1];
			if ( descArray[3] != "" ) {
				tempArrayMinor = descArray[3].split("r");
			} else {
				tempArrayMinor = descArray[4].split("r");
			}
      		var versionRevision = tempArrayMinor[1] > 0 ? tempArrayMinor[1] : 0;
            var flashVer = versionMajor + "." + versionMinor + "." + versionRevision;
		}
	}
	// MSN/WebTV 2.6 supporta Flash 4
	else if (navigator.userAgent.toLowerCase().indexOf("webtv/2.6") != -1) flashVer = 4;
	// WebTV 2.5 supporta Flash 3
	else if (navigator.userAgent.toLowerCase().indexOf("webtv/2.5") != -1) flashVer = 3;
	// Le versioni precedenti di WebTV supportano Flash 2
	else if (navigator.userAgent.toLowerCase().indexOf("webtv") != -1) flashVer = 2;
	return flashVer;
} 
// Se chiamato con il parametro reqMajorVer, reqMinorVer, reqRevision restituisce true se quella versione o una versione successiva è disponibile
function DetectFlashVer(reqMajorVer, reqMinorVer, reqRevision) 
{
 	reqVer = parseFloat(reqMajorVer + "." + reqRevision);
   	// Esamina ciclicamente all'indietro le versioni fino a trovare quella più recente	
	for (i=25;i>0;i--) {	
		if (isIE && isWin && !isOpera) {
			versionStr = VBGetSwfVer(i);
		} else {
			versionStr = JSGetSwfVer(i);		
		}
		if (versionStr == -1 ) { 
			return false;
		} else if (versionStr != 0) {
			if(isIE && isWin && !isOpera) {
				tempArray         = versionStr.split(" ");
				tempString        = tempArray[1];
				versionArray      = tempString .split(",");				
			} else {
				versionArray      = versionStr.split(".");
			}
			var versionMajor      = versionArray[0];
			var versionMinor      = versionArray[1];
			var versionRevision   = versionArray[2];
			
			var versionString     = versionMajor + "." + versionRevision;   // 7.0r24 == 7.24
			var versionNum        = parseFloat(versionString);
        	// è la versione maggiore >= versione maggiore richiesta E la versione minore >= versione minore richiesta
			if (versionMajor > reqMajorVer) {
				return true;
			} else if (versionMajor == reqMajorVer) {
				if (versionMinor > reqMinorVer)
					return true;
				else if (versionMinor == reqMinorVer) {
					if (versionRevision >= reqRevision)
						return true;
				}
			}
			return false;
		}
	}
}
// -->
</script>