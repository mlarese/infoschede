<%
dim ImageAlt
if Session("lingua") = "en" then
	ImageAlt = "Double click to close this window."
elseif Session("lingua") = "fr" then
	ImageAlt = "Double click pour fermer cette fenêtre."
elseif Session("lingua") = "de" then
	ImageAlt = "Doppeltes click zum Schließen dieses Fensters."
elseif Session("lingua") = "es" then
	ImageAlt = "Click doble para cerrar esta ventana."
else
	ImageAlt = "Doppio click per chiudere la finestra."
end if
%>
// variabili globali
var baseURL = "<%= "http://" & Application("SERVER_NAME") & "/" %>"
var imageURL = "<%= "http://" & Application("IMAGE_SERVER") & "/" & Session("AZ_ID") & "/images/" %>"
var imageAlt = "<%= imageAlt %>"


<!--#INCLUDE FILE="Utils4Dynalay.js" -->

function OpenPage(page, width, height){
	OpenAutoPositionedSizedWindow(baseURL + "dynalay.asp?PAGINA=" + page.toString(), "_blank", width, height)
}

 
function openimage(image, width, height){
	OpenImage(image, width, height);
}


function OpenImage(image, width, height){
	if (width=='')
		properties = ''
	else
		properties = 'width=' + width + ', height=' + height + ', resizable=yes, scrollbars=yes, status=yes, menubar=no, toolbars=no'
	var w = window.open(imageURL + image, '_blank', properties);
	if (!w.opener)
		w.opener = this;
	return void(0);
}


