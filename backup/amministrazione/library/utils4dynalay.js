
/****************************************************************** FUNZIONI PER LA GESTIONE DEGLI EVENTI E DELLE FINESTRE */
// imposta l'evento corrente
var currentEvent;
if (window.event) 
	currentEvent = window.event
else
	currentEvent = window.e;

//funzione per risolvere i problemi del target blank.
function OpenNewWindow(a, w, h){
	if (w==0 || h==0){
		a.target = '_blank';
	}
	else{
		try{
			if(event.shiftKey || event.shiftLeft){
				a.target = '_blank';
			}
			else{
				a.target = 'prova';
				return OpenSizedWindow(a.href, 'prova', w, h);
			}
		}
		catch(except){
			a.target = '_blank';
		}
	}
}

function OpenWindow(url, width, height){
	OpenSizedWindow(url, '_blank', width, height)
}

function OpenSizedWindow(url, target, width, height){
	var properties
	if (width=='')
		properties = ''
	else
		properties = 'width=' + width + ', height=' + height + ', resizable=yes, scrollbars=yes, status=yes, menubar=no, toolbars=no'
	var w = window.open(url, target, properties);
	if (!w.opener)
		w.opener = this;
	return void(0);
}

function OpenAutoPositionedSizedWindow(url, target, width, height){
	var properties
	var x, y;
	var top, left;
	try{
		//calcola coordinata Y
		y = (currentEvent.screenY - (currentEvent.offsetY ? currentEvent.offsetY : 0)) + 20;
		if ((y + height)>(screen.height-100))
			y = (screen.height - height - 100);
	
		//calcola coordinata X
		x = (currentEvent.screenX - (width/2));
		if (x < 20)
			x = 20;
		else if ((x + width)>(screen.width-20))
			x = (screen.width - width - 20);
		
		top = 'top=' + y + ', ';
		left = 'left=' + x + ', ';
	}
	catch(except){
		top='';
		left='';
	}
	
	if (width=='')
		properties = ''
	else
		properties = 'width=' + width + ', height=' + height + ', '
	properties += top + left + 'resizable=yes, scrollbars=yes, status=yes, menubar=no, toolbars=no'
	
	var w = window.open(url, target, properties);
	if (!w.opener)
		w.opener = this;
	return void(0);
}

function OpenPositionedScrollWindow(url, target, left, top, width, height, scrollbars){
	var properties
	if ((height + top)>(screen.height-60)){
		height = screen.height - top - 60;
		width += 20;
	}
	
	if (width=='')
		properties = '';
	else
		properties = 'left=' + left + ', top=' + top + ', width=' + width + ', height=' + height + 
					 ', resizable=yes, status=yes, menubar=no, toolbars=no';
	if (scrollbars)
		properties += ", scrollbars=yes";
	else
		properties += ", scrollbars=no";
		
	var pW = window.open(url, target, properties);
	return pW;
}

function OpenAutoPositionedScrollWindow(url, target, width, height, scrollbars){
	var top, left, except;
	try {
		//calcola coordinata Y
		top = (currentEvent.screenY - (currentEvent.offsetY ? currentEvent.offsetY : 0)) + 20;
		if ((top + height)>(screen.height-100))
			top = (screen.height - height - 100);
	
		//calcola coordinata X
		left = (currentEvent.screenX - (width/2));
		if (left <20)
			left = 20;
		else if ((left + width)>(screen.width-20))
			left = (screen.width - width - 20);
	}
	catch(except){
		top = 0;
		left = 0;
	}
	
	var pW = OpenPositionedScrollWindow(url, target, left, top, width, height, scrollbars)
	return pW;
}


function opensmartimage(image){
	OpenSmartImage(image);
}
var ImageWindowCount = 0;
var ImageWindow;

function OpenSmartImage(image){
	if (ImageWindowCount){
		ImageWindow.close();
		ImageWindow = null;
		ImageWindowCount = 0;
	}
	ImageWindow = window.open('', 'ImageWindow', 'width=300,height=250,resizable=yes,scrollbars=no');
	ImageWindowCount++;
	ImageWindow.document.open();
	ImageWindow.document.write('<html>\n');
	ImageWindow.document.write('\t<head>\n');
	ImageWindow.document.write('\t<meta http-equiv=Content-Type content="text/html; charset=utf-8">\n');
	ImageWindow.document.write('\t\t<title></title>\n');
	ImageWindow.document.write('\t</head>\n');
	ImageWindow.document.write('\t<body bottommargin="0" leftmargin="0" marginheight="0" marginwidth="0" rightmargin="0" topmargin="0" onload="focus();">\n');
	ImageWindow.document.write('\t\t<div style="width:100%; height:100%; text-align:center;">\n');
	if (image.indexOf("http") == -1)
		image = imageURL + image
	ImageWindow.document.write('\t\t\t<img align="center" src="'+ image +'" onLoad="opener.OpenSmartImageResizer(this.width, this.height);" ondblclick="window.close();" alt="'+ imageAlt +'" title="'+ imageAlt +'">\n');
	ImageWindow.document.write('\t\t</div>\n');
	ImageWindow.document.write('\t</body>\n');
	ImageWindow.document.write('</html>\n');
	ImageWindow.document.close();

	if (!ImageWindow.opener)
		ImageWindow.opener = this;
	return void(0);
}

function OpenSmartImageResizer(w, h){
	ImageWindow.resizeTo(w + 12, h + 61);
}

//Funzione per il debug: elenca tutte le proprieta' dell'oggetto passato come parametro
function displayproperties(obj){
	var w = window.open("", "properties","left=350, top=300, width=400,height=350, scrollbars=yes, menubar=no, toolsbar=no, status=no, resizable")
	var prop;
	var i=0;
	var str="<table border=\"1\" style=\"font:11px Courier New;\">";
	
	for (prop in obj){
		i +=1;
		str += "<tr>\n<td>\n\t" + i + "\n</td>\n<td>\n\t" + prop + "\n</td>\n<td><code>" + eval("obj." + prop) + "</code></td>\n</tr>\n";
	}
	str += "</table>"
	w.document.open();
	w.document.write(str);
	w.document.close();
}

var initWidth, initHeight;
// evento chiamato sull'onload dalla funzione ResizeWindow
function Fit() {
	window.resizeBy(MaxLayerWidth() - initWidth, MaxLayerHeight() - initHeight);
}
// ridimensiona la finestra partendo dalle dimensioni iniziali in input
function ResizeWindow(w, h) {
	initWidth = w, initHeight = h;
	window.moveTo(0, 0)
	window.onload = Fit
}

// restituisce il numero della versione dell'explorer
function getMSIEBrowserVersion() {
	try {
		var browser = navigator.appVersion.split(';')[1];
		if (browser.indexOf('MSIE') == -1)
			return -1;
		return parseInt(browser.substr(browser.lastIndexOf(' ')))
	} catch (e) {
		return -1;
	}
}
var MSIEVersione = getMSIEBrowserVersion();


//v1.1
//Copyright 2006 Adobe Systems, Inc. All rights reserved.
function AC_AX_RunContent(){
  var ret = AC_AX_GetArgs(arguments);
  AC_Generateobj(ret.objAttrs, ret.params, ret.embedAttrs);
}

function AC_AX_GetArgs(args){
  var ret = new Object();
  ret.embedAttrs = new Object();
  ret.params = new Object();
  ret.objAttrs = new Object();
  for (var i=0; i < args.length; i=i+2){
    var currArg = args[i].toLowerCase();    

    switch (currArg){	
      case "pluginspage":
      case "type":
      case "src":
        ret.embedAttrs[args[i]] = args[i+1];
        break;
      case "data":
      case "codebase":
      case "classid":
      case "id":
      case "onafterupdate":
      case "onbeforeupdate":
      case "onblur":
      case "oncellchange":
      case "onclick":
      case "ondblClick":
      case "ondrag":
      case "ondragend":
      case "ondragenter":
      case "ondragleave":
      case "ondragover":
      case "ondrop":
      case "onfinish":
      case "onfocus":
      case "onhelp":
      case "onmousedown":
      case "onmouseup":
      case "onmouseover":
      case "onmousemove":
      case "onmouseout":
      case "onkeypress":
      case "onkeydown":
      case "onkeyup":
      case "onload":
      case "onlosecapture":
      case "onpropertychange":
      case "onreadystatechange":
      case "onrowsdelete":
      case "onrowenter":
      case "onrowexit":
      case "onrowsinserted":
      case "onstart":
      case "onscroll":
      case "onbeforeeditfocus":
      case "onactivate":
      case "onbeforedeactivate":
      case "ondeactivate":
        ret.objAttrs[args[i]] = args[i+1];
        break;
      case "width":
      case "height":
      case "align":
      case "vspace": 
      case "hspace":
      case "class":
      case "title":
      case "accesskey":
      case "name":
      case "tabindex":
        ret.embedAttrs[args[i]] = ret.objAttrs[args[i]] = args[i+1];
        break;
      default:
        ret.embedAttrs[args[i]] = ret.params[args[i]] = args[i+1];
    }
  }
  return ret;
}

//v1.0
//Copyright 2006 Adobe Systems, Inc. All rights reserved.
function AC_AddExtension(src, ext)
{
  if (src.indexOf('?') != -1)
    return src.replace(/\?/, ext+'?'); 
  else
    return src + ext;
}

function AC_Generateobj(objAttrs, params, embedAttrs) 
{ 
  var str = '<object ';
  for (var i in objAttrs)
    str += i + '="' + objAttrs[i] + '" ';
  str += '>';
  for (var i in params)
    str += '<param name="' + i + '" value="' + params[i] + '" /> ';
  str += '<embed ';
  for (var i in embedAttrs)
    str += i + '="' + embedAttrs[i] + '" ';
  str += ' ></embed></object>';

  document.write(str);
}

function AC_FL_RunContent(){
  var ret = 
    AC_GetArgs
    (  arguments, ".swf", "movie", "clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"
     , "application/x-shockwave-flash"
    );
  AC_Generateobj(ret.objAttrs, ret.params, ret.embedAttrs);
}

function AC_SW_RunContent(){
  var ret = 
    AC_GetArgs
    (  arguments, ".dcr", "src", "clsid:166B1BCA-3F9C-11CF-8075-444553540000"
     , null
    );
  AC_Generateobj(ret.objAttrs, ret.params, ret.embedAttrs);
}

function AC_GetArgs(args, ext, srcParamName, classid, mimeType){
  var ret = new Object();
  ret.embedAttrs = new Object();
  ret.params = new Object();
  ret.objAttrs = new Object();
  for (var i=0; i < args.length; i=i+2){
    var currArg = args[i].toLowerCase();    

    switch (currArg){	
      case "classid":
        break;
      case "pluginspage":
        ret.embedAttrs[args[i]] = args[i+1];
        break;
      case "src":
      case "movie":	
        args[i+1] = AC_AddExtension(args[i+1], ext);
        ret.embedAttrs["src"] = args[i+1];
        ret.params[srcParamName] = args[i+1];
        break;
      case "onafterupdate":
      case "onbeforeupdate":
      case "onblur":
      case "oncellchange":
      case "onclick":
      case "ondblClick":
      case "ondrag":
      case "ondragend":
      case "ondragenter":
      case "ondragleave":
      case "ondragover":
      case "ondrop":
      case "onfinish":
      case "onfocus":
      case "onhelp":
      case "onmousedown":
      case "onmouseup":
      case "onmouseover":
      case "onmousemove":
      case "onmouseout":
      case "onkeypress":
      case "onkeydown":
      case "onkeyup":
      case "onload":
      case "onlosecapture":
      case "onpropertychange":
      case "onreadystatechange":
      case "onrowsdelete":
      case "onrowenter":
      case "onrowexit":
      case "onrowsinserted":
      case "onstart":
      case "onscroll":
      case "onbeforeeditfocus":
      case "onactivate":
      case "onbeforedeactivate":
      case "ondeactivate":
      case "type":
      case "codebase":
        ret.objAttrs[args[i]] = args[i+1];
        break;
      case "width":
      case "height":
      case "align":
      case "vspace": 
      case "hspace":
      case "class":
      case "title":
      case "accesskey":
      case "name":
      case "id":
      case "tabindex":
        ret.embedAttrs[args[i]] = ret.objAttrs[args[i]] = args[i+1];
        break;
      default:
        ret.embedAttrs[args[i]] = ret.params[args[i]] = args[i+1];
    }
  }
  ret.objAttrs["classid"] = classid;
  if (mimeType) ret.embedAttrs["type"] = mimeType;
  return ret;
}

////////////////////////////////////////////////////////////////////// FINE ADOBE

// restituisce la coordinata x del punto piu largo fra i layers della dynalay
function MaxLayerWidth() {
	var mas = 0
	var layers = document.getElementsByTagName("div")
	for(var i = 0; i < layers.length; i++)
		if (layers[i].id.indexOf("lay_") >= 0) {
			if (mas < layers[i].offsetLeft + layers[i].clientWidth)
				mas = layers[i].offsetLeft + layers[i].clientWidth
		}
	return mas
}
// restituisce la coordinata y del punto piu basso fra i layers della dynalay
var maxLayerHeightStored = 0;

function MaxLayerHeight() {
	if (maxLayerHeightStored < 1){
		var mas = 0
		var layers = document.getElementsByTagName("div")
		for(var i = 0; i < layers.length; i++)
			if (layers[i].id.indexOf("lay_") >= 0) {
				// explorer 7: se imposto l'altezza di un div ed il contenuto la supera non ridimensiona il div
				// escludo i layers perche caricati in un secondo momento via js e con IE7 non si ridimensionano
				if (navigator.userAgent.indexOf("MSIE 6.0") == -1 && layers[i].className.indexOf("layers_flash") == -1) {
					layers[i].style.minHeight = layers[i].style.height
					layers[i].style.height = "auto"
					//layers[i].style.border = "1px solid red";
				}
				if (mas < layers[i].offsetTop + layers[i].clientHeight)
					mas = layers[i].offsetTop + layers[i].clientHeight
			}
		maxLayerHeightStored = mas;
		return mas
	}
	else{
		return maxLayerHeightStored;
	}
}


/****************************************************************** FLOAT LAYERS WITH SCROLL */
// array di layer da spostare
var FloatLayers = new Array();
var prevScrollTop = 0;

// restituisce la posizione orizzontale dell'elemento in input
function getXCoord(el) {
	x = 0;
	while(el) {
		x += el.offsetLeft;
		el = el.offsetParent;
	}
	return x;
}
// restituisce la posizione verticale dell'elemento in input
function getYCoord(el) {
	y = 0;
	while(el) {
		y += el.offsetTop;
		el = el.offsetParent;
	}
	return y;
}

// rende un layer spostabile
function floatMaking(layerName, x, y, speed) {
    // aggangia il metodo di spostamento agli eventi pagina
	window.onresize = floatAllLayers;
	window.onscroll = floatAllLayers;
	
	var layer = document.getElementById(layerName);
	// setta alcune nuove proprieta al layer
	layer.prevX = getXCoord(layer);
	layer.prevY = getYCoord(layer);
	layer.floatX = x;
	layer.floatY = y;
	layer.steps = speed;
	layer.ifloatX = Math.abs(x);        // coordinata x iniziale
	layer.ifloatY = Math.abs(y);        // coordinata y iniziale
	layer.tm = null;                    // funzione in timeout (serve per evitare concorrenza)
	layer.style.position = 'absolute';
	// resetta le impostazioni del layer per calcolare quelle effettive (gestione specifica per i div della dynalay)
	layer.style.height = "auto";
	
	FloatLayers.push(layer);
	floatLayer(layer);
}

// sposta tutti i layer
function floatAllLayers() {
	for(var i = 0; i < FloatLayers.length; i++)
		floatLayer(FloatLayers[i]);
	
	// salvo lo scroll precedente per capire se sto scendendo o salendo
	if (document.documentElement)
        prevScrollTop = document.documentElement.scrollTop
    else
        prevScrollTop = document.body.scrollTop;
}
// calcola ed avvia lo spostamento (da ottimizzare quello orizzontale) del layer (div) in input e controlla che l'altezza non superi la pagina
function floatLayer(layer) {
    var marginY = 100;      // margine di visualizzazione sotto al layer
    
	// calcolo lo spostamento
	if (document.documentElement) {		// doctype transitional
	    var windowHeight;
        if (navigator.userAgent.indexOf('Opera') >= 0)
            windowHeight = document.body.clientHeight;
        else
            windowHeight = document.documentElement.clientHeight;
        
	    if (prevScrollTop > document.documentElement.scrollTop
	        || layer.prevY + layer.clientHeight + marginY < windowHeight + document.documentElement.scrollTop) {
		    // left
		    layer.floatX = document.documentElement.scrollLeft + layer.ifloatX;
			
		    // top
		    if (layer.clientHeight > windowHeight)
		        layer.floatY = document.documentElement.scrollTop;
		    else
		        layer.floatY = document.documentElement.scrollTop + layer.ifloatY;
		    /*
		    // right
		    layer.floatX = document.documentElement.scrollLeft + document.documentElement.clientWidth - layer.ifloatX - layer.offsetWidth;
		    // bottom
		    layer.floatY = document.documentElement.scrollTop + document.documentElement.clientHeight - layer.ifloatY - layer.offsetHeight;
		    */
		}
	} else {
	    if (prevScrollTop > document.body.scrollTop
	        || layer.prevY + layer.clientHeight + marginY < document.body.clientHeight + document.body.scrollTop) {
		    // left
		    layer.floatX = document.body.scrollLeft + layer.ifloatX;
			
		    // top
		    if (layer.clientHeight > document.body.clientHeight)
	            layer.floatY = document.body.scrollTop;
	        else
	            layer.floatY = document.body.scrollTop + layer.ifloatY;
		    /*
		    // right
		    layer.floatX = document.body.scrollLeft + document.body.clientWidth - layer.ifloatX - layer.offsetWidth;
		    // bottom
		    layer.floatY = document.body.scrollTop + document.body.clientHeight - layer.ifloatY - layer.offsetHeight;
		    */
        }
	}
	
	// controllo che non sfori mai le posizioni iniziali
	if (layer.floatY < layer.ifloatY)
	    layer.floatY = layer.ifloatY
	if (layer.floatX < layer.ifloatX)
	    layer.floatX = layer.ifloatX
	// controllo che non sfori mai l'altezza della pagina
	var documentHeight = MaxLayerHeight();
	if (layer.floatY + layer.clientHeight > documentHeight)
	    layer.floatY = documentHeight - layer.clientHeight
	
	// avvio lo spostamento
	if(layer.prevX != layer.floatX || layer.prevY != layer.floatY)
		if (layer.tm == null)
		    layer.tm = setTimeout(function () { floatDo(layer); }, 50);
}
// sposta il layer in input (solo in verticale)
function floatDo(layer) {
    layer.tm = null;
    if (layer.style.position != 'absolute') return;

    var dx = Math.abs(layer.floatX - layer.prevX);
	var dy = Math.abs(layer.floatY - layer.prevY);
	
	if (dx < layer.steps / 2)
		cx = (dx >= 1) ? 1 : 0;
	else
		cx = Math.round(dx / layer.steps);
	if (dy < layer.steps / 2)
		cy = (dy >= 1) ? 1 : 0;
	else
		cy = Math.round(dy / layer.steps);

	if (layer.floatX > layer.prevX)
		layer.prevX += cx;
	else if (layer.floatX < layer.prevX)
		layer.prevX -= cx;
	if (layer.floatY > layer.prevY)
		layer.prevY += cy;
	else if (layer.floatY < layer.prevY)
		layer.prevY -= cy;

    //layer.style.left = layer.prevX + "px";
	layer.style.top = layer.prevY + "px";

	if (cx != 0 || cy != 0)
		if (layer.tm == null)
		    layer.tm = setTimeout(function () { floatDo(layer); }, 50);
}
/********************************************************************************************/


/*******************************************************************************************
Funzione che allunga il layer indicato fino alla fine della pagina o della finestra 
(considerando la maggiore delle due), eventualmente aggiungendo l'offset indicato
*******************************************************************************************/
function ResizeToBottom(layer, contentOffset, divBottomOffset){

	//determina altezza della finestra
	var windowHeight;
       windowHeight = document.documentElement.clientHeight;
		
	//determina altezza totale dei contenuti
	var contentHeight = MaxLayerHeight() + contentOffset;
	
	//determina l'altezza a cui fare riferimento per il resize.
	var currentHeight;
	if (contentHeight > windowHeight){
		currentHeight = contentHeight;
	}
	else{
		currentHeight = windowHeight;
	}
	
	//determina nuova dimensione layer
	var newDivHeight = currentHeight - getYCoord(layer);
	$('#' + layer.id).animate({height: (newDivHeight + divBottomOffset) + "px"}, 10 );

	//var ef = new Effect.Morph(layer.id, {style: "height:" + (newDivHeight + divBottomOffset) + "px;", duration: 0.00, transition: Effect.Transitions.full});
}

/*******************************************************************************************
Funzioni e variabili che gestiscono l'allungamento ed il mantenimento della lunghezza dei layer
*******************************************************************************************/

//elenco dei layer da ridimensionare fino alla fine dell'elemento
var ResizedToBottomLayers = new Array();
var ResizedToBottomContentOffset = new Array();
var ResizedToBottomDivBottomOffset = new Array();
var ResizedToBottomWindowHeight = 0;

//imposta il layer indicato in modo che venga mantenuto fisso alla dimensione corretta.
function SetContinuousResizeToBottom(layer, contentOffset, divBottomOffset){
	window.onresize = ResizeToBottomAllLayers;
	window.onafterupdate = ResizeToBottomAllLayers;
	
	ResizedToBottomLayers.push(layer);
	ResizedToBottomContentOffset.push(contentOffset);
	ResizedToBottomDivBottomOffset.push(divBottomOffset);
	
	ResizeToBottom(layer, contentOffset, divBottomOffset);
}

//ridimensiona tutti i layer secondo le impostazioni di offset
function ResizeToBottomAllLayers(){
	for(var i = 0; i < ResizedToBottomLayers.length; i++){
		if (ResizedToBottomWindowHeight != document.documentElement.clientHeight){
			ResizeToBottom(ResizedToBottomLayers[i], ResizedToBottomContentOffset[i], ResizedToBottomDivBottomOffset[i]);
		}
	}
	ResizedToBottomWindowHeight = document.documentElement.clientHeight;
	
}

/*******************************************************************************************
Funzioni che convertono numeri e stringhe e viceversa
*******************************************************************************************/
//funzione che restituisce il valore numerico o NaN se se il valore non e' numerico
function toNumber(value){
	var str = value.toString();
	var value = parseFloat(str.replace(",","."));
	if (isNaN(value))
		return 0;
	else
		return value;	
}

//funzione che restituisce il numero formattato con le cifre decimali richieste
function FormatNumber(value, decimal){
	var num = toNumber(value);
	var num = num.toFixed(decimal);
	var str = num.toString();
	return str.replace('.',',');
}

/*******************************************************************************************
Funzioni che convertono date e stringhe e viceversa
*******************************************************************************************/

//Funzione che converte una stringa nel corrispondente valore data.
function stringToDate(value, language){
	var d = new String(value);
	var day, month, year;
	var splitted = d.split("/");
	if (language == 'en'){
		day = splitted[1];
		month = splitted[0];
	}
	else{
		day = splitted[0];
		month = splitted[1];
	}
	if (splitted[2].length>4)
		year = splitted[2].substr(0,4);
	else
		year = splitted[2];
	var date = new Date(year +"/"+ month +"/"+ day);
	
	if (date.getFullYear() == year && 
		parseFloat(date.getMonth()+1) == parseFloat(month) && 
		parseFloat(date.getDate()) == parseFloat(day)){
		return date;
	}
	else
		return "";
}

function lz(numero, cifre) {
	n = String(numero);
	while (n.length<cifre) { 
		n="0"+n 
	}
	return n;
}
function dateFormat(data, formato) { 
// (c) br1 - 2002 
 
	var giorni = new Array("Domenica","Lunedì","Martedì","Mercoledì","Giovedì","Venerdì","Sabato");
	var mesi = new Array("Gennaio","Febbraio","marzo","Aprile","Maggio","Giugno","Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre");
 
// preparo la data...  verificare di passarla corretta!
	var adesso = new Date(data); 
	var anno = adesso.getFullYear();
	var mese = adesso.getMonth()+1;
	var giorno = adesso.getDate();
	var settim = adesso.getDay()
	var ore = adesso.getHours()
	var minuti = adesso.getMinutes()
	var secondi = adesso.getSeconds()
 
// preparo la stringa di risposta
	var rVal = '';
 
	if (formato.length==0) { 
// assenza del secondo parametro
		return String(adesso); 
	} else {
 
	// inizio loop
		while (formato.length>0) {
 
	// verifico se c'e' qualche separatore e lo aggiungo
			while (formato.length>0 && String("ymdphnst").indexOf(formato.charAt(0).toLowerCase())<0) {
				rVal += formato.charAt(0);
				formato = formato.substr(1);
			}
 
 
	// Separo il gruppo
			if (formato.length>0) {
				ff = formato.charAt(0);
				formato = formato.substr(1);
				while (formato.length>0 && formato.charAt(0).toLowerCase()==ff.charAt(0).toLowerCase()) {
					ff += formato.charAt(0);
					formato = formato.substr(1);
				}
 
	// espando il formato nella stringa corrispondente
				ff = ff.toLowerCase();	 // operazione preliminare... tutto in minuscolo
				switch (ff) 	{ 
					case "yy" : 
						rVal += String(anno).substr(2); 
						break; 
					case "yyyy" : 
						rVal += String(anno); 
						break; 
					case "m" : 
						rVal += String(mese); 
						break; 
					case "mm" : 
						rVal += lz(mese,2);
						break; 
					case "mmm" : 
						rVal += mesi[mese-1].substr(0,3);
						break; 
					case "mmmm" : 
						rVal += mesi[mese-1];
						break; 
					case "d" : 
						rVal += String(giorno); 
						break; 
					case "dd" : 
						rVal += lz(giorno,2); 
						break; 
					case "ddd" : 
						rVal += giorni[settim].substr(0,3);
						break; 
					case "dddd" : 
						rVal += giorni[settim];
						break; 
					case "p" : 
						var inizio = new Date(anno, 0, 0); 
						rVal += Math.floor((adesso - inizio) / 86400000);
						break; 
					case "ppp" : 
						var inizio = new Date(anno, 0, 0); 
						rVal += lz(Math.floor((adesso - inizio) / 86400000),3);
						break; 
					case "h" : 
						rVal += String(ore); 
						break; 
					case "hh" : 
						rVal += lz(ore,2); 
						break; 
					case "n" : 
						rVal += String(minuti); 
						break; 
					case "nn" : 
						rVal += lz(minuti,2); 
						break; 
					case "s" : 
						rVal += String(secondi); 
						break; 
					case "ss" : 
						rVal += lz(secondi,2); 
						break; 
					case "t" : 
						rVal += lz(ore,2)+":"+lz(minuti,2)+":"+lz(secondi,2); 
						break; 
					default :  // il numero dei caratteri del formato non e' permesso
						rVal += ff.replace(/./gi,"?");
				} 
 
			}
 
		} // fine loop principale
 
		return rVal;
	}
} 

/*******************************************************************************************
Funzione per il controllo in input dei dati
*******************************************************************************************/
function onlyNumbersInput(evento){
	//var evtobj=window.event? window.event : e //distinguish between IE's explicit event object (window.event) and Firefox's implicit.
	//var unicode=evtobj.charCode? evtobj.charCode : evtobj.which
	var unicode = evento.which;
	var actualkey=String.fromCharCode(unicode);
	//unicode = 8 oppure = 0 mi serve per non escludere i tasti Canc e Backspace
	if (!(actualkey=="0" ||
		  actualkey=="1" ||
		  actualkey=="2" ||
		  actualkey=="3" ||
		  actualkey=="4" ||
		  actualkey=="5" ||
		  actualkey=="6" ||
		  actualkey=="7" ||
		  actualkey=="8" ||
		  actualkey=="9" ||
		  unicode=="8" ||
		  unicode=="0"
		  ))
	{
		//evento.which = 0;
		return false;
	}
}


/*******************************************************************************************
Funzioni che lavorano sulla dynalay
*******************************************************************************************/
function ChooseStringByLanguage() {
	var ret = "";
	switch (currentLanguage) {
		case 'it':
			ret = arguments[0];
			break;
		case 'en':
			if (arguments.length > 1) ret = arguments[1];
			break;
		case 'de':
			if (arguments.length > 2) ret = arguments[2];
			break;
		case 'fr':
			if (arguments.length > 3) ret = arguments[3];
			break;
		case 'es':
			if (arguments.length > 4) ret = arguments[4];
			break;
		case 'ru':
			if (arguments.length > 5) ret = arguments[5];
			break;
		case 'cn':
			if (arguments.length > 6) ret = arguments[6];
			break;
		case 'pt':
			if (arguments.length > 7) ret = arguments[7];
			break;
	}
	if (ret=="" && arguments.length > 1 && currentLanguage != 'it')
		ret = arguments[1];
	if (ret=="")
		ret = arguments[0];
	return ret;
}



/*******************************************************************************************
Imposta stili legati al browser: USA JQUERY: integrato dalla next-page del nextweb5
*******************************************************************************************/
 function setCssClassBrowser(){ 
	var myelements = $("div.layers_object"); 
	var navUserAgent = navigator.userAgent; 
	var browserType; 
	var indexofName; 
	 if (navUserAgent.indexOf("MSIE") != -1) { 
		 indexofName = navUserAgent.indexOf("MSIE"); 
		 browserType = 'ie' + navUserAgent.substring(indexofName + 5, navUserAgent.indexOf(".", indexofName)); 
	 } 
	 else if (navUserAgent.indexOf("Firefox") != -1) { 
		 browserType = 'firefox'; 
	} 
	 else if (navUserAgent.indexOf("iPad") != -1) { 
		 browserType = 'ipad'; 
	 } 
	 else if (navUserAgent.indexOf("iPhone") != -1) { 
		 browserType = 'iphone'; 
	 } 
	 else if (navUserAgent.indexOf("Chrome") != -1) { 
		 browserType = 'chrome'; 
	 } 
	 else if (navUserAgent.indexOf("Safari") != -1) { 
		 browserType = 'safari'; 
	 } 
	 else if (navUserAgent.indexOf("Opera") != -1) { 
		 browserType = 'opera'; 
	 } 
	 else { 
		 browserType = 'other'; 
	 } 
	 for (i = 0; i < myelements.length; i++) { 
		 var myarray = myelements[i].className.split(" "); 
		 for (j = 0; j < myarray.length; j++) { 
			 if (myarray[j] != 'layers_object') { 
				 myelements[i].className += " " + browserType + "_" + myarray[j]; 
			 } 
		 } 
	 } 
	 var myelements = $("body"); 
	 myelements[0].className += " " + browserType; 
 }