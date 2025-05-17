//variabile ausiliaria utilizzata come serbatoio dati temporaneo
var AuxVar

function OpenDeleteWindow(sezione, id){
	var pW = OpenAutoPositionedWindow("Delete.asp?SEZIONE=" + sezione + "&ID=" + id, "ELIMINA", 400, 270)
	return pW;
}

function OpenAutoPositionedScrollWindow(url, target, width, height, scrollbars){
	var top, left, except;
	try {
		//calcola coordinata Y
		//top = (event.screenY - event.offsetY) + 20;
		top = (window.screenY - window.offsetY) + 20;
		if ((top + height)>(screen.height-100))
			top = (screen.height - height - 100);
	
		//calcola coordinata X
		//left = (event.screenX - (width/2));
		left = (window.screenX - (width/2));
		if (left <20)
			left = 20;
		else if ((left + width)>(screen.width-20))
			left = (screen.width - width - 20);
	}
	catch(except){
		top = screenX + (outerWidth - innerWidth);
		left = screenY + (outerHeight - innerHeight);
	}
	
	var pW = OpenPositionedScrollWindow(url, target, left, top, width, height, scrollbars)
	return pW;
}	

function OpenAutoPositionedWindow(url, target, width, height){
	var pW = OpenAutoPositionedScrollWindow(url, target, width, height, false);
	return pW;
}

function OpenPositionedWindow(url, target, left, top, width, height){
	var pW = OpenPositionedScrollWindow(url, target, left, top, width, height, false)
	return pW;
}

function OpenPositionedScrollWindow(url, target, left, top, width, height, scrollbars){
	var properties = '';
	if ((height + top)>(screen.height-60)){
		height = screen.height - top - 60;
		width += 20;
	}
	
	if (width!=''){
        if (left!='' && top!=''){
            properties += 'left=' + left + ', top=' + top + ", "
        }
		properties += 'width=' + width + ', height=' + height + 
					  ', resizable=yes, status=yes, menubar=no, toolbars=no';
    }
	if (scrollbars)
		properties += ", scrollbars=yes";
	else
		properties += ", scrollbars=no";
		
	var pW = window.open(url, target, properties);
	return pW;
}


function OpenWindow(url, width, height){
    var pW = OpenPositionedScrollWindow(url, '_blank', '', '', width, height, true);
    return pW;
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


//Funzione per abilitare/disabilitare un controllo sulla base dello stato di un altro
function EnableIfChecked(obj_to_check, obj_to_enable){
	DisableControl(obj_to_enable, !(obj_to_check.checked))
}

//Funzione per abilitare/disabilitare un controllo sulla base dello stato di un altro
function DisableIfChecked(obj_to_check, obj_to_disable){
	DisableControl(obj_to_disable, obj_to_check.checked)
}

//Funzione per abilitare/disabilitare un controllo
function DisableControl(obj, disable){
	if (!disable) {
		//abilita il controllo
		if (obj.disabled || obj.disabled==undefined){
			if (obj.disabled){
				obj.disabled = false;
			}
			var re = /_disabled/;
			obj.className = obj.className.replace(re, '');
			var re = /disabled/;
			obj.className = obj.className.replace(re, '');
		}
	}else{
		//disabilita il controllo
		obj.disabled = true;
		if (obj.className.indexOf('disabled')<0){
			var SpacePos = obj.className.indexOf(' ')
			if (SpacePos<0){
				if (obj.className != '')
					obj.className += '_';
				obj.className += 'disabled';
			}
			else{
				obj.className = obj.className.substr(0, SpacePos) + '_disabled' + obj.className.substr(SpacePos);
			}
		}
	}
}


//Funzione per abilitare/disabilitare un controllo di tipo "picker" sulla base dello stato di un altro
function EnablePickerIfChecked(obj_to_check, picker_to_enable){
	DisablePicker(picker_to_enable, !(obj_to_check.checked))
}

//Funzione per abilitare/disabilitare un controllo di tipo "picker" sulla base dello stato di un altro
function DisablePickerIfChecked(obj_to_check, picker_to_disable){
	DisablePicker(picker_to_disable, obj_to_check.checked)
}

//Funzione per abilitare/disabilitare un controllo di tipo "picker" con pulsante visualiza/scegli o reset
function DisablePicker(objInput, disable){
	//disabilita input del valore
	DisableControl(objInput, disable);
	
	//disabilita visualizzatore
	var oView = document.getElementById("view_" + objInput.name);
	if (oView)
		DisableControl(oView, disable);
	
	//disabilita links
	var oScegli = document.getElementById("link_scegli_" + objInput.name);
	if (oScegli)
		DisableControl(oScegli, disable);
	var oScegli = document.getElementById(objInput.form.name + "_link_scegli_" + objInput.name);
	if (oScegli)
		DisableControl(oScegli, disable);
	var oReset = document.getElementById("link_reset_" + objInput.name);
	if (oReset)
		DisableControl(oReset, disable);
	var oReset = document.getElementById(objInput.form.name + "_link_reset_" + objInput.name);
	if (oReset)
		DisableControl(oReset, disable);
	var oVisualizza = document.getElementById("link_view_" + objInput.name);
	if (oVisualizza)
		DisableControl(oVisualizza, disable);
	var oVisualizza = document.getElementById(objInput.name + "_link");
	if (oVisualizza)
		DisableControl(oVisualizza, disable);
}

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

//restituisce il numero (in formato stringa) con i separatori delle migliaia
function numberWithCommas(x) {
	var parts = x.toString().split(".");
	parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
	return parts.join(".");
}

//funzione che ridimensiona la finestra corrente alla dimensione del proprio contenuto
function FitWindowSize(obj){
	var clientWidth = obj.document.documentElement.clientWidth;
	if (clientWidth == 0 || clientWidth == undefined)
		clientWidth = obj.document.body.clientWidth;

	var scrollWidth = obj.document.documentElement.scrollWidth;
	if (scrollWidth == 0 || scrollWidth == undefined)
		scrollWidth = obj.document.body.scrollWidth;
		
	var scrollHeight = obj.document.documentElement.scrollHeight;
	if (scrollHeight == 0 || scrollHeight == undefined)
		scrollHeight = obj.document.body.scrollHeight;
		
	var clientHeight = obj.document.documentElement.clientHeight;
	if (clientHeight == 0 || clientHeight == undefined)
		clientHeight = obj.document.body.clientHeight;
	//verifica se tutti i parametri interessati dal conteggio sono impostati
	if (clientWidth > 0 && clientWidth != undefined && 
	    scrollWidth > 0 && scrollWidth != undefined && 
	    scrollHeight > 0 && scrollHeight != undefined && 
	    clientHeight > 0 && clientHeight != undefined && 
	    obj.screen.availHeight > 0 && obj.screen.availHeight != undefined ){
		
		var displace_width = (obj.screen.width - obj.screen.availWidth) * 2;
		var displace_height = (obj.screen.height - obj.screen.availHeight) * 2;
		
		//calcola offset per ridimensionare la finestra
		var width_offset = scrollWidth - clientWidth
		var height_offset = displace_height + scrollHeight - clientHeight
		//controlla dimensione verticale finestra
		if ((height_offset + displace_height + clientHeight) > obj.screen.availHeight)
			height_offset = obj.screen.availHeight - displace_height - clientHeight;
			
		//verifica se la finestra sfora dallo schermo in altezza
		if ((obj.screenTop + clientHeight + height_offset) > obj.screen.availHeight){
			//la finestra sfora in altezza dallo schermo
			var top_offset = obj.screen.availHeight - displace_height - (obj.screenTop + clientHeight + height_offset);
			//sposta la finestra in alto per mantenere spazio sufficente
			if ((top_offset + obj.screenTop) > 0)
				obj.moveBy(0, top_offset)
			else
				obj.moveTo(obj.screenLeft, 0)
		}
		//ridimensiona la finestra
		obj.resizeBy(width_offset, height_offset);
	}
}

//funzione che imposta l'evento particolare al caricamento della finestra
function PageOnLoad_FocusSet(){
	window.onload = PageOnLoad_Focus;
}
		
//funzione che imposta il focus sul primo elemento
function PageOnLoad_Focus(){
	window.focus();
	var elemento = document.getElementById("primo_elemento");
	if (elemento){
		elemento.focus();
	}
}

//funzione che verifica la correttezza del colore
//il test viene effettuato sul colore di default dei link del documento perche'
//non pi&ugrave; utilizzato
function VerifyColor(color, advise_error){
	var ValidColor = true;
	
	color = color.toUpperCase( );
	
	if (color != "TRANSPARENT"){
		//verifica colore inserito
		if (color.length!=4 && color.length!=7)	
			ValidColor = false;
		
		//verifica primo carattere che deve essere #
		if (color.charAt(0)!="#")
			ValidColor = false;
		
		if (ValidColor){
			try{
				window.document.linkColor = color;
			}
			catch(e){
				ValidColor = false;
			}
		}
	}
	
	if (advise_error){
		if (!ValidColor)
			alert("Il codice colore HTML inserito non e' valido!");
	}
	
	return ValidColor;
	
}

//funzione che imposta il focus della finestra sul primo elemento con tabindex=1
function FocusOnFirstInput(form){
	for (var i = 0; i<form.length; i++){
		if(form[i].tabIndex == 1)
			form[i].focus();
	}
}

//funzione che imposta la dimensione dell'iframe che contiene la pagina corrente alla dimensione del contenuto del frame.
function SetParentFrameHeight(frameParentName){
	var pixels;
	var divContainer = document.getElementById('iframeform');
	if (divContainer){
		pixels = divContainer.offsetHeight;
	}
	else{
		pixels = document.body.scrollHeight;
	}
	//pixels-=5;
	parent.document.getElementById(frameParentName).height = pixels+"px";
}


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//funzioni per apertura immagini
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
	ImageWindow.document.write('\t<body bottommargin="0" leftmargin="0" marginheight="0" marginwidth="0" rightmargin="0" topmargin="0">\n');
	ImageWindow.document.write('\t\t<div style="width:100%; height:100%; text-align:center;">\n');
	ImageWindow.document.write('\t\t\t<img align="center" src="'+ image +'" onLoad="opener.OpenSmartImageResizer(this.width, this.height);" ondblclick="window.close();" alt="Doppio click per chiudere la finestra." title="Doppio click per chiudere la finestra.">\n');
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

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//funzioni per manipolazione HTML
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//.........................................................................................................
//funzione per l'ordinamento delle righe di una tabella
//  Table:          oggetto che contiene la tabella
//  SortMethod:     funzione per il metodo di ordinamento, da dichiarare in formato <nome_funzione>(row1, row2) 
//                  e deve restituire -1, 0, 1 in base al confrontro tra row1 e row2 il cui metodo 
//                  di confronto viene definito dal codice
//.........................................................................................................
function SortTable(Table, SortMethod){
    
    //genera copia dell'array delle righe
    var Rows = new Array();
    var r = 0;
    for (var r1 = 0; r1 < Table.rows.length; r1++, r++)
        Rows[r] = Table.rows[r1];
    
    if (Rows.length > 0) {
        //ordina array delle righe
        Rows.sort(SortMethod);
        
        //sostituisce vecchio array delle righe con nuovo array ordinato
        var RowsCopy = new Array(Rows.length)
        for (r = 0; r < Rows.length; r++) {
            RowsCopy[r] = Rows[r].cloneNode(true);
            Table.deleteRow(Rows[r].rowIndex);
        }
        
        var tableSection = Table.tBodies[Table.tBodies.length - 1];
        for (r = 0; r < Rows.length; r++) {
            var Row = RowsCopy[r];
            tableSection.appendChild(Row);
        }
    }
}


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// funzioni per google MAPS
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function addMarker(map, icon, draggable, point, lat, lng, zoom, title, description) {
	var p = point;
	if (p == null)
		p = new GLatLng(lat, lng);
		
	var marker;
	var opts = new Object();

	if(icon != null)
		opts.icon = icon;
		
	opts.draggable = draggable;
	opts.title = title;
	marker = new GMarker(p, opts);
		
	if (description != null)
		GEvent.addListener(marker, 'click', function() {
						   description.style.visibility = 'visible';
						   description.style.display = 'block';
						   marker.openInfoWindowHtml(description);
						   });
	
	if (zoom == null)
		map.setCenter(p);
	else
		map.setCenter(p, zoom);
	
	map.addOverlay(marker);
	return marker;
}


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
			//if (layers[i].id.indexOf("lay_") >= 0) {
				// explorer 7: se imposto l'altezza di un div ed il contenuto la supera non ridimensiona il div
				// escludo i layers perche caricati in un secondo momento via js e con IE7 non si ridimensionano
				if (mas < layers[i].offsetTop + layers[i].clientHeight)
					mas = layers[i].offsetTop + layers[i].clientHeight
			//}
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