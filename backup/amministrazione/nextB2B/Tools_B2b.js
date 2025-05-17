//funzione che calcola la variazione percentuale del valore Attuale rispetto al valore base 
//e la mette nell'input Variazione
function CalcolaVariazione(InputPrezzoBase, InputPrezzoAttuale, InputVariazione){
	var PrezzoBase = toNumber(InputPrezzoBase.value);
	var PrezzoAttuale = toNumber(InputPrezzoAttuale.value);
	
	InputPrezzoAttuale.value = FormatNumber(PrezzoAttuale, 2);
	InputVariazione.value = FormatNumber(((PrezzoAttuale-PrezzoBase)*10000/PrezzoBase)/100, 2);
}

//funzione che calcola la variazione in euro del valore attuale rispetto al valore base
// e la mette nell'input Variazione
function CalcolaDifferenza(InputPrezzoBase, InputPrezzoAttuale, InputVariazione){
	var PrezzoBase = toNumber(InputPrezzoBase.value);
	var PrezzoAttuale = toNumber(InputPrezzoAttuale.value);
	
	InputPrezzoAttuale.value = FormatNumber(PrezzoAttuale, 2);
	InputVariazione.value = FormatNumber(PrezzoAttuale-PrezzoBase, 2);
}

//funzione che ricalcola il valore attuale sulla base della variazione rispetto al 
//valore base e lo mette nell'input del PrezzoAttuale
function CalcolaPrezzo(InputPrezzoBase, InputPrezzoAttuale, InputVariazione){
	var PrezzoBase = toNumber(InputPrezzoBase.value);
	var Variazione = toNumber(InputVariazione.value);
	
	InputVariazione.value = FormatNumber(Variazione,2);
	InputPrezzoAttuale.value = FormatNumber((PrezzoBase*(100+Variazione))/100, 2);
}

//funzione che ricalcola il valore attuale sulla base della variazione rispetto al
//valore base e lo mette nell'input del prezzoattuale
function CalcolaPrezzoEuro(InputPrezzoBase, InputPrezzoAttuale, InputVariazione){
	var PrezzoBase = toNumber(InputPrezzoBase.value);
	var Variazione = toNumber(InputVariazione.value);
	
	InputVariazione.value = FormatNumber(Variazione,2);
	InputPrezzoAttuale.value = FormatNumber(PrezzoBase + Variazione, 2);
}

//controlla la quantita immessa nel dettaglio dell'ordine
//se conferma true chiedo conferma anche se la qta supera la giacenza
function ControllaQta(qta, giacenza, qtaMin, lottoR, conferma) {
	msg = ""
	if (qta > giacenza){
		msg += "La quantita' richiesta supera la giacenza disponibile.\n"
	}	
	if (qta < qtaMin) {
		msg += "La quantita' richiesta e' inferiore alla quantita' minima ordinabile.\n"
		conferma = false
	}

	var qtaOrd
	if (qtaMin < 2)
		qtaOrd = qta
	else
		qtaOrd = qta - qtaMin
		
	if ((qtaOrd % lottoR)!= 0 && lottoR != 0) {
		msg = msg + "La quantita' richiesta non e' conforme al lotto di riordino.\n"
		conferma = false
	}
	
	if (msg == "")
		return true
	else if (conferma)
		return confirm(msg + "Inserire comunque il dettaglio d'ordine?")
	else {
		alert(msg)
		return false
	}
}