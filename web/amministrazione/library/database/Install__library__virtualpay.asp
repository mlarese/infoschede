<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per l'applicativo VirtualPay
'...........................................................................................
'...........................................................................................

'*******************************************************************************************
'Installazione Applicazione VirtualPay
'...........................................................................................
' Luca 24/09/2015
'...........................................................................................
function Aggiornamento__VirtualPay__1(conn)
	Aggiornamento__VirtualPay__1 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__VirtualPay__1(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & BANCA_VIRTUALPAY)) <> "" then
		CALL AddParametroSito(conn, "VIRTUALPAY_TEST", _
									0, _
									"Indica se il sistema di pagamento è attivo in modalità di test o meno", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									1, null, null, null, null)
		CALL AddParametroSito(conn, "VIRTUALPAY_URL", _
									0, _
									"Url di invio del pagamento (Via POST)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "VIRTUALPAY_MERCHANT_ID", _
									0, _
									"Identificatore del negozio virtuale", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "VIRTUALPAY_DIVISA", _
									0, _
									"Valuta codice ISO International Standard 4217 (EUR=Euro)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									"EUR", "EUR", "EUR", "EUR", "EUR")
		CALL AddParametroSito(conn, "VIRTUALPAY_ABI", _
									0, _
									"ABI della banca alla quale deve essere indirizzata la transazione", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "VIRTUALPAY_URLOK", _
									0, _
									"Pagina del negozio virtuale verso la quale redirigere il cliente a pagamento avvenuto", _
									"", _
									adGUID, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "VIRTUALPAY_URLKO", _
									0, _
									"Pagina del negozio virtuale verso la quale redirigere il cliente se esso decide di interrompere l’operazione di pagamento", _
									"", _
									adGUID, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "VIRTUALPAY_URLACK", _
									0, _
									"Pagina del negozio virtuale verso la quale effettuare la GET di notifica di pagamento andato a buon fine", _
									"", _
									adGUID, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "VIRTUALPAY_URLNACK", _
									0, _
									"Pagina del negozio virtuale verso la quale effettuare la GET di notifica di pagamento non andato a buon fine per mancata autorizzazione", _
									"", _
									adGUID, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "VIRTUALPAY_MAC", _
									0, _
									"Message Authentication Code che il negozio virtuale può utilizzare per certificare l’integrità dei dati ordine", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "VIRTUALPAY_BUTTON_IMAGE", _
									0, _
									"Immagine utilizzata per il pulsante", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "VIRTUALPAY_BUTTON_TEXT", _
									0, _
									"Testo utilizzato per il pulsante", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "VIRTUALPAY_URL_TEST", _
									0, _
									"Url di invio del pagamento (Via POST) per la modalità di test", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									"https://www.servizipos.it/GT_EgipsyWeb/CheckOutEGIPSy.jsp", _
									"https://www.servizipos.it/GT_EgipsyWeb/CheckOutEGIPSy.jsp", _
									"https://www.servizipos.it/GT_EgipsyWeb/CheckOutEGIPSy.jsp", _
									"https://www.servizipos.it/GT_EgipsyWeb/CheckOutEGIPSy.jsp", _
									"https://www.servizipos.it/GT_EgipsyWeb/CheckOutEGIPSy.jsp")
		CALL AddParametroSito(conn, "VIRTUALPAY_MERCHANT_ID_TEST", _
									0, _
									"Identificatore del negozio virtuale per la modalità di test", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									"405206400002", "405206400002", "405206400002", "405206400002", "405206400002")
		CALL AddParametroSito(conn, "VIRTUALPAY_ABI_TEST", _
									0, _
									"ABI della banca alla quale deve essere indirizzata la transazione per la modalità di test", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									"03599", "03599", "03599", "03599", "03599")
		CALL AddParametroSito(conn, "VIRTUALPAY_MAC_TEST", _
									0, _
									"Message Authentication Code che il negozio virtuale può utilizzare per certificare l’integrità dei dati ordine per la modalità di test", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									BANCA_VIRTUALPAY, _
									"E5D49168DEAA74078B22524B02360B6B", _
									"E5D49168DEAA74078B22524B02360B6B", _
									"E5D49168DEAA74078B22524B02360B6B", _
									"E5D49168DEAA74078B22524B02360B6B", _
									"E5D49168DEAA74078B22524B02360B6B")
	end if
end function
'*******************************************************************************************

%>