<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per l'applicativo Paypal 2.0 
'...........................................................................................
'...........................................................................................

'*******************************************************************************************
'Installazione Applicazione Paypal 2.0
'...........................................................................................
' Luca 22/06/2015
'...........................................................................................
function Aggiornamento__Paypal__1(conn)
	Aggiornamento__Paypal__1 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__Paypal__1(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & PAYPAL_2_0)) <> "" then
		CALL AddParametroSito(conn, "PAYPAL_SANDBOX_ATTIVO", _
									0, _
									"Indica se paypal è attivo come sandbox o meno", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									PAYPAL_2_0, _
									1, null, null, null, null)
		CALL AddParametroSito(conn, "PAYPAL_SUBMIT_URL", _
									0, _
									"Url per il pagamento ed i controlli con ipn e pdt", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									PAYPAL_2_0, _
									"https://www.paypal.com/cgi-bin/webscr", _
									"https://www.paypal.com/cgi-bin/webscr", _
									"https://www.paypal.com/cgi-bin/webscr", _
									"https://www.paypal.com/cgi-bin/webscr", _
									"https://www.paypal.com/cgi-bin/webscr")
		CALL AddParametroSito(conn, "PAYPAL_EMAIL_SELLER", _
									0, _
									"Indirizzo email del venditore", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									PAYPAL_2_0, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "PAYPAL_PDT_TOKEN", _
									0, _
									"Token per il Payment Data Transfer", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									PAYPAL_2_0, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "PAYPAL_SUBMIT_URL_SANDBOX", _
									0, _
									"Url Sandbox per il pagamento ed i controlli con ipn e pdt", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									PAYPAL_2_0, _
									"https://www.sandbox.paypal.com/cgi-bin/webscr", _
									"https://www.sandbox.paypal.com/cgi-bin/webscr", _
									"https://www.sandbox.paypal.com/cgi-bin/webscr", _
									"https://www.sandbox.paypal.com/cgi-bin/webscr", _
									"https://www.sandbox.paypal.com/cgi-bin/webscr")
		CALL AddParametroSito(conn, "PAYPAL_EMAIL_SELLER_SANDBOX", _
									0, _
									"Indirizzo email Sandbox del venditore", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									PAYPAL_2_0, _
									"suppor_1216634323_biz@next-aim.com", _
									"suppor_1216634323_biz@next-aim.com", _
									"suppor_1216634323_biz@next-aim.com", _
									"suppor_1216634323_biz@next-aim.com", _
									"suppor_1216634323_biz@next-aim.com")
		CALL AddParametroSito(conn, "PAYPAL_PDT_TOKEN_SANBOX", _
									0, _
									"Token Sandbox per il Payment Data Transfer", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									PAYPAL_2_0, _
									"ZNy9FsF4wkWIrD8PYuXCNfW8WE_j_jxuMB5xgvIrR35Ow7mp-tatLCMZj-q", _
									"ZNy9FsF4wkWIrD8PYuXCNfW8WE_j_jxuMB5xgvIrR35Ow7mp-tatLCMZj-q", _
									"ZNy9FsF4wkWIrD8PYuXCNfW8WE_j_jxuMB5xgvIrR35Ow7mp-tatLCMZj-q", _
									"ZNy9FsF4wkWIrD8PYuXCNfW8WE_j_jxuMB5xgvIrR35Ow7mp-tatLCMZj-q", _
									"ZNy9FsF4wkWIrD8PYuXCNfW8WE_j_jxuMB5xgvIrR35Ow7mp-tatLCMZj-q")
		CALL AddParametroSito(conn, "PAYPAL_ORDER_STATE_PAYED_ID", _
									0, _
									"Id dello stato d'ordine pagato", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									PAYPAL_2_0, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "PAYPAL_ALERT_CODE", _
									0, _
									"Codice dell'alert per paypal", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									PAYPAL_2_0, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "PAYPAL_RETURN_PAGE", _
									0, _
									"Pagina di ritorno del sistema alla conferma del pagamento", _
									"", _
									adGUID, _
									0, _
									"", _
									1, _
									1, _
									PAYPAL_2_0, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "PAYPAL_CANCEL_RETURN_PAGE", _
									0, _
									"Pagina di ritorno nel caso in cui il cliente annulli il processo di pagamento", _
									"", _
									adGUID, _
									0, _
									"", _
									1, _
									1, _
									PAYPAL_2_0, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************

%>