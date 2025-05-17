<%
'......................................................................................................................................
'definizione oggetto che contiene le carte di credito gestite dal sistema
'......................................................................................................................................
dim CreditCardTypes
set CreditCardTypes = Server.CreateObject("Scripting.Dictionary")
CreditCardTypes.CompareMode = vbTextCompare
if not InIdList(Application("DISABLED_CC_TYPES"), "A") then     CreditCardTypes.add "A", "American Express"
if not InIdList(Application("DISABLED_CC_TYPES"), "B") then     CreditCardTypes.add "B", "Carte Blanche"
if not InIdList(Application("DISABLED_CC_TYPES"), "C") then     CreditCardTypes.add "C", "Diners Club"
if not InIdList(Application("DISABLED_CC_TYPES"), "M") then     CreditCardTypes.add "M", "Mastercard"
if not InIdList(Application("DISABLED_CC_TYPES"), "U") then     CreditCardTypes.add "U", "Eurocard"
if not InIdList(Application("DISABLED_CC_TYPES"), "V") then     CreditCardTypes.add "V", "Visa"
if not InIdList(Application("DISABLED_CC_TYPES"), "D") then     CreditCardTypes.add "D", "Discover"
if not InIdList(Application("DISABLED_CC_TYPES"), "E") then     CreditCardTypes.add "E", "EnRoute"
if not InIdList(Application("DISABLED_CC_TYPES"), "J") then     CreditCardTypes.add "J", "JCB"


'......................................................................................................................................
'nuova routine che disegna la parte di form per la richiesta della carta di credito
'DEVE ESSERE INTEGRATA IN UN FORM DI TIPO "contattaci" CON STILI GENERALI
'		oCnt			Oggetto request o recordset contente i valori da caricare sul form
'		TypeField			Nome del campo del tipo di carta di credito
'		OwnerField			Nome del campo del titolare della carta di credito
'		NumberField			Nome del campo del numero della carta di credito
'		MonthField			Nome del campo del mese di scadenza della carta di credito
'		YearField 			Nome del campo dell'anno di scadenza della carta di credito
'......................................................................................................................................
function Form_CreditCard(oCnt, disabled, TypeField, OwnerField, NumberField, CvcField, MonthField, YearField)
	dim value
	%>
	<tr>
		<td class="label"><%= ChooseByLanguage("Tipo", "Type", "Typ", "Type", "Tipo") %></td>
		<td colspan="3" class="input">
			<% CALL Form_DropDownDictionaryField(disabled, "tft_" + TypeField, CreditCardTypes, oCnt(TypeField), "", true)
			CALL Form_CampoObbligatorio(disabled, NULL, NULL) %>
		</td>
	</tr>
	<tr>
		<td class="label"><%= ChooseByLanguage("Numero", "Number", "Nummer", "Num&eacute;ro ", "N&uacute;mero") %></td>
		<td colspan="3" class="input">
			<% 
			if (cString(oCnt("pre_data")) <> "" AND not Request.ServerVariables("REQUEST_METHOD")="POST") OR IsAmministrazione() then 
				value = DecryptCreditCard(oCnt("pre_data"), oCnt("pre_str_id"), oCnt(NumberField))
			else
				value = oCnt(NumberField)
			end if
			%>
			<% CALL Form_TextField(disabled, "tft_" + NumberField, _
						IIF(disabled<>"", MaskValue(oCnt(NumberField), false), value), _
						"number", 25, 20)
			CALL Form_CampoObbligatorio(disabled, NULL, NULL) %>&nbsp;
		</td>
	</tr>
	<tr>
		<td class="label" <% if disabled = "" then %>rowspan="2"<% end if %>>CVC (CVV - AM)</td>
		<td colspan="3" class="input">
			<% 
			if (cString(oCnt("pre_data")) <> "" AND not Request.ServerVariables("REQUEST_METHOD")="POST") OR IsAmministrazione() then 
				value = DecryptCreditCard(oCnt("pre_data"), oCnt("pre_str_id"), oCnt(CvcField))
			else
				value = oCnt(CvcField)
			end if
			%>
			<% CALL Form_TextField(disabled, "tft_" + CvcField, _
						IIF(disabled<>"", MaskValue(oCnt(CvcField), true), value), _
						"number", 5, 4)
			CALL Form_CampoObbligatorio(disabled, NULL, NULL) %>&nbsp;
		</td>
	</tr>
	<% if disabled = "" then %>
		<tr>
			<td colspan="3" class="note">
				<% CALL Form_CreditCard_CvcMessage() %>&nbsp;
			</td>
		</tr>
	<% end if %>
	<tr>
		<td class="label"><%= ChooseByLanguage("Data di Scadenza", "Expiry date", "Ablaufdatum", "Date d'expiration", "Fecha de vencimiento") %></td>
		<td colspan="3" class="input">
			<% if disabled = "" then
				CALL DropDownInterval(1, 12, "tft_" + MonthField, IIF(cInteger(oCnt(MonthField))=0, Month(Date), oCnt(MonthField)), true, "", Session("LINGUA"))
				CALL DropDownInterval(Year(Date), (Year(Date) + 10), "tft_" + YearField, IIF(cInteger(oCnt(YearField))=0, Year(Date), oCnt(YearField)), true, "", Session("LINGUA"))
				CALL Form_CampoObbligatorio(disabled, NULL, NULL) %>
			<% else %>
				<% CALL Form_TextField(disabled, "tft_" + MonthField, oCnt(MonthField), "number", 5, 4) %> - 
				<% CALL Form_TextField(disabled, "tft_" + YearField, oCnt(YearField), "number", 5, 4)
			CALL Form_CampoObbligatorio(disabled, NULL, NULL) %>
			<% end if %>&nbsp;
		</td>
	</tr>
	<tr>
		<td class="label"><%= ChooseByLanguage("Nome del titolare", "Holder's name", "Name des Karteninhabers", "Nom du titulaire", "Nombre del titular") %></td>
		<td colspan="3" class="input">
			<% CALL Form_TextField(disabled, "tft_" + OwnerField, oCnt(OwnerField), "text", 66, 255)
			CALL Form_CampoObbligatorio(disabled, NULL, NULL) %>&nbsp;
		</td>
	</tr>
<% end function


'......................................................................................................................................
'routine che disegna la parte di form per la richiesta della carta di credito
'		ObjValues			Oggetto request o recordset contente i valori da caricare sul form
'		TypeField			Nome del campo del tipo di carta di credito
'		OwnerField			Nome del campo del titolare della carta di credito
'		NumberField			Nome del campo del numero della carta di credito
'		MonthField			Nome del campo del mese di scadenza della carta di credito
'		YearField 			Nome del campo dell'anno di scadenza della carta di credito
'......................................................................................................................................
function CreditCardForm(ObjValues, TypeField, OwnerField, NumberField, CvcField, MonthField, YearField)
	dim i, val, PrefixValues
	if instr(1, TypeName(ObjValues), "recordset", vbTextCompare) OR _
	   instr(1, TypeName(ObjValues), "indirizzario", vbTextCompare) then
		PrefixValues = ""
	else
		PrefixValues = "tft_"
	end if %>
	<table cellspacing="0" cellpadding="2" class="contact_table" id="CarteCredito">
		<tr>
			<td class="contact_label"><%= ChooseByLanguage("Tipo", "Type", "Typ", "Type", "Tipo") %>:</td>
			<td class="contact_input">
				<% CALL DropDownDictionary(CreditCardTypes, "tft_" & TypeField, ObjValues(PrefixValues & TypeField), true, " class=""contact_select"" ", Session("LINGUA")) %>
			</td>
			<td class="contact_mandatory">(*)</td>
		</tr>
		<tr>
			<td class="contact_label"><%= ChooseByLanguage("Numero", "Number", "Nummer", "Num&eacute;ro ", "N&uacute;mero") %>:</td>
			<td class="contact_input">
				<input type="text" class="contact_text" name="tft_<%= NumberField %>" value="<%= MaskValue(ObjValues(PrefixValues & NumberField), false) %>" maxlength="20">
			</td>
			<td class="contact_mandatory">(*)</td>
		</tr>
		<tr>
			<td class="contact_label" rowspan="2">CVC (CVV - AM):</td>
			<td class="contact_input">	
				<input type="text" class="contact_text" name="tft_<%= CvcField %>" value="<%= MaskValue(ObjValues(PrefixValues & CvcField), true) %>" maxlength="4" style="width:12%;">
			</td>
			<td class="contact_mandatory">(*)</td>
		</tr>
		<tr>
			<td class="contact_input">
				<% CALL Form_CreditCard_CvcMessage() %>
			</td>
			<td class="contact_mandatory">&nbsp;</td>
		</tr>
		<tr>
			<td class="contact_label"><%= ChooseByLanguage("Data di Scadenza", "Expiry date", "Ablaufdatum", "Date d'expiration", "Fecha de vencimiento") %>:</td>
			<td class="contact_input">
				<%
				CALL DropDownInterval(1, 12, "tft_" + MonthField, IIF(cInteger(ObjValues(PrefixValues & YearField))=0, Month(Date), ObjValues(PrefixValues & YearField)), true, " class=""contact_select"" ", Session("LINGUA"))
				CALL DropDownInterval(Year(Date), (Year(Date) + 10), "tft_" + YearField, IIF(cInteger(ObjValues(PrefixValues & YearField))=0, Year(Date), ObjValues(PrefixValues & YearField)), true, " class=""contact_select"" ", Session("LINGUA"))
				%>
			</td>
			<td class="contact_mandatory">(*)</td>
		</tr>
		<tr>
			<td class="contact_label"><%= ChooseByLanguage("Nome del titolare", "Holder's name", "Name des Karteninhabers", "Nom du titulaire", "Nombre del titular") %>:</td>
			<td class="contact_input">
				<input type="text" class="contact_text" name="tft_<%= OwnerField %>" value="<%= ObjValues(PrefixValues & OwnerField) %>" maxlength="255">
			</td>
			<td class="contact_mandatory">(*)</td>
		</tr>
	</table>
<%end function



'......................................................................................................................................
'routine che disegna la parte di form per la richiesta della carta di credito
'		ObjValues			Oggetto request o recordset contente i valori da caricare sul form
'		TypeField			Nome del campo del tipo di carta di credito
'		OwnerField			Nome del campo del titolare della carta di credito
'		NumberField			Nome del campo del numero della carta di credito
'		MonthField			Nome del campo del mese di scadenza della carta di credito
'		YearField 			Nome del campo dell'anno di scadenza della carta di credito
'......................................................................................................................................
function CreditCardView(ObjValues, TypeField, OwnerField, NumberField, CvcField, MonthField, YearField)
	dim val, PrefixValues
	if instr(1, TypeName(ObjValues), "recordset", vbTextCompare) OR _
	   instr(1, TypeName(ObjValues), "indirizzario", vbTextCompare) then
		PrefixValues = ""
	else
		PrefixValues = "tft_"
	end if%>
	<table cellspacing="0" cellpadding="2" class="contact_table" id="CarteCredito">
		<tr>
			<td class="contact_label"><%= ChooseByLanguage("Tipo", "Type", "Typ", "Type", "Tipo") %>:</td>
			<td class="contact_value"><span class="contact_value">&nbsp;<%= CreditCardTypes(ObjValues(PrefixValues & TypeField))%></span></td>
		</tr>
		<tr>
			<td class="contact_label"><%= ChooseByLanguage("Numero", "Number", "Nummer", "Num&eacute;ro ", "N&uacute;mero") %>:</td>
			<td class="contact_value">
				<span class="contact_value">
					<%= MaskValue(ObjValues(PrefixValues & NumberField), false) %>
				</span>
			</td>
		</tr>
		<tr>
			<td class="contact_label">CVC:</td>
			<td class="contact_value">	
				<span class="contact_value">
					<%= MaskValue(ObjValues(PrefixValues & CvcField), true) %>
				</span>
			</td>
		</tr>
		<tr>
			<td class="contact_label"><%= ChooseByLanguage("Data di Scadenza", "Expiry date", "Ablaufdatum", "Date d'expiration", "Fecha de vencimiento") %>:</td>
			<td class="contact_value">
				<span class="contact_value">
					<%= ObjValues(PrefixValues & MonthField) %>
					<%= IIF(cInteger(ObjValues(PrefixValues & MonthField))<>0, " / " , "") %>
					<%= ObjValues(PrefixValues & YearField) %>
				</span>
			</td>
		</tr>
		<tr>
			<td class="contact_label"><%= ChooseByLanguage("Nome del titolare", "Holder's name", "Name des Karteninhabers", "Nom du titulaire", "Nombre del titular") %>:</td>
			<td class="contact_value"><span class="contact_value"><%= ObjValues(PrefixValues & OwnerField) %></span></td>
		</tr>
	</table>
<%end function


'......................................................................................................................................
'routine che esegue il controllo dei dati della carta di credito
'		Tipo			Tipo di carta di credito
'		Titolare		Titolare della carta di credito
'		Numero			Numero della carta di credito
'		Mese			Mese di scadenza della carta di credito
'		Anno 			Anno di scadenza della carta di credito
'......................................................................................................................................
function CreditCardCheck(Tipo, Titolare, Numero, CVC, Mese, Anno)
	dim check
	if Titolare ="" then
		Session("Errore") = ChooseByLanguage("Nome del titolare della carta di credito non inserito", _
											 "Missing credit card holder's name", _
											 "Unzul&auml;ssiger Name des Karteninhabers", _
											 "Nom du propri&eacute;taire de la carte inadmissible", _
											 "Nombre del titular de la tarjeta de cr&eacute;dito inv&aacute;lido")
		CreditCardCheck = false
		Exit function
	end if
	
	if CVC = "" OR not isNumeric(CVC) then
		Session("Errore") = ChooseByLanguage("Codice CVC della carta di credito non inserito", _
											 "Missing credit card's CVC code", _
											 "Unzul&auml;ssiger Code CVC der Kreditkarte", _
											 "Code CVC de votre Carte de cr&eacute;dit inadmissible", _
											 "C&oacute;digo CVC de tu tarjeta de cr&eacute;dito inv&aacute;lido")
		CreditCardCheck = false
		Exit function
	end if
	
	check = checkcc(Mese, Anno, Numero, Tipo)
	Select case check
		case 0
			CreditCardCheck = true
		case 1
			Session("Errore") = ChooseByLanguage("Seleziona il tipo di carta di credito", _
												 "Select credit card's type", _
												 "Bitte w&auml;hlen Sie einen Kreditkartentyp aus", _
												 "S&eacute;lectionner le type de Carte de cr&eacute;dit", _
												 "Selecciona el tipo de la tarjeta de cr&eacute;dito")
			CreditCardCheck = false
		case 9
			Session("Errore") = ChooseByLanguage("Data di scadenza della carta di creditonon valida", _
												 "Invalid credit card's expiry date", _
												 "Unzul&auml;ssiges Kreditkarte Ablaufdatum", _
												 "Date d'expiration de le Carte de cr&eacute;dit inadmissible", _
												 "Fecha de vencimiento de la tarjeta de cr&eacute;dito inv&aacute;lida")
			CreditCardCheck = false
		case else
			Session("Errore") = ChooseByLanguage("Numero della carta di credito non valido", _
												 "Invalid credit card's number", _
												 "Unzul&auml;ssiges Kreditkartennummer ", _
												 "Num&eacute;ro de Carte de cr&eacute;dit inadmissible", _
												 "N&uacute;mero de la tarjeta de cr&eacute;dito inv&aacute;lido")
			CreditCardCheck = false
	end select
end function


'......................................................................................................................................
'funzione che maschera il valore se non &egrave;  attivo HTTPS
'value			valore da mascherare con X al posto di ogni carattere
'MaskAll		indica se il valore deve essere mascherato tutto o devono essere esclusi i primi tre caratteri
'......................................................................................................................................
function MaskValue(value, MaskAll)
	'......................................................................................................................................
	'......................................................................................................................................
	'Commentata il 16/05/2011 per rendere sicuro il codice e non mostrare più un chiaro la carta di credito ed il cvc
	'......................................................................................................................................
	'......................................................................................................................................

	'value = cString(value)
	'if instr(1, request.serverVariables("HTTPS"), "off", vbTextCompare) AND value <> "" then
	'	if MaskAll then
	'		MaskValue = string(len(value), "X")
	'	else
	'		MaskValue = left(value, 3)  & string(len(value)-3, "X")
	'	end if
	'else
	'	MaskValue = value
	'end if
	dim i
	i = cIntero(len(value))
	if not MaskAll then
		if i > 16 then i = 16
	else
		if i > 3 then i = 3
	end if	
	MaskValue = string(i, "X")
end function


'......................................................................................................................................
'procedura che scrive il messaggio di descrizione del cvc
'......................................................................................................................................
sub Form_CreditCard_CvcMessage()	
	Select case lcase(Session("LINGUA"))
		case LINGUA_ITALIANO 
			if not InIdList(Application("DISABLED_CC_TYPES"), "A") then %>
				<strong>Per le American Express:</strong><br />
				Il numero di verifica (CVV, Card Verification Value) &egrave; composto da 4 cifre e si trova nella parte anteriore della carta di credito, 
				al di sopra del numero.<br />
				<strong>Per le altre carte di credito:</strong><br />
			<% end if %>
			Il codice di verifica (CVC, Card Verification Code) &egrave; composto da 3 cifre, le ultime 3 del numero che si trova sul retro 
			della carta di credito, nello spazio destinato alla firma.
		<% case LINGUA_FRANCESE
			if not InIdList(Application("DISABLED_CC_TYPES"), "A") then %>
				<strong>American Express:</strong><br />
				Les quatre chiffres du code CVC sont sur le verso de votre carte American Express, en haut &agrave; droite.
				<strong>Autres cartes :</strong><br />
			<% end if %>
			Les trois chiffres du code CVC suivent le num&eacute;ro de votre carte de cr&eacute;dit, sur le verso de celle-ci.
		<% case LINGUA_TEDESCO
			if not InIdList(Application("DISABLED_CC_TYPES"), "A") then %>
				<strong>F&uuml;r American Express cards:</strong><br />
				Auf der R&uuml;ckseite Ihre Kreditkarte sehen Sie drei Ziffern, die nicht in der Kreditkartennummer enthalten sind.
				Der Sicherheitscode besteht aus den letzten drei Ziffern.
				<strong>Andere Kreditkartenarten:</strong><br />
			<% end if %>
			Rechts &uuml;ber der Kreditkartennummer auf der Vorderseite Ihrer American Express Karte sind vier Ziffern zu sehen..
			Diese vier Ziffern sind Ihr Sicherheitscode.
		<% case LINGUA_SPAGNOLO 
			if not InIdList(Application("DISABLED_CC_TYPES"), "A") then %>
				<strong>Para American Express Card:</strong><br />
				En la parte frontal de tu American Express, hay cuatro d&iacute;gitos situados en la parte superior derecha del n&uacute;mero de la tarjeta.
				Estos 4 d&iacute;gitos son el CVC.
				<strong>Otros tipos de la tarjeta de cr&eacute;dito:</strong><br />
			<% end if %>
			En la parte trasera de tu tarjeta de cr&eacute;dito, hay tres d&iacute;gitos que no forman parte del n&uacute;mero de la tarjeta.
			El CVC son los tres &uacute;ltimos d&iacute;gitos.
		<% case else 'lingua inglese
			if not InIdList(Application("DISABLED_CC_TYPES"), "A") then %>
				<strong>For American Express cards:</strong><br />
				On the front of your American Express card, above and to the right of your credit card number, you should see four digits.
				These four digits are your CVV.<br />
				<strong>Other credit card types:</strong><br />
			<% end if %>
			On the back of your credit card, you should see some digits that are not part of your credit card number.
			The CVC is the last three digits. 
	<% end select
end sub




'..........................................................................................................
'restituisce il numero della carta di credito criptata
'	pre_data:		data prenotazione (corrispondente alla colonna pre_data su vtb_prenotazioni)
'	str_id:			id struttura (corrispondente alla colonna pre_str_id su vtb_prenotazioni)
'	numero_carta:	numero carta di credito da criptare
'..........................................................................................................
Function EncryptCreditCard(pre_data, str_id, numero_carta)
	dim Cripto
	set Cripto = new CryptographyManager
	if cString(numero_carta) <> "" AND cString(str_id) <> "" AND cString(pre_data) <> "" then
		EncryptCreditCard = Cripto.aes_of_string(UCASE(numero_carta), UCASE(Year(pre_data)&Month(pre_data)&Day(pre_data)&str_id))
	else
		EncryptCreditCard = ""
	end if
	set Cripto = nothing
end function


'..........................................................................................................
'restituisce il numero della carta a partire dalla stringa criptata
'	pre_data:		data prenotazione (corrispondente alla colonna pre_data su vtb_prenotazioni)
'	str_id:			id struttura (corrispondente alla colonna pre_str_id su vtb_prenotazioni)
'	crypted:		numero carta di credito criptata
'..........................................................................................................
Function DecryptCreditCard(pre_data, str_id, crypted)
	dim Cripto
	set Cripto = new CryptographyManager
	if cString(crypted) <> "" AND cString(str_id) <> "" AND cString(pre_data) <> "" then
		DecryptCreditCard = Cripto.string_from_aes(crypted, UCASE(Year(pre_data)&Month(pre_data)&Day(pre_data)&str_id))
	else
		DecryptCreditCard = ""
	end if
	set Cripto = nothing
end function



'......................................................................................................................................
'......................................................................................................................................
'......................................................................................................................................
'......................................................................................................................................
'......................................................................................................................................
'......................................................................................................................................
' Credit Card check routine for ASP
' (c) 1998 by Click Online
' You may use these functions only if this header is not removed
' http://www.click-online.de
' info@click-online.de

' Developers may use the following numbers as dummy data:
' Visa:					430-00000-00000
' American Express:		372-00000-00000
' Mastercard:			521-00000-00000
' Discover:				620-00000-00000


'results of check
'0		Card is ok!
'1		Wrong card type
'2		Wrong length
'3		Wrong length and card type
'4		Wrong checksum
'5		Wrong checksum and card type
'6		Wrong checksum and length
'7		Wrong checksum, length and card type
'8		unknown cardtype
'9		data errata



function trimtodigits(tstring)
	dim s, ts, x, ch
'removes all chars except of 0-9
  s="" 
  ts=tstring
  for x=1 to len(ts)
    ch=mid(ts,x,1)
    if asc(ch)>=48 and asc(ch)<=57 then
      s=s & ch
    end if
  next
  trimtodigits=s
end function

function checkcc(mese, anno, ccnumber,cctype)
	dim ctype, cclength, ccprefix, prefixes, lengths, number, prefixvalid, lengthvalid
	dim prefix, length, result, qsum, x, ch, sum
  'checks credit card number for checksum,length and type
  'ccnumber= credit card number (all useless characters are
  '	being removed before check)
  '
  'cctype:
  '       "V" VISA
  '       "M" Mastercard
  '       "U" Eurocard
  '       "A" American Express
  '       "B" Carte Blanche
  '       "C" Diners Club
  '       "D" Discover
  '       "E" enRoute
  '       "J" JCB
  'returns:  checkcc=0 (Bit0)  : card valid
  '          checkcc=1 (Bit1) : wrong type
  '          checkcc=2 (Bit2) : wrong length
  '          checkcc=4 (Bit3) : wrong checksum (MOD10-Test)
  '          checkcc=8 (Bit4) : cardtype unknown

'controllo data
if mese = "" then
	mese = 13
end if
if anno = "" then
	anno = Year(Date())+1
end if
dim annoA
annoA = Year(Date())
if CInt(mese) <= Month(Date()) AND CInt(left(annoA, len(annoA)-len(anno)) & anno) <= annoA then
  	CheckCC = 9
else
  
  if ccnumber="0000000" then
  	checkcc = 0
	exit function
  end if
    
  ctype=ucase(cctype)
  select case ctype
    case "V"
      cclength="13;16"
      ccprefix="4"
    case "M"
      cclength="16"
      ccprefix="51;52;53;54;55"
    case "U"
      cclength="16"
      ccprefix="51;52;53;54;55"
    case "A"
      cclength="15"
      ccprefix="34;37"
    case "C"
      cclength="14"
      ccprefix="300;301;302;303;304;305;36;38"
  	case "B"
      cclength="14"
      ccprefix="300;301;302;303;304;305;36;38"
    case "D"
      cclength="16"
      ccprefix="6011;620"
    case "E"
      cclength="15"
      ccprefix="2014;2149"
    case "J"
      cclength="15;16"
      ccprefix="3;2131;1800"
    case else
      cclength=""
      ccprefix=""
  end select
  prefixes=split(ccprefix,";",-1)
  lengths=split(cclength,";",-1)
  number=trimtodigits(ccnumber)
  prefixvalid=false
  lengthvalid=false
  for each prefix in prefixes
    if instr(number,prefix)=1 then
      prefixvalid=true
    end if
  next  
  for each length in lengths
    if cstr(len(number))=length then
      lengthvalid=true
    end if
  next
  result=0
  if not prefixvalid then
    result=result+1
  end if  
  if not lengthvalid then
    result=result+2
  end if  
  qsum=0
  for x=1 to len(number)
    ch=mid(number,len(number)-x+1,1)
    if x mod 2=0 then
      sum=2*cint(ch)
      qsum=qsum+(sum mod 10)
      if sum>9 then 
        qsum=qsum+1
      end if
    else
      qsum=qsum+cint(ch)
    end if
  next
  if qsum mod 10<>0 then
    result=result+4
  end if
  if cclength="" then
    result=result+8
  end if
  checkcc=result

end if			'fine controllo date
end function
'......................................................................................................................................
'......................................................................................................................................
'......................................................................................................................................
'......................................................................................................................................
'......................................................................................................................................
%>
