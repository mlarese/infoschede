<%

'.................................................................................................
'.................................................................................................
'FUNZIONI COMUNI PER GESTIONE DEI FORM
'.................................................................................................
'.................................................................................................

function GetPrivacyText()
	if request.querystring("HTML_FOR_EMAIL")<>"" then
		'testo della privacy allegato automaticamente alle email
		GetPrivacyText = ChooseByLanguage("<strong>Ai sensi dell' articolo 13 Legge 196/03 si comunica che:</strong><br>" + _
										  "le informazioni contenute in questo messaggio sono riservate ed a uso esclusivo del destinatario. " + _
										  "Qualora il messaggio in parola Le fosse pervenuto per errore, la preghiamo di eliminarlo senza copiarlo e di " + _
										  "non inoltrarlo a terzi, dandocene gentilmente comunicazione. Grazie.", _
										  "<strong>In accordance with article 13 of Italian law 196/03 we inform you that:</strong><br>" + _
										  "This message may contain confidential and/or privileged information. If you are not the addressee or " + _
										  "authorized to receive this for the addressee, you must not use, copy, disclose or take any action based " +_
										  "on this message or any information herein. If you have received this message in error, please advise the " + _
										  "sender immediately by reply e-mail and delete this message. Thank you for your cooperation.", _
										  "", _
										  "<strong>Dans le respect de l’article 13 de la Loi 196/03 nous vous informons que:</strong><br />" + _
										  "les informations fournies ont comme unique but de satisfaire le service demand&eacute;.<br />Ces informations pourront &ecirc;tre communiqu&eacute;es seulement au personnel pr&eacute;pos&eacute; &agrave; notre activit&eacute; et diffus&eacute; exclusivement dans le domaine des finalit&eacute;s du service rendu. " + _
										  "De telles informations pourront &ecirc;tre trait&eacute;es &eacute;lectroniquement en conformit&eacute; avec les lois en vigueur.", _
										  "")
	else
		GetPrivacyText = ChooseByLanguage("<strong>Nel rispetto dell' articolo 13 Legge 196/03 si comunica che:</strong><br>" + _
										  "I dati qui raccolti hanno l'unico scopo di poter soddisfare il servizio richiesto. " + _
										  "Tali informazioni potranno essere comunicate solo al personale preposto a tale attivit&agrave; " + _
										  "e diffusi esclusivamente nell'ambito delle finalit&agrave; del servizio reso. " + _
										  "Tali informazioni potranno essere trattate anche elettronicamente in conformit&agrave; con le leggi vigenti.", _
										  "<strong>In accordance with article 13 of Italian law 196/03 we inform you that:</strong><br>" + _
										  "The data collected here are used solely to meet the needs of the requested service. " + _
										  "This information may be used only by the personnel responsible for this activity and only " + _
										  "divulged within the scope and purposes of the service provided.", _
										  "<strong>Gem&auml;&szlig; Artikel 13 des italienischen Datenschutzgesetzes 196/03 wird auf folgendes hingewiesen:</strong><br> " + _
										  "Die gesammelten Daten dienen unseren Mitarbeiten ausschliesslich zur Bearbeitung der angefragten Dienstleistung.", _
										  "<strong>Dans le respect de l'article 13 de la Loi 196/03 nous vous informons que:</strong><br />" + _
										  "les informations fournies ont comme unique but de satisfaire le service demand&eacute;.<br />Ces informations pourront &ecirc;tre communiqu&eacute;es seulement au personnel pr&eacute;pos&eacute; &agrave; notre activit&eacute; et diffus&eacute; exclusivement dans le domaine des finalit&eacute;s du service rendu. " + _
										  "De telles informations pourront &ecirc;tre trait&eacute;es &eacute;lectroniquement en conformit&eacute; avec les lois en vigueur." , _
										  "<strong>Con arreglo al art&iacute;culo 13 de la ley italiana 196/03 se comunica que:</strong><br> " + _
										  "El &uacute;nico objetivo de los datos aqu&iacute; mencionados es el de satisfacer el servicio requerido. " + _
										  "Dichas informaciones podr&aacute;n ser utilizadas s&oacute;lo por el personal encargado de dichas actividades " + _
										  "y difundidos exclusivamente en el &aacute;mbito y con las finalidades del servicio ofrecido.")
	end if
end function


function Form_TextField( disabled, fname, value, classname, size, maxlength )
	if disabled="" then %>
		<input type="text" class="<%= classname %>" name="<%= fname %>" size="<%= size %>" maxlength="<%= maxlength %>" value="<%= value %>">
	<% else 
		if value="" then value ="&nbsp;"%>
		<span class="<%= classname %>" id="<%= fname %>"><%= value %></span>	
	<% end if
end function


function Form_TextAreaField( disabled, fname, value, classname, rows )
	if disabled="" then %>
		<textarea class="<%= classname %>" name="<%= fname %>" rows="<%= rows %>"><%= value %></textarea>
	<% else %>
		<span class="<%= classname %>" id="<%= fname %>"><%= IIF(value="", "&nbsp;", TextEncode(value)) %></span>	
	<% end if
end function


function Form_CheckboxField( disabled, fname, value, classname, status)
	if disabled="" then%>
		<input type="checkbox" class="<%= classname %>" name="<%= fname %>" value="<%= value %>" <%= status %>>
	<% else %>
		<input disabled type="checkbox" class="<%= classname %>" name="<%= fname %>" value="<%= value %>" <%= status %>>	
	<% end if
end function

function Form_DropDownDictionaryField(disabled, fname, list, value, classname, mandatory)
	CALL DropDownDictionary(list, fname, value, mandatory, IIF(disabled <> "", " disabled", "") & IIF(classname<>"", " class=""" & classname & """ ", ""), Session("LINGUA"))
end function


function Form_DataPickerField( disabled, fname, value, classname)
	CALL Form_DataPickerFieldEX( disabled, fname, value, classname, "")
end function 

function Form_DataPickerFieldEX( disabled, fname, value, classname, fname2)
	if disabled="" then
		CALL WriteDataPicker_Input_Ex("form1", fname, value, "../../stili.css", "", false, false, Session("LINGUA"), fname2)
	else 
		if value="" then value ="&nbsp;"%>
		<span class="<%= classname %>" id="<%= fname %>"><%= value %></span>	
	<% end if
end function


function Form_EMailField( disabled, fname, value, classname, size, maxlength )
	if disabled="" then %>
		<input type="text" class="<%= classname %>" name="<%= fname %>" size="<%= size %>" maxlength="<%= maxlength %>" value="<%= value %>">
	<% else
		if value="" then value ="&nbsp;"%>
		<span class="<%= classname %>" id="<%= fname %>"><a class="<%= classname %>" href="mailto:<%= value %>"><%= value %></a></span>	
	<% end if
end function


sub Form_InitConfig(Config)
	Select Case Config.Lingua 
		Case LINGUA_INGLESE
			Config.AddDefault "lbl_Error", "Data not saved:"
			Config.AddDefault "lbl_save", "send"
			Config.AddDefault "lbl_undo", "reset"
			Config.AddDefault "lbl_messaggio", "message"
		Case LINGUA_SPAGNOLO
			Config.AddDefault "lbl_Error", "Datos no ahorrados:"
			Config.AddDefault "lbl_save", "envíe"
			Config.AddDefault "lbl_undo", "reajuste"
			Config.AddDefault "lbl_messaggio", "mensaje"
		Case LINGUA_TEDESCO
			Config.AddDefault "lbl_Error", "Daten nicht gespeichert:"
			Config.AddDefault "lbl_save", "senden Sie"
			Config.AddDefault "lbl_undo", "Zurückstellen"
			Config.AddDefault "lbl_messaggio", "Anzeige"
		Case LINGUA_FRANCESE
			Config.AddDefault "lbl_Error", "Donn&eacute;es non sauv&eacute;es:"
			Config.AddDefault "lbl_save", "envoyez"
			Config.AddDefault "lbl_undo", "remise"
			Config.AddDefault "lbl_messaggio", "message"
		Case else
			Config.AddDefault "lbl_Error", "Registrazione non eseguita:"
			Config.AddDefault "lbl_save", "invia"
			Config.AddDefault "lbl_undo", "annulla"
			Config.AddDefault "lbl_messaggio", "messaggio"
	end select
end sub


sub Form_CampoObbligatorio(disabled, field, MandatoryFields)
	if disabled = "" AND (MandatoryFields<>"" OR IsNull(MandatoryFields)) then
		if instr(1, MandatoryFields & ";", Field & ";", vbTextCompare)>0 OR _
		    IsNull(MandatoryFields) then %>
			<span class="mandatory">(*)</span>
		<%end if
	end if
end sub


sub Form_CampiObbligatori() %>
	<tr>
		<td class="label">&nbsp;</td>
		<td colspan="3" class="input">
			<span class="mandatory">
				(*) <%= ChooseByLanguage("Campi obbligatori.", "Mandatory fields.", "Vorgeschriebene Felder.", "Champs obligatoires.", "Campos obligatorios.") %>
			</span>
		</td>
	</tr>
<% end sub


function Form_Errore(label, errore)
	if cString(errore)<>"" then %>
		<tr>
			<td colspan="4" class="errore">
				<span class="errore">
					<% if label<>"" then %><%= label %><br><% end if %>
					<%= errore %>
				</span>
			</td>
		</tr>
		<% errore = ""
	end if
	Form_Errore = errore
end function


sub Form_Captcha(oCnt, disabled)
	if disabled = "" then %>
		<tr>
			<td class="label" rowspan="3">
				<%= ChooseByLanguage("Codice di Verifica", "Verification code", "", "Code de v&eacute;rification", "") %>
			</td>
			<td class="input captcha_img" colspan="3">
				<img src="../amministrazione/library/aspcaptcha.asp" alt="This Is CAPTCHA Image" width="86" height="21" />
			</td>
		</tr>
		<tr>
			<td colspan="3" class="label captcha_text">
			<%= ChooseByLanguage("Il codice &egrave; composto di 8 cifre numeriche", "The code is composed of 8 numeric digit", "", "Le code est compos&eacute; de 8 chiffres", "") %>
			</td>
		</tr>
		<tr>
			<td colspan="3" class="input captcha_input">
				<%= ChooseByLanguage("Digita il codice visualizzato:", "Enter the code here:", "", "Entrez le code affich&eacute;:", "") %>
				<input name="strCAPTCHA" type="text" id="strCAPTCHA" maxlength="8" />&nbsp;<span class="mandatory">(*)</span>
			</td>
		</tr>
	<% end if
end sub



function Form_CaptchaCheck()
	
	dim SessionCAPTCHA
	SessionCAPTCHA = Trim(Session("CAPTCHA"))
	Session("CAPTCHA") = vbNullString
	if Len(SessionCAPTCHA) < 1 then
        Form_CaptchaCheck = False
		Session("ERRORE")=Session("ERRORE") & "<br>" & ChooseByLanguage("Codice di Verifica non valido", "Verification code not valid", "", "Code de v&eacute;rification invalide", "")
        exit function
    end if
	if CStr(SessionCAPTCHA) = CStr(request("strCAPTCHA")) then
	    Form_CaptchaCheck = True
	else
	    Form_CaptchaCheck = False
		Session("ERRORE")=Session("ERRORE") & "<br>" & ChooseByLanguage("Codice di Verifica non valido", "Verification code not valid", "", "Code de v&eacute;rification invalide", "")
	end if
	
end function



sub Form_DatiContattoSintetica(oCnt, disabled, MandatoryFields)
	disabled = cString(disabled) %>
	<tr>
		<td class="label"><%= ChooseByLanguage("Nome", "Name", "Name", "Nom", "Nombre") %></td>
		<td colspan="3" class="input">
			<% CALL Form_TextField(disabled, "tft_NomeElencoIndirizzi", oCnt("NomeElencoIndirizzi"), "text", 40, 250)
			CALL Form_CampoObbligatorio(disabled, "tft_NomeElencoIndirizzi", MandatoryFields) %>
		</td>
	</tr>
	<tr>
		<td class="label"><%= ChooseByLanguage("Cognome", "Surname", "Familienname", "Pr&eacute;nom", "Apellido") %></td>
		<td colspan="3" class="input">
			<% CALL Form_TextField(disabled, "tft_CognomeElencoIndirizzi", oCnt("CognomeElencoIndirizzi"), "text", 40, 250)
			CALL Form_CampoObbligatorio(disabled, "tft_CognomeElencoIndirizzi", MandatoryFields) %>
		</td>
	</tr>
	<tr>
		<td class="label"><%= ChooseByLanguage("Telefono", "Phone", "Telefon", "T&eacute;l&eacute;phone", "Tel&eacute;fono") %></td>
		<td colspan="3" class="input">
			<% CALL Form_TextField(disabled, "tft_telefono", oCnt("telefono"), "text", 40, 250)
			CALL Form_CampoObbligatorio(disabled, "tft_telefono", MandatoryFields) %>
		</td>
	</tr>
	<tr>
		<td class="label">Fax</td>
		<td colspan="3" class="input">
			<% CALL Form_TextField(disabled, "tft_fax", oCnt("fax"), "text", 40, 250)
			CALL Form_CampoObbligatorio(disabled, "tft_fax", MandatoryFields) %>
		</td>
	</tr>
	<tr>
		<td class="label">Email</td>
		<td colspan="3" class="input">
			<% CALL Form_EMailField(disabled, "tft_email", oCnt("email"), "text", 66, 250)
			CALL Form_CampoObbligatorio(disabled, "tft_email", MandatoryFields) %>
		</td>
	</tr>
<% end sub 


' DropCountry = true mostra il drop down degli stati
sub Form_DatiContattoEx(oCnt, disabled, MandatoryFields,DropCountry)
	disabled = cString(disabled) %>
	<tr>
		<td class="label"><%= ChooseByLanguage("Nome", "Name", "Name", "Pr&eacute;nom", "Nombre") %></td>
		<td colspan="3" class="input">
			<% CALL Form_TextField(disabled, "tft_NomeElencoIndirizzi", oCnt("NomeElencoIndirizzi"), "text", 40, 250)
			CALL Form_CampoObbligatorio(disabled, "tft_NomeElencoIndirizzi", MandatoryFields) %>
		</td>
	</tr>
	<tr>
		<td class="label"><%= ChooseByLanguage("Cognome", "Surname", "Familienname", "Nom", "Apellido") %></td>
		<td colspan="3" class="input">
			<% CALL Form_TextField(disabled, "tft_CognomeElencoIndirizzi", oCnt("CognomeElencoIndirizzi"), "text", 40, 250)
			CALL Form_CampoObbligatorio(disabled, "tft_CognomeElencoIndirizzi", MandatoryFields) %>
		</td>
	</tr>
	<tr>
		<td class="label"><%= ChooseByLanguage("Indirizzo", "Address",  "Addresse","Adresse", "Direcci&oacute;n") %></td>
		<td colspan="3" class="input">
		<% CALL Form_TextField(disabled, "tft_IndirizzoElencoIndirizzi", oCnt("IndirizzoElencoIndirizzi"), "text", 66, 250)
		CALL Form_CampoObbligatorio(disabled, "tft_IndirizzoElencoIndirizzi", MandatoryFields) %>
		</td>
	</tr>
	<tr>
		<td class="label1"><%= ChooseByLanguage("Citt&agrave;", "City", "Stadt", "Ville", "Ciudad") %></td>
		<td class="input1">
			<% CALL Form_TextField(disabled, "tft_CittaElencoIndirizzi", oCnt("CittaElencoIndirizzi"), "text1", 20, 100)
			CALL Form_CampoObbligatorio(disabled, "tft_CittaElencoIndirizzi", MandatoryFields) %>
		</td>
		<td class="label2"><%= ChooseByLanguage("Cap", "P. code", "Rei&szlig;verschlu&szlig;", "Code postal", "C&oacute;digo postal") %></td>
		<td class="input2">
			<% CALL Form_TextField(disabled, "tft_CAPElencoIndirizzi", oCnt("CAPElencoIndirizzi"), "text2", 10, 50)
			CALL Form_CampoObbligatorio(disabled, "tft_CAPElencoIndirizzi", MandatoryFields) %>
		</td>
	</tr>
	<tr>
		<td class="label1"><%= ChooseByLanguage("Provincia", "Province", "Provinz", "D&eacute;partement", "provincia") %></td>
		<td class="input1">
			<% CALL Form_TextField(disabled, "tft_StatoProvElencoIndirizzi", oCnt("StatoProvElencoIndirizzi"), "text1", 20, 250)
			CALL Form_CampoObbligatorio(disabled, "tft_StatoProvElencoIndirizzi", MandatoryFields) %>
		</td>
		<td class="label2"><%= ChooseByLanguage("Nazione", "Country", "Zustand", "Pays", "Pa&iacute;s") %></td>
		<td class="input2">
		<% if DropCountry then 
				CALL dropDown(oCnt.conn, "SELECT * FROM stati", "CodiceIso", "nome", "tft_CountryElencoIndirizzi", oCnt("CountryElencoIndirizzi"), true, "class=""stati""", Session("LINGUA"))  
		else 
				CALL Form_TextField(disabled, "tft_CountryElencoIndirizzi", oCnt("CountryElencoIndirizzi"), "text2", 10, 250)
		 end if 
			CALL Form_CampoObbligatorio(disabled, "tft_CountryElencoIndirizzi", MandatoryFields) %>
		</td>
	</tr>
	<tr>
		<td class="label"><%= ChooseByLanguage("Telefono", "Phone", "Telefon", "T&eacute;l&eacute;phone", "Tel&eacute;fono") %></td>
		<td colspan="3" class="input">
			<% CALL Form_TextField(disabled, "tft_telefono", oCnt("telefono"), "text", 40, 250)
			CALL Form_CampoObbligatorio(disabled, "tft_telefono", MandatoryFields) %>
		</td>
	</tr>
	<tr>
		<td class="label">Fax</td>
		<td colspan="3" class="input">
			<% CALL Form_TextField(disabled, "tft_fax", oCnt("fax"), "text", 40, 250)
			CALL Form_CampoObbligatorio(disabled, "tft_fax", MandatoryFields) %>
		</td>
	</tr>
	<tr>
		<td class="label">Email</td>
		<td colspan="3" class="input">
			<% CALL Form_EMailField(disabled, "tft_email", oCnt("email"), "text", 66, 250)
			CALL Form_CampoObbligatorio(disabled, "tft_email", MandatoryFields) %>
		</td>
	</tr>
<% end sub 

' Senza drop degli stati
sub Form_DatiContatto(oCnt, disabled, MandatoryFields)
	CALL Form_DatiContattoEx(oCnt, disabled, MandatoryFields,false)
end sub

%>