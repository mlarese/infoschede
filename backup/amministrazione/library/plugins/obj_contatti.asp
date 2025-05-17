<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<!--#INCLUDE FILE="../TOOLS.ASP"-->
<!--#INCLUDE FILE="../TOOLS4PlugIn.ASP"-->
<!--#INCLUDE FILE="../TOOLS4Forms.ASP"-->
<!--#INCLUDE FILE="../CLASSCONFIGURATION.ASP"-->
<!--#INCLUDE FILE="../CLASSIndirizzarioLock.ASP"-->
<!--#INCLUDE FILE="../CLASS_MAILER.ASP"-->

<%
dim Config
set Config = new Configuration
'impostazione delle proprieta' di default
CALL Form_InitConfig(Config)

'impostazione delle proprieta' di default
Config.AddDefault "Rubrica", ""
Config.AddDefault "PageEmail",""
Config.AddDefault "Captcha","false"
Config.AddDefault "PageConfirm",""
Config.AddDefault "RubricaMailingList", ""
Config.AddDefault "EmailSenderAdministratorID", ""

'impostazione dati di default: da sovrascrivere via propriet&agrave; del plugin o dagli stili
Config.AddDefault "Object_IT", Request.ServerVariables("SERVER_NAME") & " - contattaci"
Config.AddDefault "Object_EN", Request.ServerVariables("SERVER_NAME") & " - contact us"
Config.AddDefault "Object_FR", Request.ServerVariables("SERVER_NAME") & " - contactez-nous"
Config.AddDefault "Object_DE", Request.ServerVariables("SERVER_NAME") & " - kontakt"
Config.AddDefault "Object_ES", Request.ServerVariables("SERVER_NAME") & " - contacta con nosotros"

Config.AddDefault "MailingList_IT", "Acconsento alla ricezione di email con le migliori offerte e promozioni da voi proposte."
Config.AddDefault "MailingList_EN", "I consent the reception of email with the best offers and promotions from You proposed."
Config.AddDefault "MailingList_FR", "Je consens &agrave; la réception de l'email avec le meilleur offres et des promotions de Vous propos&eacute;."
Config.AddDefault "MailingList_DE", "Zustimmung zur Aufnahme von email mit den besten Angeboten und Foerderungen von Sie vorgeschlagen."
Config.AddDefault "MailingList_ES", "Consentimiento a la recepci&ograve;n del email con los mejores ofrecidos y las promociones de usted propuesto."

Config.AddDefault "Message_Base_IT", "Compila i campi obbligatori (*) del modulo sottostante e appena possibile riceverai un'email di risposta alle tue richieste dal nostro personale."
Config.AddDefault "Message_Base_EN", "Fill in the form and you will receive a reply from our staff as soon as possible."
Config.AddDefault "Message_Base_FR", "Remplissez les champs obligatoires (*) du formulaire ci-dessous et dès que possible vous recevrez de la part de notre personnel un e-mail répondant à toutes vos questions."
Config.AddDefault "Message_Base_DE", "Füllen Sie bitte die Pflichtfelder (*) des nachfolgenden Formulars aus. Sie werden so schnell wie möglich auf Ihre Anfrage eine Rückantwort per E-Mail von unserem Personal erhalten."
Config.AddDefault "Message_Base_ES", "Llene los campos obligatorios (*) del siguiente formulario y en breve nuestro personal enviará un e-mail de respuesta a su solicitud."

Config.AddDefault "Message_OK_IT", "Il suo messaggio &egrave; stato spedito correttamente.<br> Ricever&agrave; al pi&ugrave; presto una e-mail di conferma."
Config.AddDefault "Message_OK_EN", "Your message has been sent correctly.<br> You'll receive a confirm email as soon as possible."
Config.AddDefault "Message_OK_FR", "Votre message a &eacute;t&eacute; envoy&eacute; correctement.<br> Vous recevrez un email de confirmation aussitôt que possible."
Config.AddDefault "Message_OK_DE", "Ihre Anzeige ist richtig gesendet worden.<br> Sie empfangen ein Bestätigungs-email so bald wie möglich."
Config.AddDefault "Message_OK_ES", "Su mensaje se ha enviado correctamente.<br> Usted recibirá un email del confirmar cuanto antes."

'caricamento proprieta' specifiche
Config.SetConfigurationString(Session("CONFSTR"))
dim disabled, cnt_ID, OBJ_contatto

set OBJ_contatto = new IndirizzarioLock

if request("CNT_ID")<>"" then
	disabled = " DISABLED "
	OBJ_contatto.loadFromDb(request("CNT_ID"))
else
	disabled = ""
	OBJ_contatto.LoadFromForm("isSocieta")
	if Request("salva")<> "" then
		'eventuale verifica del captcha
		if not cBoolean(Config("captcha"), false) OR _
		   Form_CaptchaCheck() then
			if OBJ_contatto.ValidateFields("CognomeElencoIndirizzi;NomeElencoIndirizzi;Noteelencoindirizzi", true) then
				cnt_ID = OBJ_contatto.InsertIntoDB()
				OBJ_contatto.ResyncConnection
				
				dim EmailUrl, RedirectUrl, rs
				set rs = server.createobject("adodb.recordset")
				'salvataggio andato a buon fine
				CALL PrepareFormUrls(Config, Config("PageEmail"), Config("PageConfirm"), EmailUrl, RedirectUrl)
				
				EmailUrl = EmailUrl & "&CNT_ID=" & cnt_ID
				RedirectUrl = RedirectUrl & "&CNT_ID=" & cnt_ID
				
				CALL SendPageFromAdminToContact(OBJ_contatto.conn, rs, Config, CBL(Config, "Object"), _
												EmailUrl, Config("EmailSenderAdministratorID"), cnt_id, true)
				
				'redirige a pagina di conferma (stessa dell'email)
				response.redirect RedirectUrl
				
				set rs = nothing
			end if
		end if
	end if
end if
%>
<form action="" method="post">
<table cellspacing="0" cellpadding="0" class="form">
		<input type="Hidden" name="tft_rubrica" value="<%= Config("Rubrica") %>">
		<input type="Hidden" name="IsSocieta" value="">
		<% If disabled <> "" AND CBL(Config, "Message_OK")<>"" then %>
			<tr>
				<td colspan="4" class="message_OK"><%= CBL(Config, "Message_OK") %></td>
			</tr>
		<% Else
			if CBL(Config, "Message_Base")<>"" then %>
				<tr>
					<td colspan="4" class="message_Base"><%= CBL(Config, "Message_Base") %></td>
				</tr>
			<% end if
			
			Session("ERRORE") = Form_Errore(Config("lbl_error"), Session("ERRORE"))

		end if
		
		CALL Form_DatiContatto(OBJ_contatto, disabled, "tft_NomeElencoIndirizzi;tft_CognomeElencoIndirizzi;tft_email")
		
		if cInteger(Config("RubricaMailingList"))>0 then
			dim value, sql
			
			if disabled = "" then
				'in compilazione: recupera informazioni da form
				value = request.ServerVariables("REQUEST_METHOD")<>"POST" OR instr(1, OBJ_contatto("rubrica"), Config("RubricaMailingList"), vbTextCompare)
			else
				'in recupero informazioni da database
				sql = "SELECT COUNT(*) FROM rel_rub_ind WHERE id_rubrica=" & cIntero(Config("RubricaMailingList")) & " AND id_indirizzo=" & cIntero(Obj_Contatto("IDElencoIndirizzi"))
				value = cInteger(GetValueList(Obj_contatto.conn, NULL, sql))>0
			end if%>
			<tr>
				<td class="label">&nbsp;</td>
				<td colspan="4" class="MailingList">
					<% CALL Form_CheckboxField( disabled,"tft_rubrica", Config("RubricaMailingList") ,"checkbox",chk(value)) %>
					<%= CBL(Config, "MailingList") %>
				</td>
			</tr>
		<% end if 
		if cBoolean(Config("captcha"), false) then 
			CALL Form_Captcha(OBJ_contatto, disabled)
		end if %>
		<tr>
			<td class="label"><%= config("lbl_messaggio") %></td>
			<td colspan="3" class="input">
				<% CALL Form_TextAreaField(disabled, "tft_NoteElencoIndirizzi", OBJ_contatto("NoteElencoIndirizzi"), "message", "12")
				CALL Form_CampoObbligatorio(disabled, "tft_NoteElencoIndirizzi", "tft_NoteElencoIndirizzi") %>
			</td>
		</tr>
		<% If disabled = "" then 
			CALL Form_CampiObbligatori()
		end if %>
		<tr> 
			<td colspan="4" class="privacy">
				<%= GetPrivacyText()%>
			</td>
		</tr>
		<% If disabled = "" then %>
			<tr>
				<td colspan="4" class="button">
					<input type="reset" name="annulla" id="annulla" value="<%= Config("lbl_undo") %>" class="submit">
					<input type="submit" name="salva" id="salva" value="<%= Config("lbl_save") %>" class="submit">
				</td>
			</tr>
		<% End If %>
	</table>
</form>	
<% 
set OBJ_contatto = nothing
set Config = nothing
%>