<!--#INCLUDE FILE="../library/ClassCryptography.asp"-->
<% 
'classe per la gestione degli stili del next-web
class CssManager

	'costanti generali
	public STANDARD_CSS_CONTENT
	'oggetti e costanti per la definizione delle corrispondenze tra valori Editor e CSS
	public FONT_FAMILY
	public FONT_FAMILY_UNICODE
	public FONT_SIZE
	public FONT_SIZE_rapporto_conversione
	public FONT_SIZE_unita_editor
	public FONT_SIZE_unita_css
	public FONT_SIZE_unita_px
	public FONT_WEIGHT
	public FONT_WEIGHT_DEFAULT
	public FONT_WEIGHT_USER
	public FONT_STYLE
	public FONT_STYLE_DEFAULT
	public FONT_STYLE_USER
	public TEXT_ALIGN
	public TEXT_ALIGN_DEFAULT
	public TEXT_ALIGN_USER
	public LINE_HEIGHT
	public LINE_HEIGHT_DEFAULT
	public LINE_HEIGHT_USER
	public LINE_HEIGHT_rapporto_conversione
	public LINE_HEIGHT_unita_editor
	public LINE_HEIGHT_unita_css
	public LETTER_SPACING
	public LETTER_SPACING_DEFAULT
	public LETTER_SPACING_USER
	public LETTER_SPACING_rapporto_conversione
	public LETTER_SPACING_unita_editor
	public LETTER_SPACING_unita_css
	public TEXT_DECORATION_USER

	'**************************************************************************************************************
	'FUNZIONI DI INIZZIALIZZAZIONE
	'**************************************************************************************************************
	
	Private Sub Class_Initialize()
	
		'parte standard degli stili (ex contenuto del file standard.css)
		STANDARD_CSS_CONTENT = "body{padding:0px;margin:0px;}" + vbCrLf + _
							   "form,div#NextForm,div.nextform_div {margin: 0px auto;text-align: left;}" + vbCrLf + _
							   "h1,h2,h3,h4,h5,h6,ul li,ol li{margin:0px;padding:0px;}" + vbCrLf + _
							   "p{margin: 0px;padding: 0px;}" + vbCrLf + _
							   "img{border:0px;}" + vbCrLf + _
							   "div.dynamail{display: none;}" + vbCrLf + _
							   "img.ga{visibility:hidden;}" + vbCrLF + _
							   "div.dynamail{display:none;visibility:hidden;}"
	
		'..........................................................................................................
		'definizione elenchi di valore per le varie proprieta'
		'chiave: valore per CSS
		'valore: valore per EDITOR
		'..........................................................................................................
		'definizione valori font-family
		set FONT_FAMILY = Server.CreateObject("Scripting.Dictionary")
		FONT_FAMILY.CompareMode = vbTextCompare
		CALL FONT_FAMILY.Add("Arial, Helvetica, sans-serif", 							"Arial")	'DEFAULT
		CALL FONT_FAMILY.Add("Verdana, Geneva, Arial, Helvetica, sans-serif", 			"Verdana")
		CALL FONT_FAMILY.Add("Tahoma, Verdana, Geneva, Arial, Helvetica, sans-serif", 	"Tahoma")
		CALL FONT_FAMILY.Add("""Trebuchet MS"", Arial, Helvetica, sans-serif", 			"Trebuchet MS")
		CALL FONT_FAMILY.Add("""Times New Roman"", Times, serif", 						"Times New Roman")
		CALL FONT_FAMILY.Add("Georgia, ""Times New Roman"", Times, serif", 				"Georgia")
		CALL FONT_FAMILY.Add("""Courier New"", Courier, monospace", 					"Courier New")
		CALL FONT_FAMILY.Add("Calibri, Arial, Helvetica, sans-serif", 					"Calibri")
		FONT_FAMILY_UNICODE ="Arial Unicode MS"
		
		'definizione dati font-size
		FONT_SIZE_rapporto_conversione = 	6.25
		FONT_SIZE_unita_editor =			""	'px
		FONT_SIZE_unita_css =				"%"
		FONT_SIZE_unita_px = 				"px"
		set FONT_SIZE = Server.CreateObject("Scripting.Dictionary")
		CALL FONT_SIZE.Add(50, 		8)
		CALL FONT_SIZE.Add(56.25, 	9)
		CALL FONT_SIZE.Add(62.5, 	10)
		CALL FONT_SIZE.Add(68.75, 	11)
		CALL FONT_SIZE.Add(75, 		12)			'DEFAULT P
		CALL FONT_SIZE.Add(81.65, 	13)
		CALL FONT_SIZE.Add(87.5, 	14)
		CALL FONT_SIZE.Add(100, 	16)			'DEFAULT h1
		CALL FONT_SIZE.Add(112.5, 	18)
		CALL FONT_SIZE.Add(125, 	20)
		CALL FONT_SIZE.Add(150, 	24)
		CALL FONT_SIZE.Add(200, 	32)
		CALL FONT_SIZE.Add(300, 	48)
		
		'definizione dati font-weight
		set FONT_WEIGHT = Server.CreateObject("Scripting.Dictionary")
		FONT_WEIGHT.CompareMode = vbTextCompare
		CALL FONT_WEIGHT.Add("normal", 	"plain")	'DEFAULT
		CALL FONT_WEIGHT.Add("bold", 	"bold")
		FONT_WEIGHT_DEFAULT = "normal"
		'definizione font-weight per interfaccia utenti
		set FONT_WEIGHT_USER = Server.CreateObject("Scripting.Dictionary")
		FONT_WEIGHT_USER.CompareMode = vbTextCompare
		CALL FONT_WEIGHT_USER.Add("normal", "normale (plain)")	'DEFAULT
		CALL FONT_WEIGHT_USER.Add("bold", 	"grassetto (bold)")
		
		'definizione dati font-style
		set FONT_STYLE = Server.CreateObject("Scripting.Dictionary")
		FONT_STYLE.CompareMode = vbTextCompare
		CALL FONT_STYLE.Add("normal", 	"plain")	'DEFAULT
		CALL FONT_STYLE.Add("italic", 	"italic")
		FONT_STYLE_DEFAULT = "normal"
		'definizione font-style per interfaccia utenti
		set FONT_STYLE_USER = Server.CreateObject("Scripting.Dictionary")
		FONT_STYLE_USER.CompareMode = vbTextCompare
		CALL FONT_STYLE_USER.Add("normal", 	"normale (plain)")	'DEFAULT
		CALL FONT_STYLE_USER.Add("italic", 	"corsivo (italic)")
				
		'definizione text-align
		set TEXT_ALIGN = Server.CreateObject("Scripting.Dictionary")
		TEXT_ALIGN.CompareMode = vbTextCompare
		CALL TEXT_ALIGN.Add("left", 	"left")		'DEFAULT
		CALL TEXT_ALIGN.Add("center", 	"center")
		CALL TEXT_ALIGN.Add("right", 	"right")
		CALL TEXT_ALIGN.Add("justify", 	"full")
		TEXT_ALIGN_DEFAULT = "left"
		'definizinoe text-align per interfaccia utenti
		set TEXT_ALIGN_USER = Server.CreateObject("Scripting.Dictionary")
		TEXT_ALIGN_USER.CompareMode = vbTextCompare
		CALL TEXT_ALIGN_USER.Add("left", 	"a sinistra")		'DEFAULT
		CALL TEXT_ALIGN_USER.Add("center", 	"centrato")
		CALL TEXT_ALIGN_USER.Add("right", 	"a destra")
		CALL TEXT_ALIGN_USER.Add("justify", 	"giustificato")
		
		'definizione line-height (su editor: fixed line-spacing)
		LINE_HEIGHT_rapporto_conversione = 	1
		LINE_HEIGHT_unita_editor =			""	'%
		LINE_HEIGHT_unita_css =				"%"
		set LINE_HEIGHT = Server.CreateObject("Scripting.Dictionary")
		'il primo valore del LINE_HEIGHT deve essere quello di DEFAULT causa compilatore
		CALL LINE_HEIGHT.Add(100, 	100)		'DEFAULT
		CALL LINE_HEIGHT.Add(110, 	110)
		CALL LINE_HEIGHT.Add(120, 	120)
		CALL LINE_HEIGHT.Add(130, 	130)
		CALL LINE_HEIGHT.Add(140, 	140)
		CALL LINE_HEIGHT.Add(150, 	150)
		CALL LINE_HEIGHT.Add(160, 	160)
		CALL LINE_HEIGHT.Add(180, 	180)
		CALL LINE_HEIGHT.Add(200, 	200)
		CALL LINE_HEIGHT.Add(250, 	250)
		CALL LINE_HEIGHT.Add(300, 	300)
		LINE_HEIGHT_DEFAULT = 100
		'definizione line height per interfaccia utente
		set LINE_HEIGHT_USER = Server.CreateObject("Scripting.Dictionary")
		CALL LINE_HEIGHT_USER.Add(100, 	"normale")
		CALL LINE_HEIGHT_USER.Add(110, 	"+ 10%")
		CALL LINE_HEIGHT_USER.Add(120, 	"+ 20%")
		CALL LINE_HEIGHT_USER.Add(130, 	"+ 30%")
		CALL LINE_HEIGHT_USER.Add(140, 	"+ 40%")
		CALL LINE_HEIGHT_USER.Add(150, 	"+ 50% (1,5 righe)")
		CALL LINE_HEIGHT_USER.Add(160, 	"+ 60%")
		CALL LINE_HEIGHT_USER.Add(180, 	"+ 80%")
		CALL LINE_HEIGHT_USER.Add(200, 	"+ 100% (doppia)")
		CALL LINE_HEIGHT_USER.Add(300, 	"+ 200% (tripla)")
		
		'definizione letter-spacing (su editor: char-spacing)
		LETTER_SPACING_rapporto_conversione = 	1
		LETTER_SPACING_unita_editor =			""	'%
		LETTER_SPACING_unita_css =				"px"
		set LETTER_SPACING = Server.CreateObject("Scripting.Dictionary")
		CALL LETTER_SPACING.Add(0, 	0)
		CALL LETTER_SPACING.Add(1, 	1)
		CALL LETTER_SPACING.Add(2, 	2)
		CALL LETTER_SPACING.Add(3, 	3)
		CALL LETTER_SPACING.Add(4, 	4)
		CALL LETTER_SPACING.Add(5, 	5)
		CALL LETTER_SPACING.Add(6, 	6)
		CALL LETTER_SPACING.Add(7, 	7)
		CALL LETTER_SPACING.Add(8,	8)
		CALL LETTER_SPACING.Add(9,	9)
		CALL LETTER_SPACING.Add(10,	10)
		LETTER_SPACING_DEFAULT = 0
		'definizione letter spacingi per interfaccia utente
		set LETTER_SPACING_USER = Server.CreateObject("Scripting.Dictionary")
		CALL LETTER_SPACING_USER.Add(0, 	"normale")
		CALL LETTER_SPACING_USER.Add(1, 	"+ 1 pixel")
		CALL LETTER_SPACING_USER.Add(2, 	"+ 2 pixel")
		CALL LETTER_SPACING_USER.Add(3, 	"+ 3 pixel")
		CALL LETTER_SPACING_USER.Add(4, 	"+ 4 pixel")
		CALL LETTER_SPACING_USER.Add(5, 	"+ 5 pixel")
		CALL LETTER_SPACING_USER.Add(6, 	"+ 6 pixel")
		CALL LETTER_SPACING_USER.Add(7, 	"+ 7 pixel")
		CALL LETTER_SPACING_USER.Add(8,		"+ 8 pixel")
		CALL LETTER_SPACING_USER.Add(9,		"+ 9 pixel")
		CALL LETTER_SPACING_USER.Add(10,	"+ 10 pixel")
		
		'definizione dati text-decoration per html ed interfaccia utente (NON USATO IN EDITOR)
		set TEXT_DECORATION_USER = Server.CreateObject("Scripting.Dictionary")
		TEXT_DECORATION_USER.CompareMode = vbTextCompare
		CALL TEXT_DECORATION_USER.Add("none", "non sottolineato")
		CALL TEXT_DECORATION_USER.Add("underline", 	"sottolineato")	'DEFAULT
	end sub
		
		
	Private Sub Class_Terminate()
		'distruzione oggetti per corrispondenze
		set FONT_FAMILY = nothing
		set FONT_SIZE = nothing
		set FONT_WEIGHT = nothing
		set FONT_WEIGHT_USER = nothing
		set FONT_STYLE = nothing
		set FONT_STYLE_USER = nothing
		set TEXT_ALIGN = nothing
		set TEXT_ALIGN_USER = nothing
		set LINE_HEIGHT = nothing
		set LINE_HEIGHT_USER = nothing
		set LETTER_SPACING = nothing
		set LETTER_SPACING_USER = nothing
	end sub
	
	
	'**************************************************************************************************************
	'FUNZIONI PUBBLICHE PER LA CONVERSIONE DEI VALORI
	'**************************************************************************************************************
	
	'...........................................................................
	'gestione font size
	'...........................................................................
	
	'converte il valore della dimensione carattere da CSS/database/HTML a valori per editor
	public function FontSize_CSS_to_EditorXml(CSS_value)
		CSS_value = cReal(CSS_value)
		'ricerca valore corrispondente per editor
		if FONT_SIZE.exists(CSS_value) then
			FontSize_CSS_to_EditorXml = FONT_SIZE(CSS_value)
		else
			FontSize_CSS_to_EditorXml = NULL
		end if
	end function
	
	
	'converte il valore della dimensione carattere dal valore gestito dall'editor nel valore per il CSS/database/HTML
	public function FontSize_EditorXml_to_CSS(EditorXml_value)
		EditorXml_value = cInteger(EditorXml_value)
		FontSize_EditorXml_to_CSS = FormatNumericValue(GetKeyByValue(FONT_SIZE, EditorXml_value), "")
	end function
	
	
	'ritorna la proprieta' completa per l'xml da mandare all'editor
	public function FontSize_EditorXml(CSS_value)
		'deve convertire anche il valore da css/database/html a valore per editor 
		FontSize_EditorXml = PropertyForEditorXML("font-size", FormatNumericValue(FontSize_CSS_to_EditorXml(CSS_value), FONT_SIZE_unita_editor))
	end function
	
	
	'ritorna la proprieta' completa per il css/html
	public function FontSize_CSS(CSS_value)
		'deve aggiungere l'eventuale unita' di misura per css
		FontSize_CSS = PropertyForCss("font-size", FormatNumericValue(Css_value, FONT_SIZE_unita_css))
	end function
	
	
	'ritorna la proprieta' completa per il css/html
	public function FontSize_px_CSS(CSS_value)
		'deve aggiungere l'eventuale unita' di misura per css in pixel
		FontSize_px_CSS = PropertyForCss("font-size", FormatNumericValue(FontSize_CSS_to_EditorXml(CSS_value), FONT_SIZE_unita_px))
	end function
	
	
	'...........................................................................
	'gestione font family
	'...........................................................................
	
	'converte il valore del font-family dal CSS/database/HTML a valori per editor
	public function FontFamily_CSS_to_EditorXml(CSS_value)
		CSS_value = cString(CSS_value)
		if FONT_FAMILY.exists(CSS_value) then
			FontFamily_CSS_to_EditorXml = FONT_FAMILY(CSS_value)
		else
			FontFamily_CSS_to_EditorXml = NULL
		end if
	end function
	
	
	'converte il valore da editor a CSS/database/HTML
	public function FontFamily_EditorXml_to_CSS(EditorXml_value)
		'replace aggiunto il 15/10/2012 per salvare anche le pagine in arial unicode per gestione del cinese.
		EditorXml_value = replace(cString(EditorXml_value), " Unicode MS", "")
		
		response.write "EditorXml_value=" & EditorXml_value & "<br>"
		response.write GetKeyByValue(FONT_FAMILY, EditorXml_value)
		
		FontFamily_EditorXml_to_CSS = replace(GetKeyByValue(FONT_FAMILY, EditorXml_value), """", "'")
	end function
	
	
	'ritorna la proprieta' completa per l'xml da mandare all'editor
	public function FontFamily_EditorXml(CSS_value, OnlyUnicode)
		'converte valore da font-family per css a editor
		if OnlyUnicode then
			'manda solo Arial Unicode per gestione cinese e lingue aggiuntive
			FontFamily_EditorXml = PropertyForEditorXML("font-family", FONT_FAMILY_UNICODE)
		else
			FontFamily_EditorXml = PropertyForEditorXML("font-family", FontFamily_CSS_to_EditorXml(CSS_value))
		end if
	end function
	
	
	'ritorna la proprieta' completa per il css/html
	public function FontFamily_CSS(CSS_value)
		FontFamily_CSS = PropertyForCss("font-family", CSS_value)
	end function
	
	
	'...........................................................................
	'gestione color
	'...........................................................................
	
	'ritorna la proprieta' completa per l'xml da mandare all'editor
	public function Color_EditorXml(CSS_value)
		Color_EditorXml = PropertyForEditorXML("color", CSS_value)
	end function
	
	
	'ritorna la proprieta' completa per il css/html
	public function Color_CSS(CSS_value)
		Color_CSS = PropertyForCss("color", CSS_value)
	end function
	
	
	'...........................................................................
	'gestione font-weight
	'...........................................................................
	
	'converte il valore del font-weight dal CSS/database/HTML a valori per editor
	public function FontWeight_CSS_to_EditorXml(CSS_value)
		CSS_value = cString(CSS_value)
		if FONT_WEIGHT.exists(CSS_value) then
			FontWeight_CSS_to_EditorXml = FONT_WEIGHT(CSS_value)
		else
			FontWeight_CSS_to_EditorXml = NULL
		end if
	end function
	
	
	'converte il valore da editor a CSS/database/HTML
	public function FontWeight_EditorXml_to_CSS(EditorXml_value)
		EditorXml_value = cString(EditorXml_value)
		FontWeight_EditorXml_to_CSS = GetKeyByValue(FONT_WEIGHT, EditorXml_value)
	end function
	
	
	'ritorna la proprieta' completa per l'xml da mandare all'editor
	public function FontWeight_EditorXml(CSS_value)
		FontWeight_EditorXml = PropertyForEditorXML("font-weight", FontWeight_CSS_to_EditorXml(CSS_value))
	end function
	
	
	'ritorna la proprieta' completa per il css/html
	public function FontWeight_CSS(CSS_value)
		FontWeight_CSS = PropertyForCss("font-weight", CSS_value)
	end function
	
	
	'...........................................................................
	'gestione font-style
	'...........................................................................
	
	'converte il valore del font-style dal CSS/database/HTML a valori per editor
	public function FontStyle_CSS_to_EditorXml(CSS_value)
		CSS_value = cString(CSS_value)
		if FONT_STYLE.exists(CSS_value) then
			FontStyle_CSS_to_EditorXml = FONT_STYLE(CSS_value)
		else
			FontStyle_CSS_to_EditorXml = NULL
		end if
	end function
	
	
	'converte il valore da editor a CSS/database/HTML
	public function FontStyle_EditorXml_to_CSS(EditorXml_value)
		EditorXml_value = cString(EditorXml_value)
		FontStyle_EditorXml_to_CSS = GetKeyByValue(FONT_STYLE, EditorXml_value)
	end function
	
	
	'ritorna la proprieta' completa per l'xml da mandare all'editor
	public function FontStyle_EditorXml(CSS_value)
		FontStyle_EditorXml = PropertyForEditorXML("font-style", FontStyle_CSS_to_EditorXml(CSS_value))
	end function
	
	
	'ritorna la proprieta' completa per il css/html
	public function FontStyle_CSS(CSS_value)
		FontStyle_CSS = PropertyForCss("font-style", CSS_value)
	end function
	
	
	'...........................................................................
	'gestione line-height
	'...........................................................................
	
	'converte il valore del line-height dal CSS/database/HTML a valori per editor
	public function LineHeight_CSS_to_EditorXml(CSS_value)
		CSS_value = cReal(CSS_value)
		if LINE_HEIGHT.exists(CSS_value) then
			'LineHeight_CSS_to_EditorXml = LINE_HEIGHT(CSS_value)
			LineHeight_CSS_to_EditorXml = 100
		else
			LineHeight_CSS_to_EditorXml = NULL
		end if
	end function
	
	
	'converte il valore da editor a CSS/database/HTML
	public function LineHeight_EditorXml_to_CSS(EditorXml_value)
		'EditorXml_value = cReal(EditorXml_value)
		'LineHeight_EditorXml_to_CSS = FormatNumericValue(GetKeyByValue(LINE_HEIGHT, EditorXml_value), "")
		LineHeight_EditorXml_to_CSS = ""
	end function
	
	
	'ritorna la proprieta' completa per l'xml da mandare all'editor
	public function LineHeight_EditorXml(CSS_value)
		'LineHeight_EditorXml = PropertyForEditorXML("line-height", FormatNumericValue(LineHeight_CSS_to_EditorXml(CSS_value), LINE_HEIGHT_unita_editor))
		if not IsNull(Css_value) then
			LineHeight_EditorXml = PropertyForEditorXML("line-height", FormatNumericValue(100, LINE_HEIGHT_unita_editor))
		end if
	end function
	
	
	'ritorna la proprieta' completa per il css/html
	public function LineHeight_CSS(CSS_value)
		'LineHeight_CSS = PropertyForCss("line-height", FormatNumericValue(Css_value, LINE_HEIGHT_unita_css))
		LineHeight_CSS = ""
	end function
	
	
	'...........................................................................
	'gestione letter spacing
	'...........................................................................
	
	'converte il valore del letter-spacing dal CSS/database/HTML a valori per editor
	public function LetterSpacing_CSS_to_EditorXml(CSS_value)
		CSS_value = cReal(CSS_value)
		if LETTER_SPACING.exists(CSS_value) then
			LetterSpacing_CSS_to_EditorXml = LETTER_SPACING(CSS_value)
		else
			LetterSpacing_CSS_to_EditorXml = NULL
		end if
	end function
	
	
	'converte il valore da editor a CSS/database/HTML
	public function LetterSpacing_EditorXml_to_CSS(EditorXml_value)
		EditorXml_value = cReal(EditorXml_value)
		LetterSpacing_EditorXml_to_CSS = FormatNumericValue(GetKeyByValue(LETTER_SPACING, EditorXml_value), "")
	end function
	
	
	'ritorna la proprieta' completa per l'xml da mandare all'editor
	public function LetterSpacing_EditorXml(CSS_value)
		LetterSpacing_EditorXml = PropertyForEditorXML("letter-spacing", FormatNumericValue(LetterSpacing_EditorXml_to_CSS(CSS_value), LETTER_SPACING_unita_editor))
	end function
	
	
	'ritorna la proprieta' completa per il css/html
	public function LetterSpacing_CSS(CSS_value)
		LetterSpacing_CSS = PropertyForCss("letter-spacing", FormatNumericValue(Css_value, LETTER_SPACING_unita_css))
	end function
	
	
	'...........................................................................
	'gestione text-align
	'...........................................................................
	
	'converte il valore del text-align dal CSS/database/HTML a valori per editor
	public function TextAlign_CSS_to_EditorXml(CSS_value)
		CSS_value = cString(CSS_value)
		if TEXT_ALIGN.exists(CSS_value) then
			TextAlign_CSS_to_EditorXml = TEXT_ALIGN(CSS_value)
		else
			TextAlign_CSS_to_EditorXml = NULL
		end if
	end function
	
	
	'converte il valore da editor a CSS/database/HTML
	public function TextAlign_EditorXml_to_CSS(EditorXml_value)
		EditorXml_value = cString(EditorXml_value)
		TextAlign_EditorXml_to_CSS = GetKeyByValue(TEXT_ALIGN, EditorXml_value)
	end function
	
	
	'ritorna la proprieta' completa per l'xml da mandare all'editor
	public function TextAlign_EditorXml(CSS_value)
		TextAlign_EditorXml = PropertyForEditorXML("text-align", TextAlign_CSS_to_EditorXml(CSS_value))
	end function
	
	
	'ritorna la proprieta' completa per il css/html
	public function TextAlign_CSS(CSS_value)
		TextAlign_CSS = PropertyForCss("text-align", CSS_value)
	end function
	
	
	'...........................................................................
	'gestione text-decoration (non presente per Editor)
	'...........................................................................
	
	'ritorna la proprieta' completa per il css/html
	public function TextDecoration_CSS(CSS_value)
		TextDecoration_CSS = PropertyForCss("text-decoration", CSS_value)
	end function
	
	
	'...........................................................................
	'gestione background-color (non presente per editor)
	'...........................................................................
	
	'ritorna la proprieta' completa per il css/html
	public function BackgroundColor_CSS(CSS_value)
		BackgroundColor_CSS = PropertyForCss("background-color", CSS_value)
	end function
	
	
	'**************************************************************************************************************
	'FUNZIONI PRIVATE PER LA GESTIONE DEI VALORI
	'**************************************************************************************************************
	
	'formatta il valore numerico con la relativa unita' di misura
	public function FormatNumericValue(Value, Unit)
		if IsNull(value) then
			FormatNumericValue = NULL
		else
			FormatNumericValue = replace(cString(Value), ",", ".") + Unit
		end if
	end function
	
	
	'ritorna la chiave corrispondente al valore nell'insieme di corrispondenze indicato
	'la chiave rappresenta il valore per CSS/database/HTML, il valore per XML/editor
	Public function GetKeyByValue(Dictio, value)
		dim key
		GetKeyByValue = NULL
		for each key in Dictio.keys
			if CString(Dictio(key)) = CString(value) then
				GetKeyByValue = key
				Exit For
			end if
		next
	end function
	
	
	'compone la proprieta' per l'XML
	private function PropertyForEditorXML(name, value)
		if IsNull(value) then
			PropertyForEditorXML = ""
		else
			PropertyForEditorXML = "				" + name + ": " + value + ";" + vbCrLf
		end if
	end function
	
	
	'compone proprieta' per CSS/HTML
	private function PropertyForCSS(name, value)
		if IsNull(value) then
			PropertyForCSS = NULL
		else
			PropertyForCSS = name + ":" + value + ";"' + vbCrLf
		end if
	end function
	

	'**************************************************************************************************************
	'FUNZIONI che operano da e sul DATABASE
	'**************************************************************************************************************
	
	
	'cancella stili esistenti reinserendo gli stili di default
	public sub ResetDbToDefault(conn, web_id)
		dim sql
		
		'cancella stili attuali
		sql = "DELETE FROM tb_css_groups WHERE grp_id_webs=" & web_id
		CALL conn.execute(sql, , adExecuteNoRecords)
		
		'rigenera stili di default
		CALL GenerateDbDefault(conn, web_id)
		
		'rignera files degli stili
		CALL GenerateCssFile(conn, web_id)
	end sub
	
	
	'inserisce nel database la struttura dati con i valori di default
	public sub GenerateDbDefault(conn, web_id)
		
		dim rsG, rsS, sql
		set rsG = Server.CreateObject("ADODB.RecordSet")
		set rsS = Server.CreateObject("ADODB.RecordSet")
		
		sql = "SELECT * FROM tb_css_groups WHERE grp_id_webs=" & web_id
		rsG.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
		
		sql = "SELECT * FROM tb_css_styles " + _
			  "WHERE style_grp_id IN (SELECT grp_id FROM tb_css_groups WHERE grp_id_webs=" & web_id & ")"
		rsS.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		
		'.......................................................................
		'TESTO STRUTTURATO
		'inserimento testata
		rsG.AddNew
		rsG("grp_id_webs") = web_id
		rsG("grp_name") = "Testo strutturato"
		rsG("grp_name_class") = LAYER_S_TEXT_NAME
		rsG("grp_checksum") = ""
		rsG("grp_for_editor") = true
		rsG("grp_for_file") = true
		CALL SetUpdateParamsRS(rsG, "grp_", true)
		rsG.Update
		
		'inserimento classi e stili
		'H1
		rsS.AddNew
		rsS("style_grp_id") = 			rsG("grp_id")
		rsS("style_class") = 			"H1"
		rsS("style_description") = 		"Titolo H1"
		rsS("style_font_family") = 		"Arial, Helvetica, sans-serif"
		rsS("style_font_size") = 		100
		rsS("style_color") = 			"#000000"
		rsS("style_line_height") = 		100
		rsS("style_letter_spacing") = 	0
		rsS("style_text_align") = 		"left"
		CALL SetUpdateParamsRS(rsS, "style_", true)
		rsS.Update
		
		'H2
		rsS.AddNew
		rsS("style_grp_id") = 			rsG("grp_id")
		rsS("style_class") = 			"H2"
		rsS("style_description") = 		"Titolo H2"
		rsS("style_font_family") = 		"Arial, Helvetica, sans-serif"
		rsS("style_font_size") = 		87.5
		rsS("style_color") = 			"#000000"
		rsS("style_line_height") = 		100
		rsS("style_letter_spacing") = 	0
		rsS("style_text_align") = 		"left"
		CALL SetUpdateParamsRS(rsS, "style_", true)
		rsS.Update
		
		'H3
		rsS.AddNew
		rsS("style_grp_id") = 			rsG("grp_id")
		rsS("style_class") = 			"H3"
		rsS("style_description") = 		"Titolo H3"
		rsS("style_font_family") = 		"Arial, Helvetica, sans-serif"
		rsS("style_font_size") = 		81.65
		rsS("style_color") = 			"#000000"
		rsS("style_line_height") = 		100
		rsS("style_letter_spacing") = 	0
		rsS("style_text_align") = 		"left"
		CALL SetUpdateParamsRS(rsS, "style_", true)
		rsS.Update
		
		'H4
		rsS.AddNew
		rsS("style_grp_id") = 			rsG("grp_id")
		rsS("style_class") = 			"H4"
		rsS("style_description") = 		"Titolo H4"
		rsS("style_font_family") = 		"Arial, Helvetica, sans-serif"
		rsS("style_font_size") = 		75
		rsS("style_color") = 			"#000000"
		rsS("style_line_height") = 		100
		rsS("style_letter_spacing") = 	0
		rsS("style_text_align") = 		"left"
		CALL SetUpdateParamsRS(rsS, "style_", true)
		rsS.Update
		
		'H5
		rsS.AddNew
		rsS("style_grp_id") = 			rsG("grp_id")
		rsS("style_class") = 			"H5"
		rsS("style_description") = 		"Titolo H5"
		rsS("style_font_family") = 		"Arial, Helvetica, sans-serif"
		rsS("style_font_size") = 		75
		rsS("style_color") = 			"#000000"
		rsS("style_line_height") = 		100
		rsS("style_letter_spacing") = 	0
		rsS("style_text_align") = 		"left"
		CALL SetUpdateParamsRS(rsS, "style_", true)
		rsS.Update
		
		'H6
		rsS.AddNew
		rsS("style_grp_id") = 			rsG("grp_id")
		rsS("style_class") = 			"H6"
		rsS("style_description") = 		"Titolo H6"
		rsS("style_font_family") = 		"Arial, Helvetica, sans-serif"
		rsS("style_font_size") = 		75
		rsS("style_color") = 			"#000000"
		rsS("style_line_height") = 		100
		rsS("style_letter_spacing") = 	0
		rsS("style_text_align") = 		"left"
		CALL SetUpdateParamsRS(rsS, "style_", true)
		rsS.Update
		
		'P
		rsS.AddNew
		rsS("style_grp_id") = 			rsG("grp_id")
		rsS("style_class") = 			"P"
		rsS("style_description") = 		"Testo del paragrafo"
		rsS("style_font_family") = 		"Arial, Helvetica, sans-serif"
		rsS("style_font_size") = 		75
		rsS("style_color") = 			"#000000"
		rsS("style_line_height") = 		100
		rsS("style_letter_spacing") = 	0
		rsS("style_text_align") = 		"left"
		CALL SetUpdateParamsRS(rsS, "style_", true)
		rsS.Update
		
		'A
		rsS.AddNew
		rsS("style_grp_id") = 			rsG("grp_id")
		rsS("style_class") = 			"A"
		rsS("style_description") = 		"Link"
		rsS("style_pseudoclass") = 		NULL
		rsS("style_color") = 			"#0000FF"
		rsS("style_letter_spacing") = 	0
		rsS("style_text_decoration") = 	"underline"
		rsS("style_background_color") = NULL
		CALL SetUpdateParamsRS(rsS, "style_", true)
		rsS.Update
		
		'A:hover
		rsS.AddNew
		rsS("style_grp_id") = 			rsG("grp_id")
		rsS("style_class") = 			"A"
		rsS("style_description") = 		"Link al passaggio del mouse"
		rsS("style_pseudoclass") = 		":hover"
		rsS("style_color") = 			"#0000FF"
		rsS("style_letter_spacing") = 	0
		rsS("style_text_decoration") = 	"underline"
		rsS("style_background_color") = "transparent"
		CALL SetUpdateParamsRS(rsS, "style_", true)
		rsS.Update
		
		'A:visited
		'rsS.AddNew
		'rsS("style_grp_id") = 			rsG("grp_id")
		'rsS("style_class") = 			"A"
		'rsS("style_description") = 		"Link gi&agrave; visitato"
		'rsS("style_pseudoclass") = 		":visited"
		'rsS("style_color") = 			"#800080"
		'rsS("style_letter_spacing") = 	0
		'rsS("style_text_decoration") = 	"underline"
		'rsS("style_background_color") =	NULL
		'CALL SetUpdateParamsRS(rsS, "style_", true)
		'rsS.Update

		
		'.......................................................................
		'TESTO NORMALE
		'inserimento testata
		rsG.AddNew
		rsG("grp_id_webs") = web_id
		rsG("grp_name") = "Testo normale"
		rsG("grp_name_class") = LAYER_TEXT_NAME
		rsG("grp_checksum") = ""
		rsG("grp_for_editor") = true
		rsG("grp_for_file") = false
		CALL SetUpdateParamsRS(rsG, "grp_", true)
		rsG.Update
		'inserimento classi e stili
		
		'P
		rsS.AddNew
		rsS("style_grp_id") = 			rsG("grp_id")
		rsS("style_class") = 			"P"
		rsS("style_description") = 		"Testo del paragrafo"
		rsS("style_font_family") = 		"Arial, Helvetica, sans-serif"
		rsS("style_font_size") = 		75
		rsS("style_color") = 			"#000000"
		rsS("style_font_weight") = 		"normal"
		rsS("style_font_style") = 		"normal"
		rsS("style_line_height") = 		100
		rsS("style_letter_spacing") = 	0
		rsS("style_text_align") = 		"left"
		CALL SetUpdateParamsRS(rsS, "style_", true)
		rsS.Update
		
		'A
		rsS.AddNew
		rsS("style_grp_id") = 			rsG("grp_id")
		rsS("style_class") = 			"A"
		rsS("style_description") = 		"Link"
		rsS("style_pseudoclass") = 		NULL
		rsS("style_color") = 			"#0000FF"
		rsS("style_font_weight") = 		"normal"
		rsS("style_font_style") = 		"normal"
		rsS("style_letter_spacing") = 	0
		rsS("style_text_decoration") = 	"underline"
		rsS("style_background_color") = NULL
		CALL SetUpdateParamsRS(rsS, "style_", true)
		rsS.Update
		
		'A:hover
		rsS.AddNew
		rsS("style_grp_id") = 			rsG("grp_id")
		rsS("style_class") = 			"A"
		rsS("style_description") = 		"Link al passaggio del mouse"
		rsS("style_pseudoclass") = 		":hover"
		rsS("style_color") = 			"#0000FF"
		rsS("style_font_weight") = 		NULL
		rsS("style_font_style") = 		NULL
		rsS("style_letter_spacing") = 	NULL
		rsS("style_text_decoration") = 	"underline"
		rsS("style_background_color") = "transparent"
		CALL SetUpdateParamsRS(rsS, "style_", true)
		rsS.Update
		
		'A:visited
		'rsS.AddNew
		'rsS("style_grp_id") = 			rsG("grp_id")
		'rsS("style_class") = 			"A"
		'rsS("style_description") = 		"Link gi&agrave; visitato"
		'rsS("style_pseudoclass") = 		":visited"
		'rsS("style_color") = 			"#800080"
		'rsS("style_font_weight") = 		NULL
		'rsS("style_font_style") = 		NULL
		'rsS("style_letter_spacing") = 	NULL
		'rsS("style_text_decoration") = 	"underline"
		'rsS("style_background_color") = NULL
		'CALL SetUpdateParamsRS(rsS, "style_", true)
		'rsS.Update
		
		rsS.close
		rsG.close
		
		'generazione checksum degli stili
		sql = "SELECT * FROM tb_css_groups WHERE grp_id_webs=" & web_id
		rsG.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		while not rsG.eof
			rsG("grp_checksum") = GetChecksum(conn, rsS, rsG("grp_id"))
			rsG.update
			rsG.movenext
		wend
		rsG.close
		
		set rsG = nothing
		set rsS = nothing
		
	end sub
	
	
	'funzione per generare e scrivere il file degli stili
	public sub GenerateCssFile(conn, web_id)
		
		CALL WriteFileContent(Application("IMAGE_PATH") & web_id & "\css\stili_testo.css", _
							  STANDARD_CSS_CONTENT + GenerateCss(conn, web_id, true), true)
		
	end sub
	
	
	'funzione che restituisce i css standard del NextWeb
	public function GetCssStandard(conn, web_id)
		GetCssStandard = STANDARD_CSS_CONTENT + GenerateCss(conn, web_id, false)
	end function
	
	
	public function GenerateCss(conn, web_id, ForFile)
		dim rsS, sql, css
		set rsS = Server.CreateObject("ADODB.RecordSet")
		
		'recupera elementi da scrivere nel file
		sql = " SELECT * FROM tb_css_groups INNER JOIN tb_css_styles ON tb_css_groups.grp_id = tb_css_styles.style_grp_id " + _
			  " WHERE grp_id_webs=" & web_id
		if ForFile then
			sql = sql & " AND ( "& SQL_IsTrue(conn, "grp_for_file") &" OR NOT "& SQL_IsNull(conn, "style_pseudoclass") & ")"
		end if
		rsS.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		'stili di default del sito.
		

		'genera stringa da scrivere nel file
		while not rsS.eof
			if ForFile then
				css = css + "body.cke_show_borders " & lCase(cString(rsS("style_class"))) & lCase(cString(rsS("style_pseudoclass"))) & ", "
			end if
			css = css + _
				  "div." & lCase(cString(rsS("grp_name_class"))) & " " & lCase(cString(rsS("style_class"))) & lCase(cString(rsS("style_pseudoclass"))) & "{" & _
				  FontFamily_CSS(		rsS("style_font_family")) & _
				  FontSize_CSS(			rsS("style_font_size")) & _
				  Color_CSS(			rsS("style_color")) & _
				  FontWeight_CSS(		rsS("style_font_weight")) & _
				  FontStyle_CSS(		rsS("style_font_style")) & _
				  LineHeight_CSS(		rsS("style_line_height")) & _
				  LetterSpacing_CSS(	rsS("style_letter_spacing")) & _
				  TextAlign_CSS(		rsS("style_text_align")) & _
				  TextDecoration_CSS(	rsS("style_text_decoration")) & _
				  BackgroundColor_CSS(	rsS("style_background_color")) & _
				  "}" + vbCrLf
			rsS.movenext
		wend
		
		rsS.close
		
		set rsS = nothing
		
		GenerateCss = css
	end function
	
	'funzione per generare l'xml per l'editor
	public function GenerateEditorXml(conn, web_id, OnlyUnicode)
		dim rsS, rsG, sql, xml
		set rsS = Server.CreateObject("ADODB.RecordSet")
		set rsG = Server.CreateObject("ADODB.RecordSet")
		
		'recupera gruppi di stili e genera xml
		sql = "SELECT * FROM tb_css_groups WHERE grp_id_webs=" & web_id & " AND " & SQL_IsTrue(conn, "grp_for_editor")
		rsG.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		
		xml = "<?xml version=""1.0"" encoding=""UTF-8""?>" + vbCrLf
		xml = xml + _
			  "<xml>" + vbCrLf
		
		while not rsG.eof
			'apre elenco stili del gruppo
			sql = EditorXML_Query(conn, rsG("grp_id"))
			rsS.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			if not rsS.eof then
				xml = xml + "	<group name=""" & rsG("grp_name") & """ checksum=""" & rsG("grp_checksum") & """>" & vbCrLf
				while not rsS.eof
					xml = xml + "		<style name=""" & rsS("style_class") & """>" + vbCrLf + _
						  		"			<values>" + vbCrLf + _
								FontFamily_EditorXml(		rsS("style_font_family"), OnlyUnicode) & _
								FontSize_EditorXml(			rsS("style_font_size")) & _
								Color_EditorXml(			rsS("style_color")) & _
								FontWeight_EditorXml(		rsS("style_font_weight")) & _
								FontStyle_EditorXml(		rsS("style_font_style")) & _
								LineHeight_EditorXml(		rsS("style_line_height")) & _
								LetterSpacing_EditorXml(	rsS("style_letter_spacing")) & _
								TextAlign_EditorXml(		rsS("style_text_align")) & _
						  		"			</values>" + vbCrLf + _
						  		"		</style>" + vbCrLf
					rsS.movenext
				wend
				xml = xml + "	</group>" + vbCrLf
			end if
			rsS.close
			rsG.movenext
		wend
		
		xml = xml + "</xml>" + vbCrLf
		
		rsG.close
		
		set rsS = nothing
		set rsG = nothing
		
		GenerateEditorXml = xml
	end function
	
	
	'ritorna query recordset degli stili per l'editor
	private function EditorXML_Query(conn, grp_id)
		EditorXML_Query = " SELECT style_class, " + _
						  " style_font_family, " + _
						  " style_font_size, " + _
						  " style_color, " + _
						  " style_font_weight, " + _
						  " style_font_style, " + _
						  " style_line_height, " + _
						  " style_letter_spacing, " + _
						  " style_text_align " + _
						  " FROM tb_css_styles " + _
						  " WHERE style_grp_id=" & grp_id & " AND "& SQL_IsNull(conn, "style_pseudoclass") + _
						  " ORDER BY left(style_class, 1) DESC, style_id"
	end function
	
	
	'**************************************************************************************************************
	'FUNZIONI che operano sul checksum
	'**************************************************************************************************************
	
	'funzione per la generazione del checksum per il gruppo indicato
	function GetChecksum(conn, rsS, grp_id)
		dim sql, Cripto
		'recupera stili
		sql = EditorXML_Query(conn, grp_id)
		rsS.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		
		'gestione checksum
		set Cripto = new CryptographyManager
		GetChecksum = Cripto.md5_of_string(rsS.GetString(adClipString, rsS.recordcount, ";", vbCrLf, "NULL"))
		set Cripto = nothing
		
		rsS.close
	end function
	
	
	'funzione per aggiornare il checksum per il gruppo indicato
	sub UpdateChecksum(conn, rsS, grp_id)
		dim sql
		sql = "UPDATE tb_css_groups SET grp_checksum='" + GetChecksum(conn, rsS, grp_id) + "' WHERE grp_id=" + grp_id
		CALL conn.execute(sql, , adExecuteNoRecords)
	end sub
	
	
end class

%>