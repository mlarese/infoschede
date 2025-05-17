<%@ Language=VBScript CODEPAGE=65001 %>
<%'@ Language=VBScript CODEPAGE=1252 %>
<% option explicit %>

<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<%
	'server.scripttimeout = 50
	response.buffer = false
	
	'salva la lista in un file di testo
	'dim Fso, TextFile, TextFilePath
	'Set Fso = CreateObject("Scripting.FileSystemObject")
	'TextFilePath = Application("IMAGE_PATH") & "\pagina_" & request.form("id_pag") & "_UltimoPost.txt"
	'Set TextFile = fso.CreateTextFile(TextFilePath, True, true)
	'TextFile.Write(request.form("layers"))
	'TextFile.close
	'set TextFile = nothing
	'set Fso = nothing
	
	On Error Resume Next		'per qualsiasi errore di parsing spedisco una mail
	
	dim p
	set p = new Parser
	if NOT p.ParseAndSave(request.form("layers")) OR Err.number > 0 then		'se errore
		CALL SendEmailSupportEX("ERRORE: compilazione pagina ("& Request.ServerVariables("SERVER_NAME") &")", request.form("layers"))
	end if



Class Parser

    Private conn		'connessione al dbLayer
    Private sql
    Private rs
    Private rsAux
    Private pagID		'ID della pagina da modificare
    
    Private webID		    'ID del sito in relazione con la pagina
    Private webAccessibile  'sito web accessibile
    
    Private css			'classe CssManager
    
    Private HtmlAltDefault  'valore di default dell'attributo html alt
    
    'sottostringhe di match
    Private regExLayer
    Private regExProp
    Private regExPropName
    Private regExPropValue
    Private regExPropFormatValue
    Private regExArrayOfArray
    Private regExArray
    Private regExStr
    Private regExNum
    Private regExCost
    Private regExTipi
    
    Private testo		'rs("testo") modificato del layer corrente
    
    Private Sub Class_Initialize()
    	pagID = CIntero(request.form("id_pag"))
    	
    	if pagID > 0 then
    		set conn = server.createobject("ADODB.connection")
    		set rs = server.createobject("ADODB.recordset")
    		set rsAux = server.createobject("ADODB.recordset")
    		conn.open Application("DATA_ConnectionString")
    		
            'recupera dati del sito
            sql = "SELECT tb_webs.* FROM tb_webs INNER JOIN tb_pages ON tb_webs.id_webs = tb_pages.id_webs WHERE id_page=" & pagID
            rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
            webID = rs("id_webs")
            webAccessibile = rs("sito_accessibile")
            if webAccessibile then
                HtmlAltDefault = """ """
            else
                HtmlAltDefault = """"""
            end if
            rs.close
            
            set css = New CssManager
    		
    		'inizializzo le stringhe di parsing
    		regExStr = "(?:""[^""]*"")"				'match stringhe
    		regExNum = "(?:\d+)"					'match numeri
    		regExCost = "(?:#[\w]+|<Void>)"			'match costanti: #alfanumerici o <Void>
    												'match array
    		regExArray = "(?:\[(?:"& regExStr &"(?:, )?|"& regExNum &"(?:, )?|"& regExCost &"(?:, )?)*\])"
    												'match array di array
    		regExArrayOfArray = "(?:\[(?:"& regExStr &"(?:, )?|"& regExNum &"(?:, )?|"& regExArray &"(?:, )?|"& regExCost &"(?:, )?)*\])"
    												'match del valore della proprieta "format"
    		regExPropFormatValue = "\[(?:"& regExArrayOfArray &"(?:, )?)*\]"
    												'match di un valore di una generica proprieta
    		regExPropValue = "("& regExStr &"|"& regExPropFormatValue &")"
    												'match del nome di una generica proprieta
    		regExPropName = "("& regExStr &": )"
    												'match di una generica proprieta
    		regExProp = "(?:"& regExPropName & regExPropValue &")"
    												'match di un layer
    		regExLayer = "(?:\[(?:"& regExProp &"(?:, )?)+\])"
    		
    												'match del singolo valore di array di array
    												'usato solo nel parsing del format e non all'interno di regExLayer
    		regExTipi = "(?:"& regExStr &"|"& regExNum &"|"& regExArray &"|"& regExCost &")"
    	end if
    End Sub
    
    Private Sub Class_Terminate()
    	if pagID > 0 then
    		rs.close
    		set rs = nothing
    		conn.close
    		set conn = nothing
    	end if
    End Sub
    
    'esegue il parse della stringa e salva nel DB
    'restituisce true se l'operazione e andata a buon fine
    Public Function ParseAndSave(txt)
    	if pagID > 0 then
    		dim re, layers, layer
    		set re = New RegExp
    		re.ignoreCase = true
    		re.global = true
    		re.multiline = false
    		
    		'parsing dei layers
    		re.pattern = regExLayer
    		set layers = re.Execute(txt)
    		
    		response.write txt
    		response.write "<br><br>"& regExLayer
    		response.write "<br><br>N.:"& layers.count
    		
    		conn.BeginTrans
    		UpdatePage
    		
    		'GESTIONE LAYERS
            sql = "SELECT * FROM tb_layers WHERE id_pag=" & pagId
    		rs.open sql, conn, adOpenKeySet, adLockOptimistic
    		for each layer in layers
    			'GESTIONE SINGOLO LAYER
    			ParseAndSaveLayer(layer.value)
    		next
    		rs.close
    		conn.CommitTrans
    		ParseAndSave = true
    		
    		response.write "<br><br>OK-----------------------------------------OK------------------------------------------OK"
    	else
    		ParseAndSave = false
    	end if
    End Function
    
    'esegue il parse e il salvataggio del singolo layer
    'layer:			txt del singolo layer
    'restituisce true se l'operazione e andata a buon fine
    Public Sub ParseAndSaveLayer(layer)
    	dim re, props, i, propName, propValue, html
    	set re = New RegExp
    	re.ignoreCase = true
    	re.global = true
    	re.multiline = false
    	
    	'parsing delle proprieta
    	re.pattern = regExProp
    	set props = re.Execute(layer)
    	
    	response.write "<br><br><br>LAY:"& layer
    	
    	rs.AddNew
    	for i = 0 to props.count - 1
    		propName = props(i).subMatches(0)
    		'tolgo gli apici e i ":" dal nome
    		propName = Left(Right(propName, Len(propName)-1), Len(propName)-4)
    		'tolgo gli a capo dal valore
    		propValue = Replace(props(i).subMatches(1), "\n", vbCrLf)
    		if Left(propValue, 1) = """" then			'se stringa tolgo gli apici
    			propValue = TogliApici(propValue)
    		end if
    		
    		response.write "<br><br>PROP:"& propName &" = "& propValue
    		if propName <> "id_lay" then
    			rs(propName) = propValue
    		end if
    	next
    	'calcolo i campi in em
    	rs("em_x") = PxToEm(rs("x"))
    	rs("em_y") = PxToEm(rs("y"))
    	rs("em_largo") = PxToEm(rs("largo"))
    	rs("em_alto") = PxToEm(rs("alto"))
    	
    	'Tolgo i vbCrLf dall'RTF causa errore nel parsing xml in editor (makelist)
    	'rs("rtf") = Replace(rs("rtf"), "}"& vbCrLf &"{", "}{")
    	'rs("rtf") = Replace(rs("rtf"), vbCrLf &"\", "\")
    	
    	'gestione salvataggio campi aggiuntivi plugin
    	if rs("id_tipo") = 4 then
            sql = "SELECT id_objects, identif_objects FROM tb_objects WHERE (tb_objects.name_objects LIKE '" + rs("nome") + "') AND id_webs=" & webID
    		rsAux.open sql, conn
    		rs("id_objects") = rsAux("id_objects")
    		rs("aspcode") = rsAux("identif_objects")
    		rsAux.close
    	end if
    	rs.Update
    	
    	'CREAZIONE HTML
    	testo = Replace(CString(rs("testo")), vbCrLf, vbCr)
    	SELECT CASE rs("id_tipo")
    		CASE 1					'testo normale
    			html = HTMLtxtNormal
    		CASE 2					'immagini
    			html = HTMLimg
    		CASE 3					'flash
    			html = HTMLflashJS
    		CASE 4					'plug-in
    			html = HTMLplugIn
    		CASE 5					'testo strutturato
    			html = HTMLtxtStructured
    	END SELECT
    	rsAux.open "SELECT * FROM tb_layers WHERE id_lay = "& rs("id_lay"), conn, adOpenDynamic, adLockOptimistic
    	rsAux("html") = CString(html)
    	rsAux.Update
    	rsAux.Close
    End Sub
    
    'toglie gli apici o le parantesi (i due caratteri esterni) da una stringa
    Private Function TogliApici(str)
    	TogliApici = Left(Right(str, Len(str)-1), Len(str)-2)
    End Function
    
    'crea il DIV del layer principale
    Private Function HTMLdiv(classe, id, html, chiudi)
    	HTMLdiv = "<div class="""& classe &""" id="""& id &"""" & _
    			  " style=""position:absolute; left:" & rs("x") & "px;" & _
    								 	     " top:" & rs("y") & "px;" & _
    										 " width:" & rs("largo") & +"px;" & _
    										 " height:" & rs("alto") & "px;" & _
    										 " z-index:" & rs("z_order") & """" &_
    			  ">" & vbCrLf & _
    			  html
    	if chiudi then
    		HTMLdiv = HTMLdiv &"</div>"& vbCrLf
    	end if
    End Function
    
    'formatta correttamente il testo preso dal recordset
    Private Function HTMLtesto(testo, aCapo)
    	HTMLtesto = Replace(Replace(Server.HtmlEncode(testo), vbTab, " &nbsp; &nbsp;"), "  ", " &nbsp;")
    	if aCapo <> vbCr then
    		HTMLtesto = Replace(HTMLtesto, vbCr, aCapo)
    	end if
    End Function
    
    'crea l'HTML per lo style
    Private Function HTMLstyle(font, fontSize, fontColor, letterSpacing)
    	if font <> "" then				'se font <> "" anche gli altri lo sono, puo essere vuoto per i link semplici
    		HTMLstyle = "style=""" & _
    				    "font-family: "& css.FontFamily_EditorXML_To_Css(TogliApici(font)) &"; " & _
    					"font-size: "& css.FontSize_EditorXML_To_Css(fontSize) & css.FONT_SIZE_unita_css &"; " & _
    					"color: "& TogliApici(fontColor) &";"
    		if letterSpacing <> CString(css.LETTER_SPACING_DEFAULT) then
    			HTMLstyle = HTMLstyle &" letter-spacing: "& css.LetterSpacing_EditorXML_To_Css(letterSpacing) & _
    													   css.LETTER_SPACING_unita_css &";"
    		end if
    		HTMLstyle = HTMLstyle &""""
    	end if
    End Function
    
    'crea l'HTML per il bold e l'italic
    Private Function HTMLbEi(stile, html)
    	HTMLbEi = html
    	if InStr(stile, "#bold") > 0 then
    		if html = "" then
    			HTMLbEi = " font-weight: bold;"
    		else
    			HTMLbEi = "<strong>"& HTMLbEi &"</strong>"
    		end if
    	end if
    	if InStr(stile, "#italic") > 0 then
    		if html = "" then
    			HTMLbEi = " font-style: italic;"
    		else
    			HTMLbEi = "<em>"& HTMLbEi &"</em>"
    		end if
    	end if
    End Function
    
    'crea l'HTML per il br
    Private Function HTMLbr(font, stile, fontSize, fontColor, letterSpacing)
    	dim aux
    	
    	HTMLbr = "<br"
    	
    	'inserisco lo style
    	aux = HTMLstyle(font, fontSize, fontColor, letterSpacing)
    	if aux <> "" then
    		HTMLbr = HTMLbr &" "& aux
    	end if
    	
    	'inserisco weight e style
    	aux = HTMLbEi(stile, "")
    	if aux <> "" then
    		if Right(HTMLbr, 1) = "r" then		'se non ho inserito lo style
    			HTMLbr = HTMLbr &" style="""& aux &""""
    		else
    			HTMLbr = Left(HTMLbr, Len(HTMLbr)-1) & aux &""""
    		end if
    	end if
    	
    	HTMLbr = HTMLbr &" />"
    End Function
    
    'crea l'HTML per lo span
    Private Function HTMLspan(font, stile, fontSize, fontColor, letterSpacing, html)
    	HTMLspan = "<span "& HTMLstyle(font, fontSize, fontColor, letterSpacing) &">"& html &"</span>"
    	
    	'inserisco lo strong e l'em
    	HTMLspan = HTMLbEi(stile, HTMLspan)
    End Function
    
    'crea l'HTML per il tag p
    Private Function HTMLp(align, lineHeight, html)
    	HTMLp = vbTab &"<p"
    	
    	'inserisco l'align e il line-height
    	align = LCase(Right(align, len(align)-1))
    	if align <> css.TEXT_ALIGN_DEFAULT OR lineHeight <> CString(css.LINE_HEIGHT_DEFAULT) then
    		HTMLp = HTMLp &" style="""
    		if align <> css.TEXT_ALIGN_DEFAULT then
    			HTMLp = HTMLp &" text-align: "& css.TextAlign_EditorXML_To_Css(align) &";"
    		end if
    		if lineHeight <> CString(css.LINE_HEIGHT_DEFAULT) then
    			HTMLp = HTMLp & css.LineHeight_Css(lineHeight)
    		end if
    		HTMLp = HTMLp &""""
    	end if
    	
    	HTMLp = HTMLp &">"& html &"</p>" & vbCrLf
    End Function
    
    'crea l'HTML per un tag strutturato
    Private Function HTMLtag(tag, html)
    	HTMLtag = vbTab &"<"& LCase(TogliApici(tag)) &">"& html &"</"& LCase(TogliApici(tag)) &">"& vbCrLf
    End Function
    
    'crea l'HTML per l'anchor
    Private Function HTMLa(href, title, target, font, stile, fontSize, fontColor, letterSpacing, html)
    	dim aux
    	
    	HTMLa = "<a href="& Replace(href, """default.aspx", """"& tagPath) & _
    			IIF(CString(title) <> "" AND title <> "qui metti il titolo", " title="& title, "")
    	
    	'gestione target
    	if target <> "[0]" then
    		dim targetParams
    		targetParams = Split(TogliApici(target), ", ")
    		HTMLa = HTMLa &" onclick=""javascript:OpenNewWindow(this, "& TogliApici(targetParams(1)) &", "& TogliApici(targetParams(2)) &")"""
    	end if
    	
    	'inserisco gli stili
    	aux = HTMLstyle(font, fontSize, fontColor, letterSpacing)
    	if aux <> "" then
    		HTMLa = HTMLa &" "& aux
    	end if
    	
    	HTMLa = HTMLa &">"& html &"</a>"
    	
    	'inserisco lo strong e l'em
    	HTMLa = HTMLbEi(stile, HTMLa)
    End Function
    
    'crea l'HTML per il testo normale
    'N.B.: dato che l'editor considera il vbCr come uno "" lo sostituisco con <br/> che non ha spessore e prima di un </tag>
    'viene ignorato. per compatibilità con editor se ho un elemento che contiene solo " " lo sostituisco con "&nbsp;"
    Private Function HTMLtxtNormal()
    	if testo <> "" then
    		dim re, props, prop, vals, val, i, k, html, txtDa, txtLun, isFormat
    		Redim parags(3, 0), formats(9, 0)
    		const x = 0
    		const y = 1
    		set re = New RegExp
    		re.ignoreCase = true
    		re.global = true
    		re.multiline = false
    		
    		'parsing delle formattazioni all'interno del format
    		re.pattern = regExArrayOfArray
    		set props = re.Execute(TogliApici(rs("format")))
    		
    		'inizializzazioni strutture dati
    		re.pattern = regExTipi
    		for each prop in props
    			set vals = re.Execute(TogliApici(prop.value))
    			SELECT CASE vals.count
    				CASE 4			'paragrafo
    					'al posto del secondo parametro "p" calcolo l'y
    					i = UBound(parags, 2)+1
    					Redim Preserve parags(3, i)
    					for k = 2 to vals.count-1
    						parags(k, i-1) = vals(k)
    					next
    					parags(x, i-1) = CInt(vals(0))
    					if i > 1 then
    						'calcolo l'y del precedente in base all'x dell'attuale
    						parags(y, i-2) = CInt(vals(0)-1)
    					end if
    				CASE ELSE		'link e span
    					i = UBound(formats, 2)+1
    					Redim Preserve formats(9, i)
    					for k = 0 to vals.count-1
    						formats(k, i-1) = vals(k)
    					next
    					formats(x, i-1) = Cint(formats(x, i-1))
    					formats(y, i-1) = Cint(formats(y, i-1))
    			END SELECT
    			'calcolo la y per l'ultimo paragrafo
    			parags(y, UBound(parags, 2)-1) = Len(testo)
    		next
    		
    		for i = 0 to UBound(parags, 2)-1					'livello paragrafi
    			response.write "<br><br>PARAG(x, y, align, lineH) = ("& parags(x, i) &", "& parags(y, i) &", "& parags(2, i) &", "& parags(3, i) &")"
    			html = ""
    			isFormat = false
    			
    			for k = 0 to UBound(formats, 2)-1			'livello anchor e span
    				response.write "<br><br>ASPAN = ("& formats(x, k) &", "& formats(y, k) &", "& formats(2, k) &", " & _
    													formats(3, k) &", "& formats(4, k) &", "& formats(5, k) &", " & _
    													formats(6, k) &", "& formats(7, k) &", "& formats(8, k) &", " & _
    													formats(9, k) &")"
    				
    				'se ho un'intersezione col paragrafo
    				if formats(x, k) <= parags(y, i) AND formats(y, k) >= parags(x, i) then
    					isFormat = true
    					
    					'aggiungo l'eventuale testo precedente non formattato o linkato
    					if k = 0 then
    						if formats(x, k) > parags(x, i) then
    							html = html & HTMLtesto(Mid(testo, parags(x, i), formats(x, k) - parags(x, i)), "<br />")
    						end if
    					else
    						if formats(y, k-1) >= parags(x, i) then
    							txtDa = formats(y, k-1)+1
    							txtLun = formats(x, k) - 1 - formats(y, k-1)
    						else
    							txtDa = parags(x, i)
    							if formats(x, k) >= txtDa then
    								txtLun = formats(x, k) - txtDa
    							else
    								txtLun = 0
    							end if
    						end if
    						html = html & HTMLtesto(Mid(testo, txtDa, txtLun), "<br />")
    					end if
    					
    					'controllo accavallamento con paragrafo successivo
    					if formats(x, k) >= parags(x, i) then							'se sono interno al paragrafo
    						txtDa = formats(x, k)								'indice del primo carattere da formattare
    					else																'mi accavallo a due paragrafi
    						txtDa = parags(x, i)
    					end if
    					if formats(y, k) <= parags(y, i) then							'se sono interno al paragrafo
    						txtLun = formats(y, k) - txtDa + 1					'lunghezza del testo da formattare
    					else																'mi accavallo a due paragrafi
    						txtLun = parags(y, i) - txtDa + 1
    					end if
    					
    					'formattazione testo
    					if formats(9, k) = "" then			'se sono uno span
    						if HTMLtesto(Mid(testo, txtDa, txtLun), "") = " " then
    							'se sono uno spazio lo sostituisco con &nbsp; per l'editor
    							html = html & HTMLspan(formats(2, k), formats(3, k), formats(4, k), formats(5, k), formats(6, k), _
    												   "&nbsp;")
    						elseif Mid(testo, txtDa, txtLun) = vbCr then
    							response.write "<br>DA:"& txtDa &";LUN:"& txtLun &";"
    							html = html & HTMLspan(formats(2, k), formats(3, k), formats(4, k), formats(5, k), formats(6, k), _
    												   "<br />")
    						else
    							html = html & HTMLspan(formats(2, k), formats(3, k), formats(4, k), formats(5, k), formats(6, k), _
    												   HTMLtesto(Mid(testo, txtDa, txtLun), ""))
    						end if
    					else								'sono un anchor
    						html = html & HTMLa(formats(2, k), formats(8, k), formats(3, k), formats(4, k), _
    									 		formats(5, k), formats(6, k), formats(7, k), formats(9, k), _
    									 		HTMLtesto(Mid(testo, txtDa, txtLun), "<br />"))
    					end if
    				elseif formats(x, k) > parags(y, i) then
    					exit for
    				end if
    			next
    			'aggiungo l'eventuale testo non formattato mancante
    			if isFormat then
    				if formats(x, k-1) < parags(y, i) AND formats(y, k-1) < parags(y, i) then
    					html = html & HTMLtesto(Mid(testo, formats(y, k-1) + 1, parags(y, i) - formats(y, k-1)), "")
    				end if
    			else			'non ho testo formattato in questo paragrafo
    				html = html & HTMLtesto(Mid(testo, parags(x, i), parags(y, i) - parags(x, i) + 1), "")
    			end if
    			
    			if html = "" then			'se sono una riga vuota
    				html = "<br />"
    			elseif html = " " then		'se sono uno spazio vuoto
    				html = "&nbsp;"
    			end if
    			
    			HTMLtxtNormal = HTMLtxtNormal & HTMLp(parags(2, i), parags(3, i), html)
    		next
    	end if
    End Function
    
    'crea l'HTML per il testo strutturato
    'N.B.: dato che l'editor considera il vbCr come uno "" lo sostituisco con <br/> che non ha spessore e prima di un </tag>
    'viene ignorato. per compatibilita con editor se ho un elemento che contiene solo " " lo sostituisco con "&nbsp;"
    Private Function HTMLtxtStructured()
    	if testo <> "" then
    		dim re, props, prop, vals, val, i, k, html, txtDa, txtLun, isFormat
    		Redim parags(2, 0), formats(5, 0)
    		const x = 0
    		const y = 1
    		set re = New RegExp
    		re.ignoreCase = true
    		re.global = true
    		re.multiline = false
    		
    		'parsing delle formattazioni all'interno del format
    		re.pattern = regExArrayOfArray
    		set props = re.Execute(TogliApici(rs("format")))
    		
    		'inizializzazioni strutture dati
    		re.pattern = regExTipi
    		for each prop in props
    			set vals = re.Execute(TogliApici(prop.value))
    			if vals.count = 3 AND Left(vals(2), 1) = """" then			'paragrafo
    				i = UBound(parags, 2)+1
    				Redim Preserve parags(2, i)
    				for k = 0 to vals.count-1
    					parags(k, i-1) = vals(k)
    				next
    				parags(x, i-1) = Cint(parags(x, i-1))
    				parags(y, i-1) = Cint(parags(y, i-1))
    			else														'link e stili
    				i = UBound(formats, 2)+1
    				Redim Preserve formats(5, i)
    				for k = 0 to vals.count-1
    					formats(k, i-1) = vals(k)
    				next
    				formats(x, i-1) = Cint(formats(x, i-1))
    				formats(y, i-1) = Cint(formats(y, i-1))
    			end if
    		next
    		
    		for i = 0 to UBound(parags, 2)-1					'livello paragrafi
    			response.write "<br><br>PARAG(x, y, tag) = ("& parags(x, i) &", "& parags(y, i) &", "& parags(2, i) &")"
    			html = ""
    			isFormat = false
    			
    			for k = 0 to UBound(formats, 2)-1			'livello anchor e stili
    				response.write "<br><br>ASTILE = ("& formats(x, k) &", "& formats(y, k) &", "& formats(2, k) &", " & _
    												  	 formats(3, k) &", "& formats(4, k) &", "& formats(5, k) &")"
    				
    				'se ho un'intersezione col paragrafo
    				if formats(x, k) <= parags(y, i) AND formats(y, k) >= parags(x, i) then
    					isFormat = true
    					
    					'aggiungo l'eventuale testo precedente non formattato o linkato
    					if k = 0 then
    						if formats(x, k) > parags(x, i) then
    							html = html & HTMLtesto(Mid(testo, parags(x, i), formats(x, k) - parags(x, i)), "<br />")
    						end if
    					else
    						if formats(y, k-1) >= parags(x, i) then
    							txtDa = formats(y, k-1)+1
    							txtLun = formats(x, k) - 1 - formats(y, k-1)
    						else
    							txtDa = parags(x, i)
    							if formats(x, k) >= txtDa then
    								txtLun = formats(x, k) - txtDa
    							else
    								txtLun = 0
    							end if
    						end if
    						html = html & HTMLtesto(Mid(testo, txtDa, txtLun), "<br />")
    					end if
    					
    					'controllo accavallamento con paragrafo successivo
    					if formats(x, k) >= parags(x, i) then							'se sono interno al paragrafo
    						txtDa = formats(x, k)								'indice del primo carattere da formattare
    					else																'mi accavallo a due paragrafi
    						txtDa = parags(x, i)
    					end if
    					if formats(y, k) <= parags(y, i) then							'se sono interno al paragrafo
    						txtLun = formats(y, k) - txtDa + 1					'lunghezza del testo da formattare
    					else																'mi accavallo a due paragrafi
    						txtLun = parags(y, i) - txtDa + 1
    					end if
    					
    					'formattazione testo
    					if formats(5, k) = "" then			'se sono stile
    						if HTMLtesto(Mid(testo, txtDa, txtLun), "") = " " then
    							html = html & HTMLbEi(formats(2, k), "&nbsp;")
    						elseif Mid(testo, txtDa, txtLun) <> vbCr then		'ignoro il vbCr perche non cambia la visualizzazione
    							html = html & HTMLbEi(formats(2, k), HTMLtesto(Mid(testo, txtDa, txtLun), ""))
    						end if
    					else								'sono un anchor
    						html = html & HTMLa(formats(2, k), formats(5, k), formats(3, k), "", formats(4, k), "", "", "", _
    									 		HTMLtesto(Mid(testo, txtDa, txtLun), ""))
    					end if
    				elseif formats(x, k) > parags(y, i) then
    					exit for
    				end if
    			next
    			'aggiungo l'eventuale testo non formattato mancante
    			if isFormat then
    				if formats(x, k-1) < parags(y, i) AND formats(y, k-1) < parags(y, i) then
    					html = html & HTMLtesto(Mid(testo, formats(y, k-1) + 1, parags(y, i) - formats(y, k-1)), "")
    				end if
    			else			'non ho testo formattato in questo paragrafo
    				html = html & HTMLtesto(Mid(testo, parags(x, i), parags(y, i) - parags(x, i) + 1), "")
    			end if
    			
    			if html = "" then			'se sono una riga vuota
    				html = "<br />"
    			elseif html = " " then		'se sono uno spazio vuoto
    				html = "&nbsp;"
    			end if
    			
    			HTMLtxtStructured = HTMLtxtStructured & HTMLtag(parags(2, i), html)
    		next
    	end if
    End Function
    
    'crea l'HTML dell'immagine
    Private Function HTMLimg()
    	dim re, props
    	set re = New RegExp
    	re.ignoreCase = true
    	re.global = true
    	re.multiline = false
    	
    	'parsing delle proprieta all'interno dell'unico array contenuto nel format per le immagini
    	re.pattern = regExTipi
    	set props = re.Execute(Left(Right(rs("format"), Len(rs("format"))-2), Len(rs("format"))-4))
    	
    	if props.count > 4 then
    		dim alt
    		alt = props(4)
    		if CString(alt) = "" OR alt = """testo alternativo""" then
                if webAccessibile then
        			alt = props(5)
                else
                    alt = HtmlAltDefault
                end if
    		end if
    		if CString(alt) = "" OR alt = """qui metti il titolo""" then
                alt = HtmlAltDefault
    		end if
    		'compone link su immagine
    		HTMLimg = HTMLa(props(2), props(5), props(3), "", "", "", "", "", _
    						"<img src="""& tagPathResources &"/images/"& rs("nome") &""" alt="& alt &" style=""width:"& rs("em_largo") &"em; height:"& rs("em_alto") &"em;"" />")
    	else
    		HTMLimg = "<img src="""& tagPathResources &"/images/"& rs("nome") &""" alt="& IIF(CString(props(2)) = "" OR props(2) = """testo alternativo""", HtmlAltDefault, props(2)) &" style=""width:"& rs("em_largo") &"em; height:"& rs("em_alto") &"em;"" />"
    	end if
    	HTMLimg = vbTab & HTMLimg
    End Function
    
    'crea l'HTML dei plug-in
    Private Function HTMLplugIn()
    	HTMLplugIn = " "
    	'segue in dynalay...
    End Function
    
    'crea l'HTML dei flash
    Private Function HTMLflash()
    	dim params, param
    	set params = server.CreateObject("Scripting.Dictionary")
    	
    	if testo <> "" then
    		'creazione dictionary di parametri
    		for each param in Split(Left(testo, Len(testo)-1), "; ")
    			params.Add Left(param, InStr(param, "=")-1), Right(param, Len(param) - InStr(param, "="))
    		next
    	end if
    	
    	HTMLflash = HTMLflash &"	<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" " & _
    				         		"codebase=""https://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"" " & _
    								"style=""width:"& rs("em_largo") &"em; height:"& rs("em_alto") &"em;"" " & _
    								"id=""flash_name_"& rs("id_lay") &""">" & vbCrLf
    	
    	'aggiungo i PARAM
    	HTMLflash = HTMLflash &"		<param name=""movie"" value="""& tagPathResources &"/flash/"& rs("nome") &""" />" & vbCrLf
    	for each param in params
    		HTMLflash = HTMLflash & "		<param name=""" & param & """ value=""" & params(param) & """ />" & vbCrLf
    	next
    
    	HTMLflash = HTMLflash &"	</object>"
    End Function
    
    Private Function HTMLflashJS()
    	dim params, param
    	set params = server.CreateObject("Scripting.Dictionary")
    	
    	if testo <> "" then
    		'creazione dictionary di parametri
    		for each param in Split(Left(testo, Len(testo)-1), "; ")
    			params.Add Left(param, InStr(param, "=")-1), Right(param, Len(param) - InStr(param, "="))
    		next
    	end if
		
    	HTMLflashJS = "	<script type=""text/javascript"">" & vbCrLf & _
    			      "	<!--" & vbCrLf & _
    			      "	AC_FL_RunContent(" & vbCrLf & _
    				  "		'codebase', 'http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0'," & vbCrLf & _
    				  "		'width', '100%'," & vbCrLf & _
    				  "		'height', '100%'," & vbCrLf & _
    				  "		'src', '"& Replace(rs("nome"), ".swf", "", 1, -1, vbTextCompare) &"'," & vbCrLf & _
    				  "		'quality', '"& IIF(params("quality") <> "", params("quality"), "high") &"'," & vbCrLf & _
    				  "		'pluginspage', 'https://www.macromedia.com/go/getflashplayer'," & vbCrLf & _
    				  "		'align', '"& IIF(params("align") <> "", params("align"), "middle") &"'," & vbCrLf & _
    				  "		'play', '"& IIF(params("play") <> "", params("play"), "true") &"'," & vbCrLf & _
    				  "		'loop', '"& IIF(params("loop") <> "", params("loop"), "true") &"'," & vbCrLf & _
    				  "		'scale', '"& IIF(params("scale") <> "", params("scale"), "showall") &"'," & vbCrLf & _
    				  "		'wmode', '"& IIF(params("wmode") <> "", params("wmode"), "transparent") &"'," & vbCrLf & _
    				  "		'devicefont', '"& IIF(params("devicefont") <> "", params("devicefont"), "false") &"'," & vbCrLf & _
    				  "		'id', 'flash_name_"& rs("id_lay") &"'," & vbCrLf & _
    				  "		'bgcolor', '"& IIF(params("bgcolor") <> "", params("bgcolor"), "#FFFFFF") &"'," & vbCrLf & _
    				  "		'name', 'flash_name_"& rs("id_lay") &"'," & vbCrLf & _
    				  "		'menu', '"& IIF(params("menu") <> "", params("menu"), "true") &"'," & vbCrLf & _
    				  "		'allowScriptAccess','sameDomain'," & vbCrLf & _
    				  "		'movie', '"& tagPathResources &"/flash/"& Replace(rs("nome"), ".swf", "", 1, -1, vbTextCompare) &"'," & vbCrLf & _
    				  "		'salign', '"& IIF(params("salign") <> "", params("salign"), "") &"'" & vbCrLf & _
    				  "	); //end AC code" & vbCrLf & _
    				  "	// -->" & vbCrLf & _
    				  "	</script>" & vbCrLf & _
    				  "	<noscript>" & vbCrLf & _
    				  "		<div>" & vbCrLf & _
    				  "			<a href=""https://www.macromedia.com/go/getflash/"">Flash Player Download</a>" & vbCrLf & _
    				  "		</div>" & vbCrLf & _
    				  "	</noscript>"
    End Function
    
    'salva i colori, update delle date nella paginasito, cancellazione layer per reinserimento
    Private Sub UpdatePage()
    	sql = " UPDATE tb_pages SET SfondoColore='"& request.form("sfondoColore") &"', SfondoImmagine='" & request.form("Sfondo") &"' "& _
    		  " WHERE id_page = "& pagID
    	CALL conn.execute(sql, , adExecuteNoRecords)
    	
      	'aggiorna data di modifica paginasito
    	CALL UpdateDataModifica(conn, pagID)
    	
        sql = "DELETE FROM tb_layers WHERE id_pag="& pagID
    	CALL conn.execute(sql, , adExecuteNoRecords)
    End Sub

End Class
%>