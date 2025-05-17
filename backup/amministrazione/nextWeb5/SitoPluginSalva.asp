<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="Tools_NEXTweb5.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_plugin_accesso, 0))

dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	Classe.Requested_Fields_List	= "tft_identif_objects; tft_name_objects; tft_param_list"
	Classe.Checkbox_Fields_List 	= ""
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_objects"
	Classe.id_Field					= "id_objects"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
	Classe.SetUpdateParams("obj_")

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	dim sql
	
	'verifica del nome plugin
	if CheckChar(request("tft_name_objects"), LOGIN_VALID_CHARSET) then
		
		sql = " SELECT COUNT(*) FROM tb_objects " + _
			  " WHERE id_webs="& Session("AZ_ID") &" AND name_objects LIKE '" & request("tft_name_objects") & "' AND id_objects<>" & ID
		if cInteger(GetValueList(conn, rs, sql))=0 then
		
			dim Changed, Parser
			Changed = (request("tft_param_list") <> request("old_param_list"))
			
			'esegue l'aggiornamento alle proprieta' solo se il testo e' stato variato
			set Parser = new ClassProperties
			'carica le proprieta' della classe
			
			CALL Parser.ClassProperiesParse(request("tft_param_list"))
			
			'aggiorna le propriet&agrave; del plugin
			sql = "SELECT * FROM tb_objects WHERE id_objects=" & cInteger(ID)
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			rs("param_list") = Parser.ClassPropertyUpdate()
			rs.update
			
			CALL UpdateSitoDataModificaPlugin(conn, rs("id_webs"))
			
			rs.close
			
			sql = "SELECT id_lay, testo, aspcode, nome FROM tb_layers WHERE id_objects=" & ID
			rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
			if not rs.eof then
				while not rs.eof
					if Changed then
						'carica le proprieta' dell'istanza e le aggiorna
						rs("testo") = Parser.InstancePropertiesUpdate(rs("testo"))
					end if
					rs("aspcode") = request("tft_identif_objects")
					rs("nome") = request("tft_name_objects")
					rs.update
					rs.movenext
				wend
			end if
						
			rs.close
			
			
			set Parser = nothing
		else
			Session("ERRORE") = "Nome gi&agrave; utilizzato da un'altra classe!"
		end if
		
	else
		Session("ERRORE") = "Nome classe non valido!"
	end if
	
	if session("ERRORE") = "" then
		'imposta parametri per passare alla pagina successiva
		Classe.isReport = FALSE
		Classe.Next_Page = "SitoPlugin.asp"
	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()



'*******************************************************************************
'classe per la gestione e modifica delle propriet&agrave; dei plugin
'*******************************************************************************
class ClassProperties
	'DICHIARAZIONE VARIABILI:*********************
	private ClassProperties
	private InstanceProperties

	'DEFINIZIONE METODI DEFAULT:*********************
	Private Sub Class_Initialize()
		
	end sub
	
	Private Sub Class_Terminate()
		
	End Sub
	
	'DEFINIZIONE METODI PUBBLICI:*********************
	
	'carica le proprieta' della classe nella variabile globale
	public sub ClassProperiesParse(properties)
		set ClassProperties = PropertiesParse(properties)
	end sub
	
	public function ClassPropertyUpdate()
		ClassPropertyUpdate = PropertiesList(ClassProperties)
	end function
	
	'carica, esegue il parsing ed aggiorna le proprieta' dell'istanza
	public function InstancePropertiesUpdate(properties)
		
		set InstanceProperties = PropertiesParse(properties)
		
		if ClassProperties.recordcount > 0 then
			ClassProperties.movefirst
		end if
		if InstanceProperties.recordcount > 0 then
			InstanceProperties.movefirst
		end if
		
		while not ClassProperties.eof
			if InstanceProperties.recordcount > 0 then
				InstanceProperties.movefirst
			end if
			InstanceProperties.find "prop_name='" & ClassProperties("prop_name") & "'"
			if InstanceProperties.eof then
				'proprieta' non presente: la aggiunge
				CALL PropertyAdd(InstanceProperties, ClassProperties("prop_name"), ClassProperties("prop_updatable"), ClassProperties("prop_value"))
			else
				'proprieta gi&agrave; presente
				if ClassProperties("prop_updatable") then
					'proprieta' modificabile
					if not InstanceProperties("prop_updatable") then
						'se non e' modificabile ne cambia lo stato impostando nuovamente il valore
						InstanceProperties("prop_updatable") = true 
						InstanceProperties("prop_value") = ClassProperties("prop_value") 
						InstanceProperties.update
					end if
				else
					'proprieta non modificabile
					if InstanceProperties("prop_updatable") then
						'proprieta' modificabile: ne cambia lo stato ed imposta il valore
						InstanceProperties("prop_updatable") = false 
						InstanceProperties("prop_value") = ClassProperties("prop_value") 
						InstanceProperties.update
					else
						'proprieta' non modificabile: ne aggiorna solo il valore
						InstanceProperties("prop_value") = ClassProperties("prop_value") 
						InstanceProperties.update
					end if
				end if
			end if
			ClassProperties.movenext
		wend
		
		'verifica che tutte le proprieta' dell'istanza siano ancora presenti nella dichiarazione della classe
		if InstanceProperties.recordcount > 0 then
			InstanceProperties.movefirst
			
			if ClassProperties.recordcount > 0 then
				'ci sono proprieta' nella dichiarazione della classe: esegue controllo
				while not InstanceProperties.eof
					ClassProperties.movefirst
					
					ClassProperties.find "prop_name='" & InstanceProperties("prop_name") & "'"
					if ClassProperties.eof then
						'proprieta' non trovata: la cancella dall'istanza
						InstanceProperties.Delete
						InstanceProperties.Update
					end if
					InstanceProperties.Movenext
				wend
			else
				'non ci sono proprieta' dichiarate nella clase: cancella tutto
				while not InstanceProperties.eof
					InstanceProperties.Delete
					InstanceProperties.Update
					InstanceProperties.Movenext
				wend
			end if
		end if
		
		InstancePropertiesUpdate = PropertiesList(InstanceProperties)
	end function
	
	
	'DEFINIZIONE METODI PRIVATI:*********************
	
	'carica le proprieta' restituendo un recordset che contiene in 3 colonne tutte le specifiche delle proprieta'
	private function PropertiesParse(properties)
		dim prs, PropList, Prop, PropParts
		dim PropUpdatable
		
		'crea contenitore
		set prs = Server.CreateObject("ADODB.recordset")
		
		'aggiunge colonne contenitore
		prs.Fields.Append "prop_name", adLongVarWChar, 1000
		prs.Fields.Append "prop_updatable", adBoolean
		prs.fields.Append "prop_value", adLongVarWChar, 8000
		
		prs.open
		'divide le proprieta' e le carica nel recordset
		PropList = split(properties, ";")
		for each Prop in PropList
			'verifica se la proprieta' e' valida
			if instr(1, prop, "=", vbTextCompare) then
				if instr(1, prop, ":=", vbTextCompare)>0 then
					'proprieta' modificabile dall'utente in editing
					PropUpdatable = true
					PropParts = split(prop, ":=", 2, vbTextCompare)
				else
					'proprieta' non modificabile dall'utente
					PropUpdatable = false
					PropParts = split(prop, "=", 2, vbTextCompare)
				end if
				'aggiunge la proprieta' con i dati recuperati
				CALL PropertyAdd( prs, Trim(PropParts(0)), PropUpdatable, Trim(PropParts(1)) )
			end if
		next
						
		set PropertiesParse = prs
	end function
	
	'aggiunge una proprieta' al recordset
	private sub PropertyAdd(byref prs, prop_name, prop_updatable, prop_value)
		prop_name = replace(prop_name, vbCrLF, "")
		prop_value = replace(prop_value, vbCrLF, "")
		
		'verifica se la propriet&agrave; c'&egrave; gia
		if prs.recordcount > 0 then
			prs.movefirst
			prs.find "prop_name LIKE '" & prop_name & "'"
		end if
		
		if prs.eof then
			'la proprieta' non c'e': la inserisce
			prs.AddNew
			prs("prop_name") = prop_name
			prs("prop_updatable") = prop_updatable
			prs("prop_value") = prop_value
			prs.update
			prs.movenext
		else
			'la proprieta' c'&egrave; pi&ugrave; di una volta: la aggiorna
			prs("prop_updatable") = prop_updatable
			prs("prop_value") = prop_value
			prs.update
			prs.movelast
		end if
	end sub

	
	'funzione che ritorna la stringa completa contentene tutte le propriet&agrave; del plugin
	private function PropertiesList(byRef prs)
		PropertiesList = ""
		if prs.recordcount>0 then
			prs.movefirst
			while not prs.eof
				PropertiesList = PropertiesList + PropertyToString(prs)
				prs.movenext
			wend
		end if
	end function
	
	'funzione che ritorna la stringa da salvare nel database corrispondente alla propriet&agrave; indicata
	private function PropertyToString(byref prs)
		PropertyToString = prs("prop_name")
		if prs("prop_updatable") then
			PropertyToString = PropertyToString + ":="
		else
			PropertyToString = PropertyToString + "="
		end if
		PropertyToString = PropertyToString & prs("prop_value") & ";" & VbCrLf
	end function
	
end class
%>