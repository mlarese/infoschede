<% 
'*******************************************************************************
'classe per il collegamento al Web Service
'*******************************************************************************

class WebService
	'variabili interne
	public Soap
	
	private nodeList
	Private nodePosition
	private pageS

	Private Sub Class_Initialize()
		pageS = 10
	End Sub
	
	Private Sub Class_Terminate()
		set soap = nothing
	End Sub

'DEFINIZIONE METODI DELLA CLASSE:*********************

Public Sub Open(WSDL_PATH)

	set soap = Server.CreateObject("MSSOAP.SOAPClient30")
	soap.ClientProperty("ServerHTTPRequest") = True
	CALL soap.mssoapinit(WSDL_PATH, "", "", "")
	soap.ConnectorProperty("Timeout") = 100000
	NodePosition = adPosUnknown

end sub

Public Sub Close()
	NodePosition = adPosUnknown
	Set NodeList = nothing
end Sub


'sposta il valore di una posizione
Public Function MoveNext()
	if NodePosition <= NodeList.length then
		NodePosition = NodePosition + 1
	else
		'se oltre la fine della lista genera un errore
		err.Raise 9		'subscript out of range
	end if
end function

'sposta il valore di una posizione indietro
Public Function MovePrevious()
	if NodePosition > 0 then
		NodePosition = NodePosition - 1
	else
		'se oltre l'inizio della lista genera un errore
		err.Raise 9		'subscript out of range
	end if
end function

'sposta il cursore all'inizio
Public Function MoveFirst()
	NodePosition = 0
end function

'sposta il cursore alla fine
Public Function MoveLast()
	NodePosition = NodeList.length-1
end function


'DEFINIZIONE PRORIETA'*******************
'restituisce la dimensione di paginazione
	Public Property Get PageSize
		pageSize = pageS
	end Property

'imposta la dimensione di paginazione
	Public Property Let PageSize(ps)
		if CIntero(ps) > 0 then
			pageS = CIntero(ps)
		end if
	end Property

'restituisce il numero delle pagine
	Public Property Get PageCount
		PageCount = (nodelist.length-1) \ pageS + 1
	end Property

'restituisce il numero di nodi presenti
	Public Property Get recordCount()
    	recordCount = NodeList.length
	End Property
	
'restituisce la posizione corrente
	Public Property Get AbsolutePosition
		AbsolutePosition = NodePosition + 1
	end Property

'imposta la posizione corrente
	Public Property Let AbsolutePosition(NewPosition)
		NodePosition = NewPosition - 1
	end Property

'restituisce la pagina corrente
	Public Property Get AbsolutePage
		AbsolutePage = NodePosition \ pageS + 1
	end Property

'imposta la posizione corrente
	Public Property Let AbsolutePage(NewPage)
		NodePosition = (newPage-1) * pageS
	end Property

'restituisce false se la posizione corrente è oltre l'ultimo nodo
	Public Property Get BOF()
    	BOF = EOF OR nodeList.length = 0
	End Property

'restituisce false se la posizione corrente è oltre l'ultimo nodo
	Public Property Get EOF()
    	EOF = (NodePosition = NodeList.length)
	End Property
	
'restituisce false se la posizione corrente è oltre l'ultimo nodo
	Public Property Get ToString()
    	ToString = NodeList.Item(NodePosition).xml
	End Property

'restituisce il valore della proprietà del nodo richiesto
	Public Default Property Get Field(Field_name)
		dim Element
		'recupera elemento richiesto
		Set Element = NodeList.Item(NodePosition).selectSingleNode("*[translate(local-name(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz') = '"& LCase(field_name) &"']")
		'restituisce il contenuto
		if instr(1,TypeName(Element), "Nothing", vbTextCompare)<1 then
			Field = Element.text
		else
			Field = ""
		end if
		set Element = nothing
	End Property


'FUNZIONI PRIVATE************************
'restituisce la lista di nodi (IXMLDOMSelection)
	Public Function getData(WS_Response, xpath)
		dim objTmpDom
		if not IsObject(WS_Response) then
			'riceve una stringa XML dal webservice
			set objTmpDom = Server.CreateObject("Msxml2.DOMDocument.4.0")
			objTmpDom.loadXML(WS_Response)
			Set NodeList = objTmpDom.selectNodes(xpath)
		else
			Set NodeList = WS_Response.item(1).selectNodes("//*[local-name()='" + xpath + "']")
		end if
		
		'imposta cursore all'inizio della lista
		NodePosition = 0
	end function
	
end class
%>