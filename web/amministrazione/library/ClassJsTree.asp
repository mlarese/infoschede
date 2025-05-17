<%
'variabile js per includere solo una volta per pagina i javascript anche se gli alberi per pagina sono piu' di uno
dim JsIncluded
JsIncluded = false

'*******************************************************************************************************************
'CLASSE CHE GENERA UN ALBERO IN JAVASCRIPT

class ObjJsTree

'******************************************************************************************************************************************
'******************************************************************************************************************************************
'VARIABILI E PROPRIETA' PUBBLICHE
'******************************************************************************************************************************************

    'proprieta' di impostazione nome (ID javascript) dell'albero
    public Property Get name()
        name = treeName
	End Property
	
	Public Property Let name(newName)
        
        'ripulisce vecchio nome
        if cString(treeName)<>"" then
            Session.contents.remove(treeName)
        end if
        
        treeName = newName
        
        'salva cookie in sessione per apertura dell'albero
        session(treeName) = DecodeExpansionCookie(request.cookies("clickedFolder" + treeName))
		
    End Property
    
    public LivelloBase
	
    'proprieta' di visualizzazione
    public TableCaption
    public Title
    public TitleLink
	Public LINK_ONCLICKNODE
	
    
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'VARIABILI E PROPRIETA' PRIVATE
'******************************************************************************************************************************************
    
    private treeName
    private nodeList
    
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'COSTRUTTORI CLASSE
'******************************************************************************************************************************************

    Private Sub Class_Initialize()
        
        'inizializza nome albero
        name = replace(GetPageName(), ".", "")
        
        'inizializza nodi e foglie
        nodeList = ""
        
        'inizializa variabili per HTML
        TableCaption = ""
        Title = ""
        TitleLink = ""
		LINK_ONCLICKNODE = true
        LivelloBase = 0
    End Sub
    
    
    Private Sub Class_Terminate()
        
    End Sub

'******************************************************************************************************************************************
'******************************************************************************************************************************************
'METODI PUBBLICI
'******************************************************************************************************************************************
    
    
	'verifica se il nodo indicato e' aparto o chiuso
    public function IsNodeExpanded(ID)		
        IsNodeExpanded = (InStr(1, session(treeName), ID &",", vbTextCompare) > 0)
    end function
    
	
	'aggiunge lo stato di espansione del nodo
	public function ExpandeNode(ID)
		if not IsNodeExpanded(ID) then
			session(treeName) = session(treeName) + ID & ","
			response.cookies("clickedFolder" + treeName) = EncodeExpansionCookie(session(treeName))
		end if
	end function
	
	
	'azzera lo stato di espansione di tutti i nodi (chiude l'albero)
	public sub ResetExpansionState()
		session(treeName) = ""
		response.cookies("clickedFolder" + treeName) = ""
		response.cookies("clickedFolder" + treeName).expires = DateAdd("d", -1, Now())
	end sub
	
    
    'aggiunge foglia all'albero
    public sub AddLeaf(level, label, title, url)
        if title<>"" then
            label = "<span title=\""" & replace(title, """", "\""") & "\"">" & label & "</span>"
        end if
        nodeList = nodeList + VBcRlF + _
                   "insDoc(" & IIF(level=cIntero(LivelloBase), "foldersTree", "aux" & (level-1)) & ", gLnk(""S"", """ & label &""", """ & url & """)) " + VBcRlF
    end sub
    
    
    'aggiunge ramo all'albero
    public sub AddNode(level, label, title, url, ID)
        if title<>"" then
            label = "<span title=\""" & replace(title, """", "\""") & "\"">" & label & "</span>"
        end if
        nodeList = nodeList + vbCrLF + _
                   " aux" & level & "= insFld(" & IIF(level=cIntero(LivelloBase), "foldersTree", "aux" & (level-1)) & ", gFld(""" & label & """, """ & url & """)) " + VBcRlf + _
                   " aux" & level & ".xID = " & ID & VBcRlf
    end sub
    
    
    'aggiunge ramo "nuovo" all'albero
    public sub AddNodeNew(livello, padre_id, Links, Labels)
        'se l'elemento non e' espanso mette solo una voce fittizia
        if IsNodeExpanded(padre_id) OR cIntero(padre_id)=0 then
            dim value, LinkList, LabelList, i
            
            LinkList = split(links, ";")
            LabelList = split(labels, ";")
            value = ""
            
            for i = lbound(LinkList) to uBound(LabelList)
                value = value + "<a href='" & LinkList(i) & IIF(cIntero(padre_id)>0, padre_id, "") & "'>" + LabelList(i) + "</a>"
                if i < uBound(LabelList) then
                    value = value + " - "
                end if
            next
            
            CALL AddLeaf(livello, "<span style='font-size:10px;'>" + value + "</span>", "", "")
        else
            CALL AddLeaf(livello, "", "", "")
        end if
    end sub
    
    
    'procedura che scrive la tabella che contiene la visualizzazione dell'albero.
    public sub Write() %>
        <table cellspacing="0" cellpadding="0" class="tabella_madre">
            <caption class="border"><%= TableCaption %></caption>
            <tr>
				<td id="td_<%= treeName %>">
                    <% if not JsIncluded then %>
                        <script src="<%= GetLibraryPath() %>Categorie/ua.js" language="JavaScript" type="text/javascript"></script>
	                    <script src="<%= GetLibraryPath() %>Categorie/ftiens4.js" language="JavaScript" type="text/javascript"></script>
                        <% JsIncluded = true
                    end if%>
                    
					<script language="JavaScript" type="text/javascript">
                        // CONFIGURAZIONE ALBERO
            	        USETEXTLINKS = 1;
            	        STARTALLOPEN = 0;
            	        USEFRAMES = 0;
            	        USEICONS = 0;
            	        WRAPTEXT = 1;
            	        PRESERVESTATE = 1;
            	        HIGHLIGHT = 0;
            	        ICONPATH = '<%= GetAmministrazionePath() %>grafica/filemanager/';
            	        LINK_ONCLICKNODE = <%= IIF(LINK_ONCLICKNODE, "true", "false") %>;
                        
            	        foldersTree = gFld("<strong><%= Title %></strong>", "<%= IIF(TitleLink="" AND Title<>"", "javascript:void(0);", TitleLink) %>");
            	        foldersTree.xID = 0;
                        
                        foldersTree.treeID = "<%= treeName %>";
                        
                        //nodi e foglie dell'albero javascript
                        <%= nodeList %>
                        
						initializeDocument()
					</script>
				</td>
			</tr>
		</table>
    <% end sub
    
    
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'METODI E FUNZIONI PRIVATI
'******************************************************************************************************************************************
	
	
	'procedura per la decodifica del cookie in lista di id
	private function DecodeExpansionCookie(value)
		DecodeExpansionCookie = replace(replace(value, " ", ""), "^", ",")
	end function
	
	
	'procedura che codifica la lista di id in lista per cookie
	private function EncodeExpansionCookie(value)
		EncodeExpansionCookie = replace(replace(value, " ", ""), ",", "^")
	end function

end class


%>