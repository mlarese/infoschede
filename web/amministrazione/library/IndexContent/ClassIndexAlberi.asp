<!--#INCLUDE FILE="../ClassJsTree.asp" -->  
<% 
'******************************************************************************************************************************************
'CLASSE:    ObjIndexLegend


'******************************************************************************************************************************************
'CLASSE :   ObjIndexTrees
'
'METODI:    AlberoIndiceCompleto()
'           AlberoIndiceByTipoContenuto()


'************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************
'CLASSE CHE GENERA ALBERI DALL'INDICE
'************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************


class ObjIndexTrees

'******************************************************************************************************************************************
'VARIABILI E PROPRIETA' PUBBLICHE
'******************************************************************************************************************************************
    'variabile indice sulla base della quale viene generato l'albero
    public Index
	
	'true per gestione esterna del nome e del link del nodo
	public GestioneEsterna
    
'******************************************************************************************************************************************
'VARIABILI E PROPRIETA' PRIVATE
'******************************************************************************************************************************************
    public tree

'******************************************************************************************************************************************
'COSTRUTTORI CLASSE
'******************************************************************************************************************************************
    
    
    Private Sub Class_Initialize()
        'inizializza oggetto gestione albero javascript
        set tree = new ObjJsTree
		GestioneEsterna = false
    End Sub
    
    Private Sub Class_Terminate()
    	set tree = nothing
    End Sub
    
	
'******************************************************************************************************************************************
'ALTRE FUNZIONI PUBBLICHE
'******************************************************************************************************************************************
    
	public sub ExpandNodes(IdxId)
		dim sql, id
		sql = "SELECT idx_tipologie_padre_lista FROM tb_contents_index WHERE idx_id=" & cIntero(IdxId)
		for each id in split(GetValueList(Index.conn, NULL, sql), ",")			
			if cIntero(id)>0 then
				tree.ExpandeNode(trim(id))
			end if
		next
	end sub
	
	
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'ALBERO COMPLETO DELL'INDICE
'******************************************************************************************************************************************
'******************************************************************************************************************************************
    
    public sub AlberoIndiceCompleto()
        'imposta dati albero
        tree.Name = "IndexAlbero"
        tree.TableCaption = "Indice generale - albero"
        
        'aggiunge elementi dell'albero
        CALL AlberoIndiceCompleto_Explore(0)
        
        'disegna albero
        CALL tree.Write()
    end sub
    
    
    'procedura che "visita" l'albero completo dell'indice per generare nodi e foglie
    private sub AlberoIndiceCompleto_Explore(padre_id)
        dim rs, sql, nome, BaseLink, title
        BaseLink = "IndexGestione.asp?FROM=" & FROM_ALBERO & "&ID="
        
        sql = " SELECT * FROM v_indice LEFT JOIN tb_pagineSito ON (v_indice.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + _
              " WHERE "& SQL_IfIsNull(Index.conn, "idx_padre_id", "0") & "=" & cIntero(padre_id) & _
              " ORDER BY idx_ordine_assoluto, co_ordine, co_titolo_it "
        set rs = Index.conn.execute(Sql, ,adCmdtext)
        
        while not rs.Eof
            nome = JSEncode(rs("co_titolo_it"), """")
            if cString(rs("nome_ps_interno"))<>"" AND uCase(rs("tab_name")) = "TB_PAGINESITO" then
                nome = nome + " <span style='font-size:90%;'>(" & JSEncode(rs("nome_ps_interno"), """") & ")</span> "
            end if
            nome = nome & " ("& rs("tab_titolo") &")"
            'evidenzio il colore del tipo
		    if CString(rs("tab_colore")) <> "" then
			    nome = "<span style='color:"& rs("tab_colore") &"'>"& nome &"</span>"
		    end if
			title = "ordine voce: " & rs("idx_ordine") & " - ordine contenuto: " & rs("co_ordine")
            
            if rs("idx_foglia") then 
                CALL tree.AddLeaf(rs("idx_livello"), nome, title, BaseLink & rs("idx_id"))
            else
                CALL tree.AddNode(rs("idx_livello"), nome, title, BaseLink & rs("idx_id"), rs("idx_id") )
            end if
            
            if tree.IsNodeExpanded(rs("idx_id")) then
                CALL AlberoIndiceCompleto_Explore(rs("idx_id"))
            end if
		    
            'inserisco il ramo NUOVO
            if not rs("idx_foglia") then
                CALL tree.AddNodeNew(rs("idx_livello") + 1, rs("idx_id"), _
                                     "IndexGestione.asp?FROM=" & FROM_ALBERO & "&idx_padre_id=;IndexRaggruppamentoGestione.asp?FROM=" & FROM_ALBERO & "&idx_padre_id=", _
                                     "nuova voce;nuovo raggruppamento")
            end if
		    
            rs.MoveNext
	    wend
	    
        rs.close
        set rs = nothing
    end sub
    
	
	
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'ALBERO PER LA SELEZIOEN DELL'INDICE 
'******************************************************************************************************************************************
'******************************************************************************************************************************************
    
    public sub AlberoIndiceSelezione(Selected, SoloFoglie, SelectableWebId)
        'imposta dati albero
        tree.TableCaption = ""
        'aggiunge elementi dell'albero
        CALL AlberoIndiceSelezione_Explore(cIntero(Selected), SoloFoglie, SelectableWebId, 0)
        
        'disegna albero
        CALL tree.Write()
    end sub
	
    
    'procedura che "visita" l'albero completo dell'indice per generare nodi e foglie per la selezione
    private sub AlberoIndiceSelezione_Explore(Selected, SoloFoglie, SelectableWebId, padre_id)
        dim rs, sql, nome, BaseLink, Link
        BaseLink = "?SELEZIONA="
        
        sql = " SELECT * FROM v_indice LEFT JOIN tb_pagineSito ON (v_indice.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + _
              " WHERE "& SQL_IfIsNull(Index.conn, "idx_padre_id", "0") & "=" & cIntero(padre_id) & _
              " ORDER BY idx_ordine_assoluto, co_ordine, co_titolo_it "
        set rs = Index.conn.execute(Sql, ,adCmdtext)        			
		
        while not rs.Eof			
            nome = JSEncode(rs("co_titolo_it"), """")
            if cString(rs("nome_ps_interno"))<>"" AND uCase(rs("tab_name")) = "TB_PAGINESITO" then
                nome = nome + " <span style='font-size:90%;'>(" & JSEncode(rs("nome_ps_interno"), """") & ")</span> "
            end if
            nome = nome & " <span style='font-size:90%;'>("& rs("tab_titolo") &")</span> "
			
			if Selected = cIntero(rs("idx_id")) then
				nome = "<span class='warning' title='voce attualmente selezionata'>"& nome &"</span>"
			elseif not SoloFoglie OR rs("idx_foglia") then
				nome = "<span title='click per selezionare questa voce'>"& nome &"</span>"
			end if
			
			if (cIntero(SelectableWebId) = 0 OR cIntero(SelectableWebId) = cIntero(rs("idx_webs_id"))) AND _
			   (rs("idx_foglia") OR not SoloFoglie) then
				Link = BaseLink & rs("idx_id")
			else
				Link = ""
			end if
			
			if rs("idx_foglia") then 
                CALL tree.AddLeaf(rs("idx_livello"), nome, "", Link)
            else
                CALL tree.AddNode(rs("idx_livello"), nome, "", Link, rs("idx_id") )
            end if            
			
            if tree.IsNodeExpanded(rs("idx_id")) then			
                CALL AlberoIndiceSelezione_Explore(Selected, SoloFoglie, SelectableWebId, rs("idx_id"))
			elseif not rs("idx_foglia") then			
				CALL tree.AddNodeNew(rs("idx_livello") + 1, rs("idx_id"), "", "")
            end if
		    
            rs.MoveNext
	    wend
	    
        rs.close
        set rs = nothing
    end sub
    
    
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'ALBERO DELL'INDICE FITLRATO PER CONTENUTO
'******************************************************************************************************************************************
'******************************************************************************************************************************************
    
    
    public sub AlberoIndiceByTipoContenuto(ContentTabName, ContentNewLink, ContentNewLabel, ContentModLink)
        'imposta dati albero
        tree.Name = "IndexAlbero_" + ContentTabName
        
		if tree.TableCaption = "" then
	        tree.TableCaption = "Indice delle pagine - albero"
		end if
        
        'aggiunge elementi dell'albero
        CALL AlberoIndiceByTipoContenuto_Explore(0, ContentTabName, ContentNewLink, ContentNewLabel, ContentModLink)
        
        'disegna albero
        CALL tree.Write()
    end sub
    
    
    'procedura che "visita" l'albero completo dell'indice per generare nodi e foglie
    private sub AlberoIndiceByTipoContenuto_Explore(padre_id, ContentTabName, ContentNewLink, ContentNewLabel, ContentModLink)
        dim rs, sql, nome, link, title
        
        sql = " SELECT * FROM v_indice " + _
			  " INNER JOIN tb_siti ON v_indice.tab_sito_id = tb_siti.id_sito " + _
              " LEFT JOIN tb_pagineSito ON (v_indice.idx_link_pagina_id = tb_pagineSito.id_pagineSito AND tb_pagineSito.id_web = "& CIntero(Session("AZ_ID")) &") " + _
              " WHERE "& SQL_IfIsNull(Index.conn, "idx_padre_id", "0") & "=" & cIntero(padre_id) & _
              " ORDER BY idx_ordine_assoluto, co_ordine, co_titolo_it "
        set rs = Index.conn.execute(Sql, ,adCmdtext)
        while not rs.Eof
        
            if ( uCase(ContentTabName) <> "TB_PAGINESITO" OR cIntero(Session("AZ_ID")) = cIntero(rs("idx_webs_id")) ) AND _
			   ( uCase(ContentTabName) = uCase(rs("tab_name")) OR _
                 ( uCase(ContentTabName) = "TB_PAGINESITO" AND cIntero(rs("id_pagineSito")) > 0 ) ) then
               'tipo contenuto uguale a quello filtrato o sono nell'albero delle pagine ed e' presente una pagina
               	link = ContentModLink & rs("idx_id")
            else
                link = ""
            end if
            
            'genera nome del contenuto
            if uCase(ContentTabName) = "TB_PAGINESITO" AND cIntero(rs("id_pagineSito")) > 0 then
                nome = JSEncode(rs("nome_ps_it"), """")
                if cString(rs("nome_ps_interno"))<>"" then
                    nome = nome + " <span style='font-size:90%;'>(" & JSEncode(rs("nome_ps_interno"), """") & ")</span> "
                end if
                nome = "<b>" + nome + "</b>"
                if uCase(rs("tab_name")) <> "TB_PAGINESITO" then
                    nome = nome + ": <span style='color:" & JSEncode(rs("tab_colore"), """") & "'>" + JSEncode(rs("co_titolo_it"), """") + " ("& rs("tab_titolo") &")</span>"
                else
                    nome = "<span style='color:" & JSEncode(rs("tab_colore"), """") & "'>" + nome + "</span>"
                end if
            else
                nome = JSEncode(rs("co_titolo_it"), """")
                if cString(rs("nome_ps_interno"))<>"" AND uCase(rs("tab_name")) = "TB_PAGINESITO" then
                    nome = nome + " <span style='font-size:90%;'>(" & JSEncode(rs("nome_ps_interno"), """") & ")</span> "
                end if
                if uCase(ContentTabName) = uCase(rs("tab_name")) then
                    'contenuto corrente: lo evidenzio
                    nome = "<b>" + nome + "</b>"
                end if
				
                if cString(rs("tab_colore"))<>"" and link <>"" then
                    nome = "<span style='color:" & JSEncode(rs("tab_colore"), """") & "'>" + nome + " </span>"					
				else
					nome = "<span style='color:#888;'>" + nome + " </span>"					
                end if
            end if
			
			'gestione esterna del nodo
			if GestioneEsterna then
				CALL GestioneNodo(Index.conn, rs, nome, link)
			end if
			
			title = "ordine voce: " & rs("idx_ordine") & " - ordine contenuto: " & rs("co_ordine")
			if link = "" then
				title = "---------------------------------------\n" + _
						" CONTENUTO NON MODIFICABILE DA QUI:\n utilizzare l'applicativo: " + rs("sito_nome") + _
					    "\n---------------------------------------\n" + _
						title
				link = "javascript:alert('CONTENUTO NON MODIFICABILE DA QUI:\\n\\nutilizzare l\\'applicativo: " + rs("sito_nome") + "');"
			end if
			
            if rs("idx_foglia") then 
                CALL tree.AddLeaf(rs("idx_livello"), nome, title, link)
            else
                CALL tree.AddNode(rs("idx_livello"), nome, title, link, rs("idx_id") )
            end if
    		
            if tree.IsNodeExpanded(rs("idx_id")) then
                CALL AlberoIndiceByTipoContenuto_Explore(rs("idx_id"), ContentTabName, ContentNewLink, ContentNewLabel, ContentModLink)
            end if
            
    		'inserisco il ramo NUOVO
            if not rs("idx_foglia") then
				if ContentNewLink <> "" AND ContentNewLabel <> "" AND _
				   ( uCase(ContentTabName) <> "TB_PAGINESITO" OR cIntero(Session("AZ_ID")) = cIntero(rs("idx_webs_id")) ) then
					CALL tree.AddNodeNew(rs("idx_livello") + 1, rs("idx_id"), ContentNewLink, ContentNewLabel)
				elseif NOT tree.IsNodeExpanded(rs("idx_id")) then
					'visualizzo un nodo figlio qualsiasi altrimenti non appare il +
					CALL tree.AddNodeNew(rs("idx_livello") + 1, rs("idx_id"), "NODO DA ESPANDERE", "NODO DA ESPANDERE")
				end if
            end if
            
    		rs.MoveNext
    	wend
    	
    	set rs = nothing
        
    end sub
    
    
    
'******************************************************************************************************************************************
'******************************************************************************************************************************************
'ALBERO DELL'INDICE FITLRATO PER CONTENUTO
'******************************************************************************************************************************************
'******************************************************************************************************************************************
    
    
    Public sub AlberoBanner()
        'imposta dati albero
        tree.Name = "AlberoBanner"
        
        tree.TableCaption = "Contratti - visualizzazione ad albero "
        
        'aggiunge elementi dell'albero
        CALL AlberoBanner_Explore(0)
        
        'disegna albero
        CALL tree.Write()
    end sub
    
    
    'procedura che "visita" l'albero completo dell'indice per generare nodi e foglie con i relativi banner pubblicati
    private sub AlberoBanner_Explore(padre_id)
       dim rs, rsb,  sql, nome, ConBanner
        set rsb = Server.CreateObject("ADODB.Recordset")
        
        sql = " SELECT * FROM v_indice LEFT JOIN tb_pagineSito ON (v_indice.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + _
              " WHERE "& SQL_IfIsNull(Index.conn, "idx_padre_id", "0") & "=" & cIntero(padre_id) & _
              " ORDER BY idx_ordine_assoluto, co_ordine, co_titolo_it "
        set rs = Index.conn.execute(Sql, ,adCmdtext)
        
        while not rs.Eof
            nome = JSEncode(rs("co_titolo_it"), """")
            if cString(rs("nome_ps_interno"))<>"" AND uCase(rs("tab_name")) = "TB_PAGINESITO" then
                nome = nome + " <span style='font-size:90%;'>(" & JSEncode(rs("nome_ps_interno"), """") & ")</span> "
            end if
            nome = nome & " ("& rs("tab_titolo") &")"
            'evidenzio il colore del tipo
		    if CString(rs("tab_colore")) <> "" then
			    nome = "<span style='color:"& rs("tab_colore") &"'>"& nome &"</span>"
		    end if
            
            'genera elenco banner per nodo dell'indice
            if instr(1, rs("tab_name"), "tb_contents", vbTextCompare) > 0 then
                'raggruppamento
                ConBanner = false
            else
                sql = " SELECT cb_banner_title_IT, cb_data_inizio, cb_data_fine, cb_impression_iniziali, cb_impression_attuali, cb_click_iniziali, cb_click_attuali, cb_attivo, " + _
                      " cb_id, cb_riferimento_interno, tipoB_nome " + _
                      " FROM ADtb_contratti_banner INNER JOIN ADtb_tipiBanner ON ADtb_contratti_banner.cb_tipo_id = ADtb_tipiBanner.tipoB_id " + _
                      " WHERE cb_id IN ( SELECT pub_cb_id FROM ADtb_contratti_banner_pubblicazioni " + _
                                       " WHERE (IsNull(pub_su_ramo,0)=0 AND pub_index_id=" & cIntero(rs("idx_id")) & ") OR " + _
                                             " (IsNull(pub_su_ramo,0)=1 AND " & SQL_IdListSearch(Index.conn, " '" & rs("idx_tipologie_padre_lista") & "'", "' + CAST(pub_index_id AS nvarchar(10)) + '") & ") " + _
                                       " ) "
                rsb.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
                ConBanner = (not rsb.eof)
            end if
            
            if rs("idx_foglia") AND not ConBanner then 
                CALL tree.AddLeaf(rs("idx_livello"), nome, "", "")
            else
                CALL tree.AddNode(rs("idx_livello"), nome + IIF(ConBanner, " [ con banner ]", ""), "", "", rs("idx_id") )
            end if
            
            if tree.IsNodeExpanded(rs("idx_id")) then
                if ConBanner then
                    'elenca banners
                    CALL AlberoBanner_ElencoBannerNodo(rsb, rs("idx_livello")+1)
                end if
                
                CALL AlberoBanner_Explore(rs("idx_id"))
            end if
		    
            'inserisco il ramo fittizio per generare l'apertura dei rami
            if (not rs("idx_foglia") OR ConBanner ) AND not tree.IsNodeExpanded(rs("idx_id")) then
                CALL tree.AddNodeNew(rs("idx_livello") + 1, rs("idx_id"), "", "")
            end if
		    
            if rsb.State = adStateOpen then
                rsb.close
            end if 
            
            rs.MoveNext
	    wend
	    
        rs.close
        set rs = nothing
    end sub
    
    
    'procedura che ritorna l'elenco dei banner pubblicati nel punto indicato
    private sub AlberoBanner_ElencoBannerNodo(rs, livello)
        dim stato, nome
        if not rs.eof then
            while not rs.eof
                stato = ContrattoBanner_Stato(rs)
                nome = "<table>" + _
                        "<tr>" + _
                            "<td class='content_b" + stato +"'>" + _
                                "<a class='" + stato + "' href='ContrattiMod.asp?ID=" & rs("cb_id") & "' title='apri la scheda del contratto banner'> " + _
                                    JsEncode(rs("cb_banner_title_IT"), """") + _
                                    " <span style='font-weight:normal;'>( banner: " + JsEncode(rs("tipoB_nome"), """") + " )</span> " + _
                                "</a>" + _
                            "</td>" + _
                        "</tr>" + _
                       "</table>"
                CALL tree.AddLeaf(livello, nome, "", "")
                rs.movenext
            wend
        end if
    end sub
    
    
end class



'************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************
'CLASSE CHE GENERA LA LEGENDA PER L'INDICE
'************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************

class ObjIndexLegend
    
    
'******************************************************************************************************************************************
'VARIABILI E PROPRIETA' PUBBLICHE
'******************************************************************************************************************************************
    'variabile indice sulla base della quale viene generata la legenda
    public Index
    public ContentTabName
    public ContentTitle
    public PositionLeft
    public PositionTop
    
'******************************************************************************************************************************************
'VARIABILI E PROPRIETA' PRIVATE
'******************************************************************************************************************************************
    private ExtraContentRows
    
    
'******************************************************************************************************************************************
'COSTRUTTORI CLASSE
'******************************************************************************************************************************************
    
    Private Sub Class_Initialize()
       set ExtraContentRows = Server.CreateObject("Scripting.Dictionary")
		ExtraContentRows.CompareMode = vbTextCompare
        
        PositionLeft = 540
        PositionTop = 65
    End Sub
    
    Private Sub Class_Terminate()
    	set ExtraContentRows = nothing
    End Sub
    
    
'******************************************************************************************************************************************
'METODI PUBBLICI
'******************************************************************************************************************************************

    'aggiunge una riga aggiuntiva non derivata dal database
    Public sub AddExtra(color, label)
        if not ExtraContentRows.Exists(cString(color)) then
            CALL ExtraContentRows.Add(cString(color), cString(label))
        else
            ExtraContentRows(cString(color)) = cString(label)
        end if
    end sub
    
    
    'scrive la lengenda
    Public sub Write() 
        %>
        <script type="text/javascript">
            function LegendaVisualizza() {
                var elenco;
            	elenco = document.getElementById("legendaElenco");
            	if (elenco.style.display == "none")
            	    elenco.style.display = "block";
            	else
            	    elenco.style.display = "none";
            }
        </script>
        <div id="legenda" style="position: absolute; left: <%= PositionLeft %>px; top: <%= PositionTop %>px; width: 180px; height: 200px; z-index: 3;">
		    <table cellpadding="0" cellspacing="0" class="tabella_madre" style="background-color:#FFF;">
			    <caption onclick="LegendaVisualizza()" style="cursor: help;">
                    <span style="float: right;">&laquo;&raquo;</span>
            	    LEGENDA
                </caption>
                <tr id="legendaElenco" style="display: none;">
				    <td style="width:100%; display:block;">
					    <table cellpadding="0" cellspacing="1" width="100%">
                            <% if cString(ContentTabName)<>"" then 
                                if ContentTitle<>"" then%>
                                    <tr>
                                        <th><%= ContentTitle %></th>
                                    </tr>
                                <% end if
                                
                                'elenca righe del contenuto
                                CALL WriteContentRows(false, ContentTabName, "")
                                
                                'elenca righe aggiuntive
                                if ExtraContentRows.count > 0 then
                                    dim color
                                    'scrive righe aggiuntive per il contenuto
                                    for each color in ExtraContentRows.keys 
                                        CALL WriteContentRow(ContentTabName, color, ExtraContentRows(color)) 
                                    next
                                end if
                                
                                CALL WriteContentRows(true, "", ContentTabName)
                            else
                                CALL WriteContentRows(false, "", "")
                            end if %>
                            
                        </table>
                    </td>
                </tr>
            </table>
        </div>
        <%    
    end sub
    
    
'******************************************************************************************************************************************
'METODI PRIVATI
'******************************************************************************************************************************************
    
    private sub WriteContentRows(Internal, ContentTabFilter, ContentTabExclude)
        dim rs, sql, CurrentSite
        set rs = Server.createObject("ADODB.recordset")
        
        sql = " SELECT * FROM tb_siti INNER JOIN tb_siti_tabelle ON tb_siti.id_sito = tb_siti_tabelle.tab_sito_id " + _
              " WHERE tab_colore <> '' AND NOT "& SQL_IsNull(Index.conn, "tab_colore")
        if ContentTabFilter<>"" then
            sql = sql + " AND tab_name LIKE '" & ParseSql(ContentTabName, adChar) & "' "
        end if
        if ContentTabExclude<>"" then
            sql = sql + " AND NOT(tab_name LIKE '" & ParseSql(ContentTabExclude, adChar) & "') "
        end if
        sql = sql + " ORDER BY sito_nome, tab_titolo"
        
        rs.open sql, Index.conn, adOpenStatic, adLockOptimistic
        CurrentSite = 0
        if Internal AND not rs.eof then %>
            <tr>
               <th>ALTRI CONTENUTI</th>
            </tr> 
        <%end if
        while not rs.eof 
            if CurrentSite <> rs("id_sito") then %>
               <% dim value
				if InStr(rs("sito_nome"), "[") > 0 then
					value = Left(rs("sito_nome"), InStr(rs("sito_nome"), "[") - 1)
				else
					value = rs("sito_nome")
				end if				
			   %>
			   <tr>
                    <th <% if Internal then %>class="L2"<% end if %>><%= value %></th>
               </tr> 
                <% CurrentSite = rs("id_sito") 
            end if
            
            CALL WriteContentRow(rs("tab_name"), rs("tab_colore"), rs("tab_titolo"))
            
            rs.movenext
        wend
        rs.close
    end sub
    
    
    'scrive riga della legenda
    private sub WriteContentRow(ContentType, color, label)%>
       <tr>
            <% if instr(1, color, "#", vbTextCompare)>0 then
                'evidenziato da un colore
				%>
                <td class="<%= IIF(ContentType = ContentTabName, "content_b", "content") %>">
                    <span class="icona" style="background-color:<%= color %>;">&nbsp;</span>
                    <span style="color: <%= color %>;"><%= label %></span>
                </td>
            <% else
                'evidenzato da una classe di stili
				%>
                <td class="<%= IIF(ContentType = ContentTabName, "content_b", "content") %><%= color %>">
                    <span class="icona<%= color %>">&nbsp;</span>
                    <span><%= label %></span>
                </td>
            <% end if %>
        </tr> 
    <% end sub
    
end class
 %>