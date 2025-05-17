<% 


'procedura che disegna l'area di form per la scelta o l'immissione della rubrica
sub FORM_SelezioneRubrica(conn) %>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
    <tr>
    	<td class="label" rowspan="4" style="width:18%;">rubrica di destinazione:</td>
        <td class="label_no_width" style="width:12%;">
            <input type="radio" class="checkbox" name="sel_tipo_rubrica" id="sel_tipo_rubrica_0" <%=chk(cInteger(request("sel_tipo_rubrica"))=0)%> value="0" onclick="SetStato_TIPO()">
            esistente:
        </td>
    	<td class="content" colspan="3">
    		<% sql = " SELECT id_rubrica, nome_rubrica FROM tb_rubriche " &_
    				 " ORDER BY nome_rubrica"
    		CALL dropDown(conn, sql, "id_rubrica", "nome_rubrica", "rubrica_import", request("rubrica_import"), false, "", LINGUA_ITALIANO)%>(*)<br>
    		<span class="note">Selezionare la rubrica nella quale verranno inseriti i contatti.</span>
    	</td>
    </tr>  
    <tr>
        <td class="label_no_width" rowspan="3">
            <input type="radio" class="checkbox" name="sel_tipo_rubrica" id="sel_tipo_rubrica_1" <%=chk(cInteger(request("sel_tipo_rubrica"))=1)%> value="1" onclick="SetStato_TIPO()">
            nuova:
        </td>
        <td class="label" rowspan="3" style="width:8%;">nome:</td>
        <td class="content_center" style="width:5%;">
            <input type="radio" class="checkbox" name="sel_nome_rubrica" id="sel_nome_rubrica_2" <%=chk(cInteger(request("sel_nome_rubrica"))=2)%> value="2" onclick="SetStato_TIPO()">
        </td>
        <td class="content">
            uguale al nome della tabella
        </td>
    </tr>
    <tr>
        <td class="content_center" style="width:5%;">
            <input type="radio" class="checkbox" name="sel_nome_rubrica" id="sel_nome_rubrica_1" <%=chk(cInteger(request("sel_nome_rubrica"))=1)%> value="1" onclick="SetStato_TIPO()">
        </td>
        <td class="content">
            uguale al nome del file
        </td>
    </tr>
    <tr>
        <td class="content_center" style="width:5%;">
            <input type="radio" class="checkbox" name="sel_nome_rubrica" id="sel_nome_rubrica_0" <%=chk(cInteger(request("sel_nome_rubrica"))=0)%> value="0" onclick="SetStato_TIPO()">
        </td>
        <td class="content">
            inserito manualmente:<br>
            <input type="text" class="text" name="nuova_rubrica" value="<%= request("nuova_rubrica") %>" maxlength="250" size="70">
        </td>
    </tr>
</table>
    <script language="JavaScript1.1" type="text/javascript">
        function SetStato_TIPO(){
            EnableIfChecked(form1.sel_tipo_rubrica_0, form1.rubrica_import);
            DisableIfChecked(form1.sel_tipo_rubrica_0, form1.sel_nome_rubrica_2);
            DisableIfChecked(form1.sel_tipo_rubrica_0, form1.sel_nome_rubrica_1);
            DisableIfChecked(form1.sel_tipo_rubrica_0, form1.sel_nome_rubrica_0);
            DisableControl(form1.nuova_rubrica, form1.sel_tipo_rubrica_0.checked || !form1.sel_nome_rubrica_0.checked);
        }
        
        SetStato_TIPO();
    </script>
<%end sub


'funzione che ritorna la prima tabella del database
function GetFirstTable(conn, Obbligatorio)
    dim rs
    'recupera elenco delle tabelle del database
    set rs = Conn.OpenSchema(adSchemaTables)
    
    if not rs.eof then
        GetFirstTable = rs("table_name")
    elseif Obbligatorio then %> 
        <tr>
            <td class="errore" colspan="3">TABELLA PRINCIPALE NON TROVATA</td>
        </tr>
    <%end if
    
    set rs = nothing
end function


'funzione che ritorna l'id della rubrica nella quale inserire i dati importati
function GestioneRubrica(conn, FilePath, TableName)
    
    if cInteger(request("rubrica_import"))>0 then
        GestioneRubrica = cInteger(request("rubrica_import"))
    else
        'inserimento nuova rubrica
        dim sql, NomeRubrica
        
        Select case cInteger(request("sel_nome_rubrica"))
            case 0
                NomeRubrica = request("nuova_rubrica")
            case 1
                NomeRubrica = right(FilePath, len(FilePath) - instrrev(replace(FilePath, "/", "\"), "\"))
            case else
                NomeRubrica = TableName
        end select
        
        sql = " INSERT INTO tb_rubriche (nome_rubrica, rubrica_esterna, locked_rubrica, note_rubrica) " + _
			  " VALUES ('" & ParseSql(NomeRubrica, adChar) & "', 0, 0, 'Rubrica importata') "
		CALL conn.execute(sql, ,adExecuteNoRecords)
				
		sql = "SELECT MAX(id_rubrica) FROM tb_rubriche"
		GestioneRubrica = cInteger(GetValueList(conn, NULL, sql))
    end if
end function


	

 %>