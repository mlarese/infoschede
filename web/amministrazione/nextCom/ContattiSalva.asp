<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassSalva.asp" -->
<!--#INCLUDE FILE="../library/Tools.asp" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<%
dim Classe
	Set Classe = New OBJ_Salva
	
	'Impostazione parametri
	Classe.ConnectionString 		= Application("DATA_ConnectionString")
	if request("isLocked")="" then
		if request("tfn_CntRel")="" then
			'contatto normale:
			'se il contatto &egrave; esterno o bloccato &egrave; collegato ad almeno una rubrica
			if cInteger(request("LockedByApplication"))=0 then
				Classe.Requested_Fields_List	 = "rubriche;"
			end if
		end if
		
		if request("chk_isSocieta")<>"" then
			'contatto registrato come societa' o sede del contatto interno
			Classe.Requested_Fields_List	 = Classe.Requested_Fields_List + "tft_NomeOrganizzazioneElencoIndirizzi"
		else
			'contatto registrato come persona fisica o contatto interno
			Classe.Requested_Fields_List	 = Classe.Requested_Fields_List + "tft_NomeElencoIndirizzi; tft_CognomeElencoIndirizzi"
		end if
	end if
	Classe.Checkbox_Fields_List 	= "chk_isSocieta"
	Classe.Page_Ins_Form			= ""
	Classe.Page_Mod_Form			= ""
	Classe.Next_Page				= ""	'impostata nella gestione delle relazioni
	Classe.Next_Page_ID				= FALSE
	Classe.Table_Name				= "tb_Indirizzario"
	Classe.id_Field					= "IDElencoIndirizzi"
	Classe.Read_New_ID				= TRUE
	Classe.isReport 				= TRUE
	Classe.Gestione_Relazioni 		= TRUE
	
	if cIntero(request("tfn_CntRel"))=0 then
		CALL Classe.AddForcedValue("CntRel", NULL)
	end if
	if cIntero(request("tfn_CntSede"))=0 then
		CALL Classe.AddForcedValue("CntSede", NULL)
	end if

'definizione eventuali operazioni su relazioni	
Sub Gestione_Relazioni_record(conn, rs, ID)
	
	'genero il codiceInserimento
	CALL SetCodiceInserimento(conn, ID)
	
	if request("tfn_cntRel") = "" then
		'contato normale: gestisce relazioni con rubriche e pagina successiva
		dim sql, i, rubriche
	
		if request("ID")<>"" then
			sql = "DELETE FROM rel_rub_ind WHERE id_indirizzo=" & ID &_
				  " AND id_rubrica NOT IN (SELECT id_rubrica FROM tb_rubriche WHERE " & SQL_IsTrue(conn, "rubrica_esterna") & ") "
			CALL conn.execute(sql, 0, adExecuteNoRecords)
		end if
		
		rubriche = split(request("rubriche"), ",")
		for i = lbound(rubriche) to ubound(rubriche)
			sql = "INSERT INTO rel_rub_ind (id_indirizzo, id_rubrica) VALUES (" & ID & ", " & rubriche(i) & ")"
			CALL conn.execute(sql, 0, adExecuteNoRecords)
		next
		
		if cInteger(request("ext_new_campagna_id")) > 0 then 
			sql = " INSERT INTO rel_cnt_campagne(rcc_cnt_id, rcc_campagna_id) VALUES("&ID&","&request("ext_new_campagna_id")&") "
			CALL conn.execute(sql, 0, adExecuteNoRecords)
		end if
		
		if request("ID") = "" then
			'inserimento di un nuovo contatto: controlla direttamente recapiti
			sql = "SELECT * FROM tb_TipNumeri" 
			rs.open sql, conn, adOpenStatic, adLockOptimistic
			while not rs.eof	
				if request("recapito_" & rs("id_TipoNumero"))<>"" then
					if rs("id_TipoNumero") <> VAL_EMAIL OR isEmail(request("recapito_" & rs("id_TipoNumero"))) then
						'registra recapito
						sql = " INSERT INTO tb_valoriNumeri (id_Indirizzario, id_TipoNumero, protetto_privacy, email_default, ValoreNumero, email_newsletter) " + _
							  " VALUES  (" & ID & ", " & rs("id_TipoNumero") & ", 0, " & IIF(rs("id_TipoNumero") = VAL_EMAIL, "1", "0") & _
							  " , '" + ParseSql(request("recapito_" & rs("id_TipoNumero")), adChar) & "', "&IIF(cBoolean(request("email_newsletter"),false),"1","0")&")"
						CALL conn.execute(sql, , adExecuteNoRecords)
					else
						'errore nell'email
						Session("ERRORE") = Session("ERRORE") & "L'email inserita non &egrave; corretta. "
					end if
				end if
				rs.movenext
			wend
			rs.close
		end if
		
		if request("ID") = "" then
			CALL UpdateParams(conn, "tb_indirizzario", "cnt_", "IDElencoIndirizzi", ID, true)
		else
			CALL UpdateParams(conn, "tb_indirizzario", "cnt_", "IDElencoIndirizzi", ID, false)
		end if
		
		'gestione caratteristiche
		CALL DesSalva(conn, ID, "rel_cnt_ctech", "ric_valore_", "ric_cnt_id", "ric_ctech_id")
		
		'imposta parametri per passare alla pagina successiva
		Classe.isReport = FALSE
		if request("salva_elenco")<>"" then
			Classe.Next_Page = "Contatti.asp"
		else
			Classe.Next_Page = "ContattiMod.asp?ID=" & ID
		end if
	else
		if request("CNT") <> "" then
			%>
			<script language="JavaScript" type="text/javascript">
				opener.location.reload(true);
				window.location = "ContattiInterniMod.asp?ID=<%=ID%>&CNT=<%=request("CNT")%>";
			</script>
			<%
		else
			'contatto interno: gestisce reload pagina via javascript
			%>
			<script language="JavaScript" type="text/javascript">
				opener.location.reload(true);
				//window.close();
			</script>
			<%
		end if
	end if
end Sub

	'salvataggio/modifica dati
	Classe.Salva()
	
%>