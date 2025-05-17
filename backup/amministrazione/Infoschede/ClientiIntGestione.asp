<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/classIndirizzarioLock.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede.asp" -->
<%
dim OBJ_contatto, sql, conn
set OBJ_contatto = new IndirizzarioLock

if request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim cnt_id, ut_id, mandatory_flds, EsitoOk
	OBJ_contatto.LoadFromForm("chk_isSocieta;chk_abilitato;chk_TipoProfilo")
	
	mandatory_flds = "CognomeElencoIndirizzi;NomeElencoIndirizzi;"
	
	'controlla altri campi obbligatori
	if OBJ_contatto.ValidateFields(mandatory_flds, OBJ_contatto("TipoProfilo")) then
	
		'controllo esito positivo: dati validi
		OBJ_contatto.conn.beginTrans
		
		if request("ID") <> "" then			'sono in modifica o ho scelto dal menu
			OBJ_contatto.UpdateDB()
			cnt_id = OBJ_contatto("IDElencoIndirizzi")
		else
			cnt_id = OBJ_contatto.InsertIntoDB()
		end if
		
		if OBJ_contatto("TipoProfilo") then
			'registrazione utente
			ut_id = OBJ_Contatto.UserFromContact(cnt_id, UTENTE_PERMESSO_SUBCLIENTE)
		else
			CALL obj_contatto.RemoveUserFormContact(cnt_id, 0, UTENTE_PERMESSO_SUBCLIENTE)
		end if
		
		'chiude transazione e conferma dati
		OBJ_contatto.conn.commitTrans %>
		<script language="JavaScript" type="text/javascript">
			opener.location.reload(true);
			window.close();
		</script>
		<%
	end if
end if

if request("ID") <> "" then
	'modifica del contatto
	OBJ_contatto.LoadFromDB(request("ID"))
	if OBJ_contatto.Fields.Exists("ut_id") then
		OBJ_contatto.Fields.Add "tipoprofilo", true
	end if
else
	OBJ_contatto.LoadFromForm("chk_isSocieta;chk_abilitato;chk_TipoProfilo")
end if



'--------------------------------------------------------
if request("ID")="" then
	sezione_testata = "Inserimento nuovo contatto interno" 
else
	sezione_testata = "Modifica dati contatto interno" 
end if%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'-----------------------------------------------------
%> 
<script language="JavaScript" type="text/javascript">
	function set_modo_registra(){
		form1.tft_modoregistra.value = form1.tft_cognomeelencoindirizzi.value;
		return true;
	}
</script>

<div id="content_ridotto">
<form action="" method="post" id="form1" name="form1" onsubmit="set_modo_registra();">
	<input type="hidden" name="tft_modoregistra" value="">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:0px;">
		<% 	if request("ID") = "" then 
			OBJ_contatto("CntRel") = request("CNT") %>
			<input type="hidden" name="tfn_CntRel" value="<%= OBJ_contatto("CntRel") %>">
			<caption>Inserimento nuovo contatto interno</caption>
		<% 	else %>
			<input type="hidden" name="tfn_IDElencoIndirizzi" value="<%= request("ID") %>">
			<caption>Modifica dati del contatto interno</caption>
		<% 	end if %>
		<tr><th colspan="4">ANAGRAFICA</th></tr>
		<tr>
			<input type="hidden" name="chk_isSocieta" id="chk_issocieta_false" value="" <%= chk(not OBJ_contatto("isSocieta"))%>>
			<td class="label" nowrap style="width:18%;">nome:</td>
			<td class="content">
				<input type="text" class="text" name="tft_nomeelencoindirizzi" value="<%= OBJ_contatto("NomeElencoIndirizzi") %>" maxlength="100" size="38">
				<span id="nome">(*)</span>
			</td>
			<td class="label_right">lingua:</td>
			<td class="content">
				<%CALL DropLingue(OBJ_contatto.conn, NULL, "tft_lingua", OBJ_contatto("lingua"), true, false, "") %>
			</td>
		</tr>
		<tr>
			<td class="label">cognome:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_cognomeelencoindirizzi" value="<%= OBJ_contatto("CognomeElencoIndirizzi") %>" maxlength="100" size="38" onChange="set_modo_registra();">
				<span id="cognome">(*)</span>
			</td>
		</tr>
		<tr>
			<td class="label">ruolo / qualifica:</td>
			<td class="content" colspan="3"><input type="text" class="text" name="tft_qualificaelencoindirizzi" value="<%= OBJ_contatto("qualificaelencoindirizzi") %>" maxlength="250" size="38"></td>
		</tr>
	</table>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<tr><th colspan="4">INDIRIZZO</th></tr>
		<tr>
			<td class="label" style="width:19%;">indirizzo:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_IndirizzoElencoIndirizzi" value="<%= OBJ_contatto("IndirizzoElencoIndirizzi") %>" maxlength="250" size="55">
			</td>
		</tr>
		<tr>
			<td class="label">localit&agrave;:</td>
			<td class="content"><input type="text" class="text" name="tft_localitaElencoIndirizzi" value="<%= OBJ_contatto("localitaElencoIndirizzi") %>" maxlength="50" size="35"></td>
			<td class="label">cap:</td>
			<td class="content"><input type="text" class="text" name="tft_CAPElencoIndirizzi" value="<%= OBJ_contatto("CAPElencoIndirizzi") %>" maxlength="20" size="8"></td>
		</tr>
		<tr>
			<td class="label">citt&agrave;:</td>
			<td class="content">
				<input type="text" class="text" name="tft_cittaElencoIndirizzi" value="<%= OBJ_contatto("cittaElencoIndirizzi") %>" maxlength="50" size="35">
			</td>
			<td class="label">provincia:</td>
			<td class="content"><input type="text" class="text" name="tft_StatoProvElencoIndirizzi" value="<%= OBJ_contatto("StatoProvElencoIndirizzi") %>" maxlength="50" size="8"></td>
		</tr>
		<tr><th colspan="4">RECAPITI</th></tr>
		<tr>
			<td class="label">telefono:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_telefono" value="<%= OBJ_contatto("telefono") %>" maxlength="250" size="45">
			</td>
		</tr>
		<tr>
			<td class="label">fax:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_fax" value="<%= OBJ_contatto("fax") %>" maxlength="20" size="45">
			</td>
		</tr>
		<tr>
			<td class="label">cellulare:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_cellulare" value="<%= OBJ_contatto("cellulare") %>" maxlength="20" size="45">
			</td>
		</tr>
		<tr>
			<td class="label">email:</td>
			<td class="content" colspan="3">
				<input type="text" class="text" name="tft_email" value="<%= OBJ_contatto("email") %>" maxlength="250" size="55">
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
				<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
			</td>
		</tr>
	</table>
	<input type="hidden" name="chk_tipoprofilo" id="chk_tipoprofilo_false" value="" <%= chk(not OBJ_contatto("tipoprofilo"))%>>
</form>
</div>
</body>
</html>
<% 
set OBJ_contatto = nothing
%>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>