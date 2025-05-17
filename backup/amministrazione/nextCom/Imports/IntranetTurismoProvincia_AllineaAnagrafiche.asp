<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% Server.ScriptTimeout = 2147483647 %>
<% response.buffer = false %>

<% 
Titolo_sezione = "Turismo.provincia.venezia.it - Allineamento anagrafiche con AOL"
Action = "INDIETRO"
href = "../default.asp"

import_no_login = true
%>
<!--#include file="Intestazione.asp"-->
<!--#INCLUDE FILE="../../library/CLASSWEBSERVICE.ASP"-->
<%
dim conn, ws, ws_r, rs, sql, ID, CntObj, ID_anagr, ID_cntRel, ID_cntSede, ID_rubrica, ID_tipoNumero
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")


dim admin_login, admin_password, lastDateModify
admin_login = "NEXTAIM"
admin_password = "5315TRS"
lastDateModify = Trim(cString(GetValueList(conn, NULL, "SELECT TOP 1 ISNULL(cnt_modData, cnt_insData) FROM tb_indirizzario WHERE PraticaCount = 99999 ORDER BY cnt_modData DESC, cnt_insData DESC")))
if lastDateModify = "" then
	lastDateModify = "2013-01-03 00:00:00"
end if


set ws = new WebService
set ws_r = new WebService
CALL ws.open(GetModuleParam(conn, "NEXT-COM_ANAGRAFICHE_WSDL_URL"))
CALL ws_r.open(GetModuleParam(conn, "NEXT-COM_ANAGRAFICHE_WSDL_URL"))



%>
<div id="content">	
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-top:0px; margin-bottom:0px; width:750px;">
		<% if request("importa_contatti")="" then %>
			<caption class="border">Conferma esecuzione import</caption>
            <form action="" method="post" id="form1" name="form1">
				<tr>
					<td class="label" style="">Vuoi importare le anagrafiche da AOL?</td>
				</tr>
				<tr>
					<td class="footer" colspan="2">
						<input style="width:20%;" type="submit" class="button" name="importa_contatti" value="SI &gt;&gt;">
					</td>
				</tr>
			</form>
		<% else %>
			<caption class="border">Esecuzione import</caption>
			<tr>
				<th style="width:15%">ID</th>
				<th>DENOMINAZIONE</th>
			</tr>
			<%
			
			conn.BeginTrans
			
			'----- Aggiorno l'elenco tipi recapiti
			CALL ws.GetData(ws.soap.GetRecapitiElenco(admin_login, admin_password) , "TipiRecapiti")
			while not ws.EOF()
				sql = "SELECT * FROM tb_tipNumeri WHERE nome_tipoNumero LIKE '"&ws("nome_tipoNumero")&"'"
				rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
				if rs.eof then
					rs.AddNew
					rs("id_tipoNumero") = cIntero(GetValueList(conn, NULL, "SELECT MAX(id_tipoNumero) + 1 FROM tb_tipNumeri"))
					rs("nome_tipoNumero") = ws("nome_tipoNumero")
					rs("nome_tiponumero_it") = ws("nome_tiponumero_it")
					rs("nome_tiponumero_en") = ws("nome_tiponumero_en")
					rs("nome_tiponumero_de") = ws("nome_tiponumero_de")
					rs("nome_tiponumero_fr") = ws("nome_tiponumero_fr")
					rs("nome_tiponumero_es") = ws("nome_tiponumero_es")
					rs.Update
				end if
				rs.Close
				ws.MoveNext()
			wend
			
			
			'----- Aggiorno l'elenco rubriche
			CALL ws.GetData(ws.soap.GetRubricheElenco(admin_login, admin_password) , "Rubriche")
			while not ws.EOF()
				sql = "SELECT * FROM tb_rubriche WHERE nome_Rubrica LIKE '" & ParseSQL(ws("nome_Rubrica"), adChar) & "'"
				rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
				if rs.eof then
					rs.AddNew
					rs("nome_Rubrica") = ws("nome_Rubrica")
					rs("note_Rubrica") = ws("note_Rubrica")
					rs("locked_rubrica") = cBoolean(ws("locked_rubrica"), false)
					rs("rubrica_esterna") = cBoolean(ws("rubrica_esterna"), false)
					rs("SyncroTable") = NULL
					rs("SyncroFilterTable") = NULL
					rs("SyncroFilterKey") = NULL
					rs.Update
				end if
				rs.Close
				ws.MoveNext()
			wend
			
			
			'----- Aggiorno le anagrafiche
			CALL ws.GetData(ws.soap.GetAnagrafiche(admin_login, admin_password, lastDateModify) , "Anagrafiche")
			
			set CntObj = new IndirizzarioLock
			set CntObj.conn = conn
				
			while not ws.EOF()
				if cIntero(ws("IDElencoIndirizzi")) > 0 then
					%>
					<tr>
						<td class="content" style="width:15%"><%= ws("IDElencoIndirizzi") %></td>
						<td class="content"><%= cString(ws("NomeOrganizzazioneElencoIndirizzi")) & " - " & ws("NomeElencoIndirizzi") & " " & cString(ws("CognomeElencoIndirizzi")) %></td>
					</tr>
					</table>
					<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-top:0px; margin-bottom:0px; width:750px;">
					<%
					
					CntObj.RemoveALL()
					CntObj("NomeElencoIndirizzi") = cString(ws("NomeElencoIndirizzi"))
					CntObj("SecondoNomeElencoIndirizzi") = cString(ws("SecondoNomeElencoIndirizzi"))
					CntObj("CognomeElencoIndirizzi") = cString(ws("CognomeElencoIndirizzi"))
					CntObj("TitoloElencoIndirizzi") = cString(ws("TitoloElencoIndirizzi"))
					CntObj("NomeOrganizzazioneElencoIndirizzi") = cString(ws("NomeOrganizzazioneElencoIndirizzi"))
					CntObj("QualificaElencoIndirizzi") = cString(ws("QualificaElencoIndirizzi"))
					CntObj("IndirizzoElencoIndirizzi") = cString(ws("IndirizzoElencoIndirizzi"))
					CntObj("CittaElencoIndirizzi") = cString(ws("CittaElencoIndirizzi"))
					CntObj("StatoProvElencoIndirizzi") = cString(ws("StatoProvElencoIndirizzi"))
					CntObj("ZonaElencoIndirizzi") = cString(ws("ZonaElencoIndirizzi"))
					if cIntero(cString(ws("CAPElencoIndirizzi"))) > 0 then
 						CntObj("CAPElencoIndirizzi") = cString(ws("CAPElencoIndirizzi"))
					end if
					CntObj("CountryElencoIndirizzi") = cString(ws("CountryElencoIndirizzi"))
					CntObj("DTNASCElencoIndirizzi") = cString(ws("DTNASCElencoIndirizzi"))
					CntObj("NoteElencoIndirizzi") = cString(ws("NoteElencoIndirizzi"))
					CntObj("isSocieta") = cBoolean(ws("isSocieta"), false)
					CntObj("ModoRegistra") = cString(ws("ModoRegistra"))
					CntObj("DataIscrizione") = DateISO(ws("DataIscrizione"))
					CntObj("LockedByApplication") = NULL
					CntObj("ApplicationsLocker") = NULL
					CntObj("SyncroKey") = NULL
					CntObj("SyncroTable") = NULL
					CntObj("SyncroApplication") = NULL
					CntObj("LocalitaElencoIndirizzi") = cString(ws("LocalitaElencoIndirizzi"))
					CntObj("LuogoNascita") = cString(ws("LuogoNascita"))
					CntObj("CF") = cString(ws("CF"))
					CntObj("lingua") = cString(ws("lingua"))
					CntObj("codiceInserimento") = cString(ws("codiceInserimento"))
					CntObj("google_maps_latitudine") = cString(ws("google_maps_latitudine"))
					CntObj("google_maps_longitudine") = cString(ws("google_maps_longitudine"))
					CntObj("partita_iva") = cString(ws("partita_iva"))

					CntObj("PraticaCount") = 99999
					CntObj("PraticaPrefisso") = ws("IDElencoIndirizzi")
					
					ID_anagr = GetIDElencoIndirizziByPraticaPrefisso(ws("IDElencoIndirizzi"))
					
					if ID_anagr > 0 then
						CntObj("IDElencoIndirizzi") = ID_anagr
						CntObj.UpdateDB()
					else
						ID_anagr = CntObj.InsertIntoDB()
					end if
				end if
				
				ws.MoveNext()
			wend
			
			
			ws.MoveFirst()
			while not ws.EOF()
				if cIntero(ws("IDElencoIndirizzi")) > 0 then
				
					ID_anagr = GetIDElencoIndirizziByPraticaPrefisso(ws("IDElencoIndirizzi"))
				
					
					if ID_anagr > 0 then
					
						'----- Imposto i contatti interni e le sedi
						CntObj.RemoveALL()
						CntObj.LoadFromDB(ID_anagr)
						
						if cIntero(ws("cntRel")) > 0 then
							ID_cntRel = GetIDElencoIndirizziByPraticaPrefisso(ws("cntRel"))
							CntObj("cntRel") = ID_cntRel
						end if
						
						if cIntero(ws("CntSede")) > 0 then
							ID_cntSede = GetIDElencoIndirizziByPraticaPrefisso(ws("CntSede"))
							CntObj("CntSede") = ID_cntSede
						end if

						CntObj.UpdateDB()
					
					
					
						'----- Importo i recapiti
						CALL ws_r.GetData(ws_r.soap.GetRecapitiElencoById(admin_login, admin_password, cIntero(ws("IDElencoIndirizzi"))) , "RecapitiAnagrafica")
						while not ws_r.EOF()
							ID_tipoNumero = cIntero(GetValueList(conn, NULL, "SELECT id_tipoNumero FROM tb_tipNumeri WHERE nome_tipoNumero LIKE '" & ParseSQL(ws_r("NomeTipoNumero"), adChar) & "'"))
							if ID_tipoNumero > 0 AND Trim(ws_r("ValoreNumero")) <> "" then
								sql = " SELECT * FROM tb_ValoriNumeri WHERE " & _
									  " id_tipoNumero IN ("&ID_tipoNumero&") AND " & _
									  " id_indirizzario IN ("&ID_anagr&") AND " & _
									  " ValoreNumero LIKE '"&ParseSQL(Trim(ws_r("ValoreNumero")), adChar)&"'"
								rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
								if rs.eof then
									rs.AddNew
									rs("id_indirizzario") = ID_anagr
									rs("id_tipoNumero") = ID_tipoNumero
									rs("ValoreNumero") = Trim(ws_r("ValoreNumero"))
									rs("email_default") = cBoolean(ws_r("email_default"), false)
									rs("SyncroField") = NULL
									rs("protetto_privacy") = cBoolean(ws_r("protetto_privacy"), false)
									rs("email_newsletter") = cBoolean(ws_r("email_newsletter"), false)
									rs.Update
								end if
								rs.Close
							end if
							ws_r.MoveNext()
						wend
						ws_r.Close()
						
						
						
						'----- Associo i clienti alle rubriche
						CALL ws_r.GetData(ws_r.soap.GetRubricheElencoById(admin_login, admin_password, cIntero(ws("IDElencoIndirizzi"))) , "RubricheContatto")
						while not ws_r.EOF()
							ID_rubrica = cIntero(GetValueList(conn, NULL, "SELECT id_rubrica FROM tb_rubriche WHERE nome_Rubrica LIKE '" & ParseSQL(ws_r("nome_Rubrica"), adChar) & "'"))
							if id_rubrica > 0 then 
								sql = " SELECT * FROM rel_rub_ind WHERE " & _
									  " id_rubrica IN ("&ID_rubrica&") AND " & _
									  " id_indirizzo IN ("&ID_anagr&")"
								rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
								if rs.eof then
									rs.AddNew
									rs("id_rubrica") = ID_rubrica
									rs("id_indirizzo") = ID_anagr
									rs.Update
								end if
								rs.Close
							end if
							ws_r.MoveNext()
						wend
						ws_r.Close()
						
					end if
					
				end if
				ws.MoveNext()
			wend
			ws.Close()
			
			conn.CommitTrans
			
			%>
			<tr>
				<td class="footer" colspan="4">
					AGGIORNAMENTO ANAGRAFICHE COMPLETATO
				</td>
			</tr>
		<% end if %>
	</table>
</div>
</body>
</html>
<%

function GetIDElencoIndirizziByPraticaPrefisso(praticaPrefisso)
	sql = "SELECT IDElencoIndirizzi FROM tb_indirizzario WHERE PraticaPrefisso LIKE '"&praticaPrefisso&"'"
	GetIDElencoIndirizziByPraticaPrefisso = cIntero(GetValueList(conn, NULL, sql))
end function


%>