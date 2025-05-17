<!--#INCLUDE FILE="Tools_Infoschede_Const.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede_Categorie.asp" -->
<!--#INCLUDE FILE="../nextB2B/Tools_B2B.asp" -->

<%
'**************************************************************************************************
'imposta valore specifico per la gestione delle foto dei pacchetti.
'oArticoliFoto.ElementTableName = "gtb_articoli"
'oArticoliFoto.ElementIndexTableName = "tv_pacchetti"
'**************************************************************************************************

const StatoSchedaConclusa = 1 'id dello stato delle schede Concluse

sub ArticoloLink(id, label, codice)%>
	<a href="javascript:void(0);" title="apri scheda dell'articolo in una nuova finestra" <%= ACTIVE_STATUS %>
	   onclick="OpenAutoPositionedScrollWindow('<%= GetAmministrazionePath() %>Infoschede/ArticoliMod.asp?ID=<%= id %>#<%= Server.HTMLEncode(codice) %>', 'articolo', 760, 400, true);">
		<%= label %>
	</a>
<%end sub


sub SchedaLink(id, label)%>
	<a href="javascript:void(0);" title="apri la scheda in una nuova finestra" <%= ACTIVE_STATUS %>
	   onclick="OpenAutoPositionedScrollWindow('<%= GetAmministrazionePath() %>Infoschede/SchedeMod.asp?ID=<%= id %>', 'scheda', 760, 400, true);">
		<%= label %>
	</a>
<%end sub


sub SpedizioneLink(id, label)%>
	<a href="javascript:void(0);" title="apri i dati della spedizione in una nuova finestra" <%= ACTIVE_STATUS %>
	   onclick="OpenAutoPositionedScrollWindow('<%= GetAmministrazionePath() %>Infoschede/SpedizioniMod.asp?ID=<%= id %>', 'spedizione', 760, 400, true);">
		<%= label %>
	</a>
<%end sub


sub RitiroLink(id, label)%>
	<a href="javascript:void(0);" title="apri i dati della richesta di ritiro in una nuova finestra" <%= ACTIVE_STATUS %>
	   onclick="OpenAutoPositionedScrollWindow('<%= GetAmministrazionePath() %>Infoschede/RitiriMod.asp?ID=<%= id %>', 'ritiro', 760, 400, true);">
		<%= label %>
	</a>
<%end sub


function GetIdCentroAssistenzaLoggato()
	dim id, cnt_id, sql
	id = 0
	if Session("INFOSCHEDE_CENTRO_ASSISTENZA")<>"" OR Session("INFOSCHEDE_OFFICINA")<>"" then
		sql = "SELECT ut_NextCom_ID FROM tb_utenti WHERE ut_admin_id = " & cIntero(Session("ID_ADMIN"))
		if cIntero(GetValueList(conn, NULL, sql)) > 0 then
			sql = "SELECT cntRel FROM tb_Indirizzario WHERE IDElencoIndirizzi = " & cIntero(GetValueList(conn, NULL, sql))
			cnt_id = cIntero(GetValueList(conn, NULL, sql))
			sql = "SELECT ut_id FROM tb_utenti WHERE ut_NextCom_ID = " & cnt_id
			id = cIntero(GetValueList(conn, NULL, sql))
		else
			sql = "SELECT ag_id FROM gtb_agenti WHERE ag_admin_id = " & cIntero(Session("ID_ADMIN"))
			id = cIntero(GetValueList(conn, NULL, sql))
		end if
	end if
	GetIdCentroAssistenzaLoggato = cIntero(id)
end function



function GetPermessoUtente(ut_id)
	dim sql, id_profilo
	sql = "SELECT riv_profilo_id FROM gtb_rivenditori WHERE riv_id = " & cIntero(ut_id)
	id_profilo = cIntero(GetValueList(conn, NULL, sql))
	select case id_profilo
		case TRASPORTATORI
			GetPermessoUtente = PERMESSO_AR_TRASPORTATORI
		case COSTRUTTORI
			GetPermessoUtente = PERMESSO_AR_COSTRUTTORI
		case CLIENTI_PRIVATI
			GetPermessoUtente = PERMESSO_AR_CLIENTI_PRIVATI
		case CLIENTI_PROFESSIONALI
			GetPermessoUtente = PERMESSO_AR_CLIENTI_PROFESSIONALI
		case RIVENDITORI
			GetPermessoUtente = PERMESSO_AR_RIVENDITORI
		case SUPERVISORI_NEGOZI
			GetPermessoUtente = PERMESSO_AR_SUPERVISORI_NEGOZI
		case else
			GetPermessoUtente = ""
	end select
end function


%>