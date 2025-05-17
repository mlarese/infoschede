<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="../library/CLASS_MAILER.ASP"-->
<!--#INCLUDE FILE="../library/classIndirizzarioLock.asp" -->
<%
dim conn, sql, rs, rsr, rsc, sql_filtri, var, id_age, obj_cnt
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")
set rsr = Server.CreateObject("ADODB.RecordSet")
set rsc = Server.CreateObject("ADODB.RecordSet")
set obj_cnt = new IndirizzarioLock

if Request.ServerVariables("REQUEST_METHOD")="POST" AND request("procedi") <> "" then
	dim url, cnt_id, cliente_id, cliente_cnt_id, cod_ins
	id_age = cIntero(request("ID_AGE"))
	if cIntero(id_age) > 0 then 
		sql = "UPDATE sgtb_schede SET sc_centro_assistenza_id = " & id_age & " WHERE sc_id = " & cIntero(request("ID_SCHEDA"))
		conn.beginTrans
		conn.execute(sql)
		
		conn.commitTrans	' spostato prima dell'invio email, altrimenti nella pagina della scheda inviata non si riesce a recuperare i dati
		
		sql = "SELECT sgtb_schede.sc_cliente_id, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.codiceInserimento" & _
			  " FROM sgtb_schede INNER JOIN gtb_rivenditori ON sgtb_schede.sc_cliente_id = gtb_rivenditori.riv_id " & _
			  " INNER JOIN tb_Utenti ON gtb_rivenditori.riv_id = tb_Utenti.ut_ID INNER JOIN tb_Indirizzario ON tb_Utenti.ut_NextCom_ID = tb_Indirizzario.IDElencoIndirizzi " & _
			  " WHERE sc_id = " & cIntero(request("ID_SCHEDA"))
	    rsc.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
		
		if request("spedisci_mail") <> "" then
			url = GetPageSiteUrl(conn, Session("INFOSCHEDE_ID_PAG_ASSEGNA_CENTRO"), "it") & "&SCHEDAID="&cIntero(request("ID_SCHEDA")) & "&CLIENTEID="&rsc("sc_cliente_id") & "&IDCNT="&rsc("IDElencoIndirizzi") & "&KEY="&rsc("codiceInserimento")
			sql = "SELECT ut_nextCom_id FROM tb_utenti WHERE ut_id = " & id_age
			cnt_id = cIntero(GetValueList(conn, rs, sql))

			CALL SendPageFromAdminToContactExtended(conn, rs, "it", "Assegnazione scheda", url, GetSiteBaseUrl(conn, Session("INFOSCHEDE_ID_PAG_ASSEGNA_CENTRO")), Session("ID_ADMIN"), cnt_id, false)
		end if
		
		'conn.commitTrans
	end if
	%>
	<script language="JavaScript" type="text/javascript">
		opener.location.reload(true);
		window.close();
	</script>
	<%
	response.end
end if


dim Pager
set Pager = new PageNavigator
'--------------------------------------------------------
sezione_testata = "Selezione del centro assistenza" 	
%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 


if Request.ServerVariables("REQUEST_METHOD")="POST" then
	%>
	<div id="content_ridotto">
	<form action="" method="post" id="ricerca" name="ricerca">
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
			<caption>Riepilogo</caption>
			<th>Scheda</th>
			<tr>
				<% sql = "SELECT * FROM sgtb_schede WHERE sc_id = " & cIntero(request("ID_SCHEDA"))
				rsr.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText %>
				<th style="text-align:center;border-top:0px;border-bottom:0px;background:#f4f4f2;">
					<%="n. " & rsr("sc_numero") & " del " & rsr("sc_data_ricevimento")%>
				</th>
				<input type="hidden" value="<%=request("ID_SCHEDA") %>" name="ID_SCHEDA">
				<% rsr.close %>
			</tr>

			<% 
			for each var in request.form
				if Instr(var, "associa_")>0 then
					if request.form(var) = "ASSOCIA" then
						id_age = cIntero(Replace(var, "associa_", ""))
					end if
				end if
			next
			sql = " SELECT ut_nextCom_id FROM tb_Utenti WHERE ut_id = " & cIntero(id_age)
			obj_cnt.LoadFromDB(cInteger(GetValueList(conn, NULL, sql)))
			%>
			<th>Centro assistenza</th>
			<tr>
				<th style="text-align:center;border-top:0px;border-bottom:0px;background:#f4f4f2;">
					<%= obj_cnt("NomeOrganizzazioneElencoIndirizzi")%>
				</th>
				<input type="hidden" value="<%=id_age%>" name="ID_AGE">
			</tr>
			<tr>
				<td class="label" style="text-align:center ;background:#f4f4f2;">
					<%= obj_cnt("email")%>
				</td>
			</tr>
			<th>E-mail</th>
			<tr>
				<td class="label" style="text-align:center; background:#f4f4f2;">
					Invia e-mail di avviso al centro assistenza
				</td>
			</tr>
			<tr>
				<td class="label" style="text-align:center; background:#f4f4f2;">
					<input type="checkbox" class="noBorder" name="spedisci_mail" value="1" checked>
				</td>
			</tr>
			<tr>
				<td class="footer" align="right">
					<input type="submit" class="button" name="procedi" value="PROCEDI" <%= ACTIVE_STATUS %>>
				</td>
			</tr>
		</table>
		
	</form>
	</div>
	<script language="JavaScript" type="text/javascript">
		FitWindowSize(this);
	</script>
	<%
	response.end
end if


'imposta ricerca
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Pager.Reset()
	CALL SearchSession_Reset("cen_")
	if not(request("tutti")<>"") then
		CALL SearchSession_Set("cen_")
	end if
end if

'filtra per nome
if Session("cen_nome")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch_Contatto_Nominativo(conn, Session("cen_nome"))
end if

'filtra per login
if Session("cen_login")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch(Session("cen_login"), "ut_login")
end if

'filtra per indirizzo
if Session("cen_indirizzo")<>"" then
	sql = sql & " AND " & SQL_FullTextSearch_Contatto_Indirizzo(conn, Session("cen_indirizzo"))
end if


sql = " SELECT * FROM gv_agenti " & _
	  " WHERE (1=1) "& sql & _
	  " ORDER BY ag_supervisore DESC, ModoRegistra"
CALL Pager.OpenSmartRecordset(conn, rs, sql, 15)
%>

<div id="content_ridotto">
<form action="" method="post" id="ricerca" name="ricerca">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption>
		<table border="0" cellspacing="0" cellpadding="1" align="right">
			<tr>
				<td style="font-size: 1px; padding-right:1px;" nowrap>
					<input type="submit" name="cerca" value="CERCA" class="button">
					&nbsp;
					<input type="submit" name="tutti" value="VEDI TUTTI" class="button">
				</td>
			</tr>
		</table>
		Opzioni di ricerca
	</caption>
	<tr>
		<th>NOME CONTATTO</th>
		<th>LOGIN CONTATTO</th>
	</tr>
	<tr>
		<td class="content">
			<input type="text" name="search_nome" value="<%= TextEncode(session("cen_nome")) %>" style="width:100%;">
		</td>
		<td class="content">
			<input type="text" name="search_login" value="<%= TextEncode(session("cen_login")) %>" style="width:100%;">
		</td>
	</tr>
	<tr><th colspan="2">INDIRIZZO</th></tr>
	<tr>
		<td class="content" colspan="2">
			<input type="text" name="search_indirizzo" value="<%= replace(session("cen_indirizzo"), """", "&quot;") %>" style="width:100%;">
		</td>
	</tr>
</table>

<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption class="border">Scheda selezionata</caption>
	<tr>
		<% sql = "SELECT * FROM sgtb_schede WHERE sc_id = " & cIntero(request("ID_SCHEDA"))
		rsr.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText %>
		<th style="text-align:center;border-top:0px;border-bottom:0px;background:#f4f4f2;">
			<%="n. " & rsr("sc_numero") & " del " & rsr("sc_data_ricevimento")%>
		</th>
		<input type="hidden" value="<%=rsr("sc_id")%>" name="id_scheda">
		<% rsr.close %>
	</tr>
</table>

<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<caption class="border">Elenco centri assistenza</caption>
	<tr>
		<td>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
				<tr>
					<td class="label_no_width" colspan="3">
						<% if rs.eof then %>
							Nessuna centro assistenza trovato.
						<% else %>
							Trovati n&ordm; <%= Pager.recordcount %> centri assistenza in n&ordm; <%= Pager.PageCount %> pagine
						<% end if %>
					</td>
				</tr>
				<% if not rs.eof then %>
					<tr>
						<th class="L2">NOME</th>
						<th class="l2_center" style="width:70px;">OPERAZIONE</th>
					</tr>
					<%rs.AbsolutePage = Pager.PageNo
					while not rs.eof and rs.AbsolutePage = Pager.PageNo%>
						<tr>
							<td class="content">
								<a href="javascript:void(0);" title="apri scheda del cliente" <%= ACTIVE_STATUS %>
									onclick="OpenAutoPositionedScrollWindow('ClientiGestione.asp?ID=<%= rs("IDElencoIndirizzi") %>', 'cliente', 760, 400, true);">
									<%= ContactFullName(rs) %>
								</a>
							</td>
							<td class="content_center">
								<% dim status 
								status = IIF(rs("ag_id")=cIntero(request("ID_CENTRO")),"disabled","")
								%>
								<input <%=status%> type="submit" class="button_L2 <%=status%>" name="associa_<%=rs("ag_id")%>" value="ASSOCIA">
							</td>
						</tr>
						<% rs.MoveNext
					wend%>
					<tr>
						<td colspan="3" class="footer">
							<table width="100%" cellpadding="0" cellspacing="0">
								<tr>
									<td><% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%></td>
									<td align="right">
										<a class="button" href="javascript:window.close();" title="chiudi la finestra" <%= ACTIVE_STATUS %>>
											CHIUDI</a>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				<% end if %>
			</table>
		</td>
	</tr>
</table>
</form>
</div>
</body>
</html>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>
<% 
rs.close
conn.close
set rs = nothing
set rsr = nothing
set rsc = nothing
set conn = nothing
%>