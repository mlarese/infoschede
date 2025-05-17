<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->

<% 
dim conn, sql, ID_STAGE, rs_Stage, rs, errore, esito, azione

ID_STAGE = cIntero(request.Querystring("ID_STAGE"))

set conn = Server.CreateObject("ADODB.Connection")
set rs_STAGE = server.CreateObject("ADODB.Recordset")

conn.open Application("DATA_ConnectionString"),"",""


'check dei permessi dell'utente
if NOT ChkPrmPages(ID_STAGE) then
	conn.close
	set conn = nothing %>
<script language="JavaScript">
	window.close()
</script>
<%
end if


sql = "SELECT * FROM tb_pages WHERE id_page=" & ID_STAGE
rs_STAGE.open sql, conn, adOpenStatic, adLockOptimistic

Errore = ""
Esito = ""

if request.Form("AZIONE")<>"" then
	'richiesto cambio/associazione template
	select case request.Form("AZIONE")
		case "AGGANCIA_TEMPLATE"
			'assegna il template alla pagina mantenendolo agganciato
			rs_STAGE("id_template") = cIntero(request.Form("sel_template"))
			rs_STAGE.Update
			Esito = "OK"
			
		case "SGANCIA_COPIA_LAYERS", "COPIA_LAYERS"
			'copia i layers del template nella pagina e sgancia il template
			if cIntero(request.Form("id_template"))>0 then
				conn.beginTrans
				
				CALL Copy_page(conn, request.Form("id_template"), ID_STAGE, true)
                
				rs_STAGE("id_template") = 0
				rs_STAGE.Update
				
				sql = "SELECT (COUNT(*)) AS N_LAY FROM tb_layers WHERE id_pag=" & ID_STAGE
				set rs = conn.execute(sql)
				if rs("N_LAY")>30 then
					'se sforato il numero di layers annulla tutto
					conn.rollbackTrans
					Errore = "Layers non copiati: raggiunto numero massimo di layers"
					Esito = ""
				else
					conn.commitTrans
					Esito = "OK"
				end if
			else
				Esito = ""
				Errore = "Errore nella copia dei layers."
			end if
			
		case "SGANCIA_TEMPLATE"
			'toglie template dalla pagina
			rs_STAGE("id_template") = 0
			rs_STAGE.Update
			Esito = "OK"
	end select
	if Esito = "OK" then
		'aggiorna data di modifica paginasito
		CALL UpdateDataModifica(conn, ID_STAGE)
	end if
	
	'agiorna recordset per recuperare modifiche
	rs_STAGE.close
	sql = "SELECT * FROM tb_pages WHERE id_page=" & ID_STAGE
	rs_STAGE.open sql, conn, adOpenStatic, adLockOptimistic
	session("ERRORE") = errore
end if
%>

<%'--------------------------------------------------------
if request("nextmail")<>"" then
	sezione_testata = "Gestione template della next-email"
else
	sezione_testata = "Gestione siti - indice delle pagine - template della pagina"
end if %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<script language="JavaScript">
	function aggancia(){
		if (((form1.sel_template.selectedIndex!=0) || (<%=cInteger(rs_STAGE("id_template"))%>) ) && 
		    (form1.sel_template.options[form1.sel_template.selectedIndex].value!=<%=cInteger(rs_STAGE("id_template"))%>)){
				form1.azione.value='AGGANCIA_TEMPLATE';
				form1.submit();
		}
	}
	
	function copia(){
		if (form1.sel_template.selectedIndex!=0){
			message = 'Non verr&agrave; fatta alcuna associazione con i template.\n';
			message += 'I layer appartenenti al template verranno copiati nella pagina e saranno modificabili.\n';
			message += 'ATTENZIONE: la copia dei layers non e\' reversibile.\n';
			message += 'Copiare i layers?';
			if (confirm(message)){
				form1.azione.value='COPIA_LAYERS'; 
				form1.id_template.value=form1.sel_template.options[form1.sel_template.selectedIndex].value; 
				form1.submit();
			}
		}
	}
			
	function sgancia(){
		form1.azione.value='SGANCIA_TEMPLATE';
		form1.submit();
	}
	
	
	function sganciacopia(){
		message = 'L\'associazione con il template sar&agrave; eliminata.\n';
		message += 'I layer appartenenti al template verranno copiati nella pagina e saranno modificabili.\n';
		message += 'La pagina non sar&agrave; riconducibile ad alcun template.\n';
		message += 'ATTENZIONE: la copia dei layers non e\' reversibile.\n';
		message += 'Sganciare la pagina e copiare i layers?';
		if (confirm(message)){
			form1.azione.value='SGANCIA_COPIA_LAYERS'; 
			form1.submit();
		}
	}
</script>
<div id="content_ridotto">
	<%if Esito <>"" then%>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
			<caption class="border ok">Operazione eseguita correttamente</caption>
			<tr>
				<td colspan="2" class="content">
					<%select case request.Form("AZIONE")
						case "AGGANCIA_TEMPLATE"%>
							<br>
							<b>Template agganciato correttamente.</b><br>
							<br>
							Ogni modifica eseguita sul template verr&agrave; riportata anche sulla pagina corrente.<br>
							I layers appatenenti al template non sono modificabili dall'interno della pagina.<br>
							<br>
						<%case "COPIA_LAYERS"%>
							<br>
							<b>Layers copiati correttamente.</b><br>
							<br>
							Tutti layers del template sono stati copiati nella pagina.<br>
							La pagina non &egrave; riconducibile ad alcun template, perci&ograve; tutti i layers sono 
							modificabili e nessuna modifica fatta al template verr&agrave; riportata nella pagina.<br>
							<br>
						<%case "SGANCIA_COPIA_LAYERS"%>
							<br>
							<b>Pagina sganciata correttamente.</b><br>
							<b>Layers copiati correttamente.</b><br>
							<br>
							L'associazione con il template &egrave; stata eliminata, ed i layers del template 
							sono stati copiati nella pagina.<br>
							La pagina non &egrave; riconducibile ad alcun template, perci&ograve; tutti i layers sono 
							modificabili e nessuna modifica fatta al template verr&agrave; riportata nella pagina.<br>
							<br>
						<%case "SGANCIA_TEMPLATE"%>
							<br>
							<b>Pagina sganciata correttamente.</b><br>
							<br>
							L'associazione con il template &egrave; stata eliminata.<br>
							Tutti i layers appatenenti al template non sono pi&uacute; visibili nella pagina corrente.<br>
							<br>
					<%end select%>	
				</td>
			</tr>
			<tr>
				<td colspan="2" class="note">
					Questa finestra si chiuder&agrave; automaticamente tra 10 secondi.
				</td>
			</tr>
		</table>
		<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
			<caption class="border">Pubblica la pagina</caption>
			<tr>
				<td class="content" style="width:65%;">
					Pubblica la pagina di lavoro appena modificata.
				</td>
				<td class="content_center">
					<% sql = "SELECT id_pagDYN_" & rs_stage("lingua") & " FROM tb_paginesito WHERE id_pagSTAGE_" & rs_stage("lingua") & "=" & rs_stage("id_page") %>
					<a href="SitoPagineCopia.asp?ID_S=<%= rs_stage("id_page") %>&ID_D=<%= GetValueList(conn, NULL, sql) %>&lingua=<%= rs_stage("lingua") %>&nome_lingua=<%= GetNomeLingua(rs_stage("lingua")) %>&azione=PUBBLICA&conferma=OK" class="button_L2_block" 
					   title="Pubblica la pagina sostituendo quella attualmente visibile al pubblico." <%= ACTIVE_STATUS %>>
						PUBBLICA LA PAGINA
					</a>
				</td>
			</tr>
		</table>
		<script language="JavaScript">
			<% if request("nextmail")="" then %>
				opener.location.reload(true);
			<% else %>
				opener.SetPreview( <%= ID_STAGE %>);
			<% end if %>
			//window.setTimeout("close();", 10000);
		</script>
	<%end if%>
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<form action="" method="post" id="form1" name="form1">
		<input type="hidden" name="azione" value="">
		<input type="hidden" name="id_template" value="<%=rs_STAGE("id_template")%>">
			<caption class="border">
				<% if request("nextmail")<>"" then %>
					Gestione template della next-email
				<% else %>
					Gestione template della pagina "<%= rs_STAGE("nomepage") %>"
				<% end if %>
			</caption>
			<tr>
				<td rowspan="4" class="content" style="width:65%;">
					<%sql = "SELECT (0) AS id_page, ('Template vuoto (nessun template associato)') AS NAME, (0) AS ordine FROM tb_pages UNION " &_
							QryElencoTemplate("", true)
					CALL dropDown(conn, sql, "id_page", "name", "sel_template", cInteger(rs_STAGE("id_template")), TRUE, " style=""width:95%;"" ", LINGUA_ITALIANO)%>
				</td>
				<td class="content_center">
					<a href="#" class="button_L2_block" onclick="aggancia();"
					   title="Associa alla pagina il template selezionato." <%= ACTIVE_STATUS %>>
						<%if cInteger(rs_STAGE("id_template"))>0 then%>
							CAMBIA TEMPLATE
						<% else %>
							AGGANCIA A TEMPLATE
						<%end if%>
					</a>
				</td>
			</tr>
			<tr>
				<td class="content_center">
					<%if cInteger(rs_STAGE("id_template"))>0 then%>
						<a class="button_l2_block_disabled"
						   title="Impossibile eseguire la copia: sganciare prima il template o usare la funzione &quot;Sgancia e copia&quot;" <%= ACTIVE_STATUS %>>
							COPIA DA TEMPLATE
						</a>
					<% else %>
						<a href="#" class="button_L2_block" onclick="copia();"
						   title="Copia il contentuto del template selezionato all'interno della pagina." <%= ACTIVE_STATUS %>>
							COPIA DA TEMPLATE
						</a>
					<% end if %>
				</td>
			</tr>
			<tr>
				<td class="content_center">
					<%if cInteger(rs_STAGE("id_template"))>0 then%>
						<a href="#" class="button_L2_block" onclick="sgancia();"
					   	   title="Associa alla pagina il template vuoto." <%= ACTIVE_STATUS %>>
							SGANCIA DA TEMPLATE 
						</a>
					<%else%>
						<a class="button_L2_block_disabled" title="Impossibile sganciare il template. Nessun template selezionato." <%= ACTIVE_STATUS %>>
							SGANCIA DA TEMPLATE 
						</a>
					<%end if%>
					
				</td>
			</tr>
			<tr>
				<td class="content_center">
					<%if cInteger(rs_STAGE("id_template"))>0 then%>
						<a href="#" class="button_L2_block" onclick="sganciacopia();"
					   	   title="Copia il contenuto del template nella pagina associandoci il template vuoto." <%= ACTIVE_STATUS %>>
							SGANCIA E COPIA
						</a>
					<%else%>
						<a class="button_L2_block_disabled" title="Impossibile sganciare e copiare. Nessun template selezionato." <%= ACTIVE_STATUS %>>
							SGANCIA E COPIA
						</a>
					<%end if%>
					
				</td>
			</tr>
			<tr>
				<td class="footer" colspan="2">
					<input type="button" class="button" name="chiudi" value="CHIUDI" onclick="window.close();">
				</td>
			</tr>
		</form>
	</table>
</div>
</body>
</html>

<% rs_stage.close
conn.close
set rs_stage = nothing
set conn = nothing %>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>