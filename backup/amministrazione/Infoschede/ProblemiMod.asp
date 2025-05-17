<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if (request("salva")<>"" OR request("salva_indietro")<>"") AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("ProblemiSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura, data
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione guasti - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Problemi.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, rsa, i
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsa = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, Session("ELENCO_PROBLEMI_SQL"), "prb_id", "ProblemiMod.asp")
end if

sql = "SELECT * FROM sgtb_problemi WHERE prb_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
%>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica dati del guasto</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= cIntero(request("ID")) %>&goto=PREVIOUS" title="guasto precedente">
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= cIntero(request("ID")) %>&goto=NEXT" title="guasto successivo">
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="6">DATI PRINCIPALI</th></tr>
		<script language="JavaScript" type="text/javascript">
			function natura_guasto(){
				var isProtetto = document.getElementById('chk_prb_riscontrato_true');
				var tab = document.getElementById('protetto')
				if (!(isProtetto.checked)){
					tab.style.visibility = "visible";
					tab.style.display = "block";
				}
				else{
					tab.style.visibility = "hidden";
					tab.style.display = "none";
				}
			}
		</script>
		<tr>
			<td class="label">guasto:</td>
			<td class="content" colspan="5">
				<table border="0" cellspacing="0" cellpadding="0" align="left">
					<tr>
						<td><input class="noBorder riscontrato" type="radio" name="prb_riscontrato" id="chk_prb_riscontrato_true" value="1" <%=chk(cBoolean(rs("prb_riscontrato"), false))%> onClick="natura_guasto()"></td>
						<td width="60%" style="padding-left:5px;">riscontrato</td>
						<td><input class="noBorder segnalato" type="radio" name="prb_riscontrato" id="chk_prb_riscontrato_false" value="" <%=chk(cBoolean(not rs("prb_riscontrato"), false))%> onClick="natura_guasto()"></td>
						<td width="30%" style="padding-left:5px;">segnalato</td>
					</tr>
				</table>
			</td>
		</tr>
		
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
			<% 	if i = 0 then %>
				<td class="label" style="width:18%;" rowspan="<%= ubound(Application("LINGUE"))+1 %>">titolo:</td>
			<% 	end if %>
				<td class="content" colspan="5">
					<img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0">
					<input type="text" class="text" name="tft_prb_nome_<%= Application("LINGUE")(i) %>" value="<%= rs("prb_nome_"& Application("LINGUE")(i)) %>" maxlength="255" size="90">
					<% 	if i = 0 then %>(*)<% end if %>
				</td>
			</tr>
		<% next %>
		<tr>
			<td class="label">ordine:</td>
			<td class="content" colspan="5">
				<input type="text" class="text" name="tfn_prb_ordine" value="<%= rs("prb_ordine") %>" maxlength="8" size="8">
			</td>
		</tr>
		<tr>
			<td class="label" style="width:20%;">visibile:</td>
			<td class="content" colspan="5">
				<input checked type="checkbox" class="noBorder" name="prb_visibile" value="1" <% if rs("prb_visibile")<>"" then %> checked <% end if %>>
			</td>
		</tr>
		<tr>
			<td class="label" rowspan="2">immagini</td>
			<td class="label">thumb:</td>
			<td class="content" colspan="4">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_prb_img_th", rs("prb_img_th"), "width:350px;", false) %>
			</td>
		</tr>
		<tr>
			<td class="label_no_width">zoom:</td>
			<td class="content" colspan="4">
				<% CALL WriteFilePicker_Input(Application("AZ_ID"), "images", "form1", "tft_prb_img_zo", rs("prb_img_zo"), "width:350px;", false) %>
			</td>
		</tr>		
		<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
			<tr>
				<% if i=lbound(Application("LINGUE")) then %>
					<td class="label" rowspan="<%=ubound(Application("LINGUE")) + 1%>">descrizione:</td>
				<% end if %>
				<td class="content" colspan="5">
					<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
						<tr>
							<td width="5%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
							<td><textarea style="width:100%;" rows="3" name="tft_prb_descrizione_<%= Application("LINGUE")(i)%>"><%= rs("prb_descrizione_" & Application("LINGUE")(i)) %></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		<%next %>
		
		<tr><th colspan="6">ASSOCIAZIONI</th></tr>
		<% sql = "SELECT * FROM gtb_profili ORDER BY pro_nome_it"
		   if cString(GetValueList(conn, NULL, sql)) <> "" then %>
			<tr>
				<td class="label" style="width:20%;">profili:</td>
				<td class="content" colspan="5">
					<% dim rsp
					   set rsp = Server.CreateObject("ADODB.RecordSet")
					   sql = " SELECT * FROM gtb_profili LEFT JOIN srel_problemi_profili " & _
							 " ON gtb_profili.pro_id = srel_problemi_profili.rpp_profilo_id AND srel_problemi_profili.rpp_problema_id = " & cIntero(request("ID")) & _
							 " WHERE gtb_profili.pro_id NOT IN ("&SUPERVISORI_NEGOZI&","&TRASPORTATORI&","&COSTRUTTORI&") " & _
							 " ORDER BY pro_nome_it"
					   CALL Write_Relations_Checker(conn, rsp, sql, 4, "pro_id", "pro_nome_it", "rpp_id", "profili_associati")%>
				</td>
			</tr>
		<% end if %>
		
		<tr><th class="L2" colspan="6">marche / categorie</th></tr>
		<tr>
			<td class="label" style="width:20%;">&nbsp;</td>
			<td colspan="5">
			<% sql = " SELECT *, ISNULL((SELECT mar_nome_it FROM gtb_marche WHERE mar_id=rpm_marchio_id),'Tutte le marche') AS MARCA, " & _
					 " ISNULL((SELECT tip_nome_it FROM gtb_tipologie WHERE tip_id=rpm_tipologia_id),'Tutti le categorie') AS TIPOLOGIA " & _
					 " FROM srel_problemi_mar_tip " & _
					 " WHERE rpm_problema_id = " & cIntero(request("ID")) & _
					 " ORDER BY MARCA "
			rsp.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
			%>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<td class="label" style="width:40%">
							<% if rsp.eof then %>
								Nessun record associato.
							<% else %>
								Trovati n&ordm; <%= rsp.recordcount %> record
							<% end if %>
						</td>
						<td colspan="3" class="content_right" style="padding-right:0px;">
							<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la selezione di un nuovo modello associato" <%= ACTIVE_STATUS %>
							   onclick="OpenAutoPositionedScrollWindow('ProblemiMarTipGestione.asp?IDPRB=<%=request("ID")%>', 'ProbMarTip', 530, 525, true)">
								AGGIUNGI MARCA/TIPOLOGIA
							</a>
						</td>
					</tr>
					<% if not rsp.eof then %>
						<tr>
							<th class="L2">marca</th>
							<th class="L2">categoria</th>
							<th colspan="2" class="l2_center" width="15%">operazioni</th>
						</tr>
						<% while not rsp.eof %>
							<tr>
								<td class="<%=IIF(cIntero(rsp("rpm_marchio_id"))=0,"content_b","content")%>"><%=rsp("MARCA")%></td>
								<td class="<%=IIF(cIntero(rsp("rpm_tipologia_id"))=0,"content_b","content")%>"><%=rsp("TIPOLOGIA")%></td>
								<!--<td class="content_center">
									<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione dell'associatzione tra problema e marca/categoria" <%= ACTIVE_STATUS %>
									   onclick="OpenAutoPositionedScrollWindow('ProblemiMarTipGestione.asp?ID=<%=rsp("rpm_id")%>', 'ProbMarTip', 530, 525, true)">
										MODIFICA
									</a>
								</td>-->
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione dell'associatzione tra problema e marca/categoria" <%= ACTIVE_STATUS %>
									   onclick="OpenDeleteWindow('PROBLEMI_MAR_TIP','<%= rsp("rpm_id") %>');">
										CANCELLA
									</a>
								</td>
							</tr>
							<%rsp.movenext
						wend 
					end if
					rsp.close 
					%>
					<tr>
						<td class="content" colspan="4">&nbsp;</td>
					</tr>
				</table>
			</td>
		</tr>
		
		<tr><th class="L2" colspan="6">modelli</th></tr>
		<tr>
			<td class="label" style="width:20%;">&nbsp;</td>
			<td colspan="5">
				<% dim lista_id
				sql = " SELECT * FROM gv_articoli INNER JOIN srel_problemi_articoli " & _
						 " ON gv_articoli.rel_id = srel_problemi_articoli.rpa_articolo_rel_id " & _
						 " WHERE rpa_problema_id = " & cIntero(request("ID")) & " ORDER BY art_nome_it "
				rsp.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText 
				sql = Replace(sql,"*","rel_id")
				lista_id = cString(GetValueList(conn, NULL, sql))
				lista_id = Replace(lista_id, " ", "")
				
				%>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<td class="label" colspan="2" style="width:60%">
							<% if rsp.eof then %>
								Nessun modello associato.
							<% else %>
								Trovati n&ordm; <%= rsp.recordcount %> record
							<% end if %>
						</td>
						<td colspan="3" class="content_right" style="padding-right:0px;">
							<a class="button_L2" href="javascript:void(0)" title="apre in una nuova finestra la selezione di un nuovo modello associato" <%= ACTIVE_STATUS %>
							   onclick="OpenAutoPositionedScrollWindow('ArticoliSeleziona.asp?IDPRB=<%=request("ID")%>&TYPE=M&EXCLUDE_IDS=<%=lista_id%>', 'ProbArt', 530, 525, true)">
								AGGIUNGI MODELLO
							</a>
						</td>
					</tr>
					<% if not rsp.eof then %>
						<tr>
							<th class="L2" width="22%">categoria</th>
							<th class="L2" width="10%">codice</th>
							<th class="L2">nome</th>
							<th colspan="2" class="l2_center" width="10%">operazioni</th>
						</tr>
						<% while not rsp.eof %>
							<tr>
								<td class="content"><%= categorie.NomeCompleto(rsp("art_tipologia_id")) %></td>
								<td class="content"><%= rsp("art_cod_int")%></td>
								<td class="content">
									<% CALL ArticoloLink(rsp("art_id"), rsp("art_nome_it"), rsp("art_cod_int")) %>
									<% if rsp("art_varianti") then %>
										<%= ListValoriVarianti(conn, rsa, rsp("rel_id")) %>
									<% else %>
										&nbsp;
									<% end if %>
								</td>
								<td class="content_center">
									<a class="button_L2" href="javascript:void(0);" title="apre in una nuova finestra la procedura di cancellazione dell'associatzione tra il problema e il modello" <%= ACTIVE_STATUS %>
									   onclick="OpenDeleteWindow('PROBLEMI_ARTICOLI','<%= rsp("rpa_id") %>');">
										CANCELLA
									</a>
								</td>
							</tr>
							<%rsp.movenext
						wend 
					end if%>
				</table>
			</td>
		</tr>
			
			
		<script language="JavaScript" type="text/javascript">
			function abilita(){
				var modalitaEasy = document.getElementById('chk_prb_modalita_easy_true');
				<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
					var tab = document.getElementById('tft_prb_avviso_per_conferma_<%= Application("LINGUE")(i)%>')
					DisableControl(tab, modalitaEasy.checked)
				<%next%>
			}
			
		</script>
		
		<tr>
			<td class="content" colspan="6" style="padding:0px;">
				<table id="protetto" border="0" cellspacing="1" cellpadding="0" align="left" style="width:100%;">
					<tr><th colspan="6">CARATTERISTICHE DEL GUASTO</th></tr>
					<tr>
						<td class="label" style="width:18%;">modalit&agrave;:</td>
						<td class="content" colspan="5">
							<table border="0" cellspacing="0" cellpadding="0" align="left">
								<tr>
									<td><input class="noBorder" type="radio" name="prb_modalita_easy" id="chk_prb_modalita_easy_true" value="1" <%=chk(cBoolean(rs("prb_modalita_easy"), true))%> onClick="abilita()"></td>
									<td width="60%"><% WriteColor(MODALITA_EASY_COLOR)%><%=MODALITA_EASY%></td>
									<td><input class="noBorder" type="radio" name="prb_modalita_easy" id="chk_prb_modalita_easy_false" value="" <%=chk(not cBoolean(rs("prb_modalita_easy"), false))%> onClick="abilita()"></td>
									<td width="30%"><% WriteColor(MODALITA_NON_EASY_COLOR)%><%=MODALITA_NON_EASY%></td>
								</tr>
							</table>
						</td>
					</tr>
					
					<tr><th class="L2" colspan="6">AVVISO PER CONFERMA</th></tr>
					<%for i=lbound(Application("LINGUE")) to ubound(Application("LINGUE"))%>
						<tr>
							<td class="content" colspan="6">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
									<tr>
										<td width="4%" valign="top"><img src="../grafica/flag_<%= Application("LINGUE")(i) %>.jpg" width="26" height="15" alt="" border="0"></td>
										<td><textarea style="width:100%;" rows="10" id="tft_prb_avviso_per_conferma_<%= Application("LINGUE")(i)%>" name="tft_prb_avviso_per_conferma_<%= Application("LINGUE")(i)%>"><%= rs("prb_avviso_per_conferma_" & Application("LINGUE")(i)) %></textarea></td>
									</tr>
								</table>
							</td>
						</tr>
					<%next %>
				</table>
			</td>
		</tr>
		
		<tr>
			<% CALL Form_DatiModifica(conn, rs, "prb_") %>
			<td class="footer" colspan="6">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="salva" value="SALVA">
				<input type="submit" class="button" name="salva_indietro" value="SALVA & TORNA ALL'ELENCO">
			</td>
		</tr>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>

<script language="JavaScript" type="text/javascript">
	abilita();		
	natura_guasto();
</script>

<% rs.close
conn.close
set rs = nothing
set conn = nothing%>
