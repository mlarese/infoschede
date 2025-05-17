<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede_Const.asp" -->
<!--#INCLUDE FILE="Tools_Infoschede_Categorie.asp" -->
<%
dim Pager, conn, sql
set Pager = new PageNavigator
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")


if request("salva")<>"" AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	dim marche, mar, tipologie, tip
	if request("chk_marche")<>"" AND request("chk_tipologie")<>"" then
		marche = Split(request("chk_marche"),",")
		tipologie = Split(request("chk_tipologie"), ",")
		
		conn.beginTrans
		for each mar in marche
			mar = cIntero(Trim(mar))
			for each tip in tipologie
				tip = cIntero(Trim(tip))
				if mar > 0 AND tip > 0 then
					sql = " DELETE FROM srel_problemi_mar_tip WHERE rpm_problema_id="&cIntero(request("IDPRB"))&_
						  " AND rpm_marchio_id="&mar&" AND rpm_tipologia_id="&tip
					conn.Execute(sql)
					sql = " INSERT INTO srel_problemi_mar_tip(rpm_problema_id,rpm_marchio_id,rpm_tipologia_id) " & _
						  " VALUES ("&cIntero(request("IDPRB"))&", "&mar&", "&tip&")"
					conn.Execute(sql)
				end if
			next
		next
		conn.CommitTrans
		%>
		<script language="JavaScript" type="text/javascript">
			opener.document.form1.submit();
			window.close();
		</script>
		<%
	else
		Session("ERRORE") = "ERRORE nell'inserimento dei dati."
	end if


	' sql = " SELECT rpm_id FROM srel_problemi_mar_tip " & _
		  ' " WHERE rpm_problema_id="&cIntero(request("IDPRB"))&" AND rpm_marchio_id="&cIntero(request("tfn_rpm_marchio_id"))& _
		  ' "		AND rpm_tipologia_id="&cIntero(request("tfn_rpm_tipologia_id"))
	' if CIntero(GetValueList(conn,NULL,sql))>0 then
		' Session("ERRORE") = "ATTENZIONE! Coppia marca/categoria gi&agrave; inserita."
	' else
		' conn.beginTrans
		' if cIntero(request("ID"))>0 then
			' sql = "DELETE FROM srel_problemi_mar_tip WHERE rpm_id = " & cIntero(request("ID"))
			' conn.Execute(sql)
		' end if
		' sql = " INSERT INTO srel_problemi_mar_tip(rpm_problema_id,rpm_marchio_id,rpm_tipologia_id) " & _
			  ' " VALUES (" & cIntero(request("IDPRB")) & ", " & cIntero(request("tfn_rpm_marchio_id")) & ", " & cIntero(request("tfn_rpm_tipologia_id")) & ") "

		' conn.Execute(sql)
		' conn.CommitTrans
		%>
		<script language="JavaScript" type="text/javascript">
			//opener.document.form1.submit();
			//window.close();
		</script>
		<%
	'end if
end if


dim rs, ID, IDPRB, marchio, tipologia
set rs = Server.CreateObject("ADODB.RecordSet")

ID = cIntero(request("ID"))
IDPRB = cIntero(request("IDPRB"))
' marchio = cIntero(request("tfn_rpm_marchio_id"))
' tipologia = cIntero(request("tfn_rpm_tipologia_id"))

' if ID > 0 then
	' sql = "SELECT * FROM srel_problemi_mar_tip WHERE rpm_id = " & ID
	' rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	' IDPRB = cIntero(rs("rpm_problema_id"))
	' marchio = cIntero(rs("rpm_marchio_id"))
	' tipologia = cIntero(rs("rpm_tipologia_id"))
	' rs.close
' end if


%>
<%'--------------------------------------------------------
sezione_testata = "selezione marchio / tipologia" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

%>

<div id="content_ridotto">
<form action="" method="post" id="form2" name="form2">
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<input type="hidden" name="ID" value="<%=ID%>">
	<input type="hidden" name="IDPRB" value="<%=IDPRB%>">
	<caption>Inserisci associazione</caption>
	<tr>
		<td>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
				<tr>
					<th style="width:50%;">marca</th>
					<th>categoria</th>
				</tr>
				<script language="JavaScript" type="text/javascript">
					function AllMarche() {
						if (form2.all_marche.checked){
							for(var i=0; i < form2.elements.length; i++)
								if (form2.elements(i).id.substring(0, 11) == "chk_mar_id_" && !form2.elements(i).checked)
									form2.elements(i).click();
						}
						else{
							for(var i=0; i < form2.elements.length; i++)
								if (form2.elements(i).id.substring(0, 11) == "chk_mar_id_" && form2.elements(i).checked)
									form2.elements(i).click();			
						}
					}
					
					function AllCategorie() {
						if (form2.all_tipologie.checked){
							for(var i=0; i < form2.elements.length; i++)
								if (form2.elements(i).id.substring(0, 11) == "chk_tip_id_" && !form2.elements(i).checked)
									form2.elements(i).click();
						}
						else{
							for(var i=0; i < form2.elements.length; i++)
								if (form2.elements(i).id.substring(0, 11) == "chk_tip_id_" && form2.elements(i).checked)
									form2.elements(i).click();			
						}
					}
				</script>
				<tr>
					<td class="content_b">						
						<input type="checkbox" class="noBorder" name="all_marche" id="all_marche" value="1" onclick="AllMarche()">
						SELEZIONA TUTTE LE MARCHE
					</td>
					<td class="content_b">
						<input type="checkbox" class="noBorder" name="all_tipologie" id="all_tipologie" value="1" onclick="AllCategorie()">
						SELEZIONA TUTTE LE CATEGORIE
					</td>
				</tr>
				<tr>
					<th colspan="2" style="font-size:1px;">&nbsp;</tr>
				</tr>
				<tr>
					<td class="content">
						<% sql = "SELECT * FROM gtb_marche WHERe mar_codice NOT LIKE 'default' ORDER BY mar_nome_it" 
						'CALL DropDownAdvanced(conn, sql, "mar_id", "mar_nome_it", "tfn_rpm_marchio_id", marchio, false, "style=""width:100%;""", "Tutte le marche", "Nessuna marca trovata")
						rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
						
						while not rs.eof
							%>
							<span style="width:100%;">
								<input type="checkbox" class="noBorder" name="chk_marche" id="chk_mar_id_<%=rs("mar_id")%>" value="<%=rs("mar_id")%>">
								<%=rs("mar_nome_it")%>
							</span>
							<%
							rs.moveNext
						wend
						rs.close
						%>						
					</td>
					<td class="content">
						<% 'CALL DropDownAdvanced(conn, cat_modelli.QueryElenco(false, ""), "tip_id", "NAME", "tfn_rpm_tipologia_id", tipologia, false, "style=""width:100%;""", "Tutte le categorie", "Nessuna categoria trovata")
						sql = cat_modelli.QueryElenco(false, " TIP_L0.tip_codice NOT LIKE 'categoria_default' ANd TIP_L0.tip_codice NOT LIKE 'MODELLI' ")
						rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	
						while not rs.eof
							%>
							<span style="width:100%;">
								<input type="checkbox" class="noBorder" name="chk_tipologie" id="chk_tip_id_<%=rs("tip_id")%>" value="<%=rs("tip_id")%>">
								<%=rs("tip_nome_it")%>
							</span>
							<%
							rs.moveNext
						wend
						rs.close
						%>									
					</td>
				</tr>
				<tr>
					<td colspan="2" class="footer">
						<table width="100%" cellpadding="0" cellspacing="0">
							<tr>
								<td style="font-size: 1px; padding-right:1px; text-align:right;" nowrap>
									<input type="submit" name="salva" value="SALVA" class="button">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</form>
</table>
</div>
</body>
</html>

<%
set rs = nothing
conn.close
set conn = nothing 
Session("ERRORE") = ""
%>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>


