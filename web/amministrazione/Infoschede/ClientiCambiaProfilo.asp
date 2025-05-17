<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<!--#INCLUDE FILE="Tools_Infoschede.asp" -->
<!--#INCLUDE FILE="../library/classIndirizzarioLock.asp" -->
<%
dim conn, sql, OBJ_contatto, rs
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")



if request("salva")<>"" AND Request.ServerVariables("REQUEST_METHOD")="POST" then
	dim permesso_area_riservata, cnt_id, ut_id
	'cambio il riferimento del profilo
	conn.beginTrans
	sql = " UPDATE gtb_rivenditori SET riv_profilo_id = " & cIntero(request("tfn_pro_id")) & " WHERE riv_id = " & cIntero(request("tfn_riv_id"))
	conn.Execute(sql)
	
	ut_id = cIntero(request("tfn_riv_id"))
	sql = "SELECT ut_nextCom_id FROM tb_utenti WHERE ut_id = " & ut_id
	cnt_id = cIntero(GetValueList(conn, NULL, sql))
	
	set OBJ_contatto = new IndirizzarioLock
	OBJ_contatto.LoadFromDB(cnt_id)
	CaricaCampiEsterni conn, rs, OBJ_contatto, "SELECT * FROM gtb_rivenditori", "riv_id", ut_id
	
	'tolgo il vecchio permesso
	CALL OBJ_contatto.UserAbilitazione_Remove(cnt_id, ut_id, cString(request("old_permesso")))
	
	'aggiungo permesso per nuovo profilo
	permesso_area_riservata = GetPermessoUtente(ut_id)
	CALL OBJ_contatto.UserAbilitazione_Add(cnt_id, ut_id, permesso_area_riservata)
	
	conn.commitTrans
	%>
	<script language="JavaScript" type="text/javascript">
		opener.location.reload(true);
		window.close();
	</script>
	<%
	response.end
end if


%>
<%'--------------------------------------------------------
sezione_testata = "selezione profilo anagrafica" %>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 

%>

<div id="content_ridotto">
<form action="" method="post" id="form2" name="form2">
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
	<input type="hidden" name="tfn_riv_id" value="<%=cIntero(request("RIV_ID"))%>">
	<input type="hidden" name="old_permesso" value="<%=GetPermessoUtente(cIntero(request("RIV_ID")))%>">
	<caption class="border">Modifica profilo</caption>
	<tr>
		<td>
			<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
				<tr>
					<th>profilo</th>
				</tr>
				<tr>
					<td class="content">
						<% sql = "SELECT * FROM gtb_profili WHERE pro_id NOT IN ("&TRASPORTATORI&","&COSTRUTTORI&") ORDER BY pro_nome_it" 
						CALL DropDownAdvanced(conn, sql, "pro_id", "pro_nome_it", "tfn_pro_id", cIntero(request("PRO_ID")), true, "style=""width:100%;""", "", "")%>						
					</td>
				</tr>
				<tr>
					<td colspan="5" class="footer">
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
conn.close
set conn = nothing 
Session("ERRORE") = ""
%>

<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>


