<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<!--#INCLUDE FILE="Tools4DataBase.asp" -->
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->

<% 
'*****************************************************************************************************************
'verifica dei permessi
if not IsLocal() then
	CALL VerificaPermessiUtente(true)
end if
'*****************************************************************************************************************
%>
<html>
<head>
	<title>Amministrazione aggiornamenti database</title>
	<link rel="stylesheet" type="text/css" href="../stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0" onload="window.focus();">
<%
dim conn, rs, field, sql, sql_list, i
set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application(request("ConnString")), "", ""
%>
<form name="form1" id="form1" action="" method="post">
<table width="740" cellspacing="0" cellpadding="0" border="0">
	<caption style="border:0px;">
		<table width="100%" cellspacing="0" cellpadding="0" border="0">
	  		<tr>
				<td align="right" style="padding-right:10px;">
					<a href="javascript:close();" class="menu" name="top">CHIUDI</a>
				</td>
	  		</tr>
		</table>
	</caption>
	<tr>
		<td style="padding-top:4px;">
			<% if Request.ServerVariables("REQUEST_METHOD")="POST" AND request("CODICE_SQL")<>"" then  %>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
					<caption>
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td class="caption">ESECUZIONE CODICE SU connessione "<%= request("ConnString") %>"</td>
								<td align="right" style="font-size: 1px;">
									<a class="button" href="javascript:void(0);" onclick="window.close();">
										CHIUDI
									</a>
								</td>
							</tr>
						</table>
					</caption>
					<% sql = request("CODICE_SQL")
					
					sql_list = split(sql, ";")
					
					conn.beginTrans
					
					for each sql in sql_list
						if trim(sql)<>"" then %>
							<tr>
								<th>ESECUZIONE CODICE</th>
							</tr>
							<tr>
								<td class="content">
									<%= TextHtmlEncode(sql) %>
								</td>
							</tr>
							<% set rs = conn.execute(sql, , adCmdText)
                            if rs.state = adStateOpen then %>
							<tr>
								<td class="content">
									rs.bof = <%= rs.bof %><br>
									rs.recordcount = <%= rs.recordcount %><br>
									rs.eof = <%= rs.eof %><br>
								</td>
							</tr>
                            <% end if %>
							<tr>
								<td class="content_b ok">
									<%if request("ESEGUI")="ESEGUI" then%>
										CODICE ESEGUITO CORRETTAMENTE
									<% else %>
										CODICE VERIFICATO CORRETTAMENTE
									<% end if%>
								</td>
							</tr>
							<%if rs.state = adStateOpen then%>
								<tr>
									<td>
										<table width="100%" border="0" cellspacing="1" cellpadding="0">
											<tr>
												<th class="center ok" nowrap>n. riga</th>
												<%for each Field in rs.Fields%>
													<th>&nbsp;<%= Field.name %></th>
												<%next%>
												<th class="center ok" nowrap>n. riga</th>
											</tr>
											<%i = 1
											while not rs.eof %>
												<tr>
													<td class="content_center ok"><%= i %></td>
													<%for each Field in rs.Fields%>
														<td class="content">
															<%if isNull(Field.value) then%>
																NULL
															<%else%>
																&nbsp;<%= Field.value %>
															<%end if%>
														</td>
													<%next%>
													<td class="content_center ok"><%= i %></td>
												</tr>
												<%i = i + 1
												rs.MoveNext
											wend %>
										</table>
									</td>
								</tr>
							<%else%>
								<tr><td class="content">Istruzione senza risultati</td></tr>
							<%end if
							set rs = nothing
						end if
					next
					
					if request("ESEGUI")="ESEGUI" then
						conn.CommitTrans
					else
						conn.RollbackTrans
					end if %>
				</table>
			<% end if %>
		</td>
	</tr>
</table>
<table width="740" cellspacing="0" cellpadding="0" border="0">
	<tr>
		<td>
			<table cellspacing="1" cellpadding="0" class="tabella_madre">
				<caption>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td class="caption">CODICE DA ESEGUIRE SU connessione "<%= request("ConnString") %>"</td>
							<td align="right" style="font-size: 1px;">
								<a class="button" href="javascript:void(0);" onclick="window.close();">
									ANNULLA
								</a>
							</td>
						</tr>
					</table>
				</caption>
				<tr>
					<th>ESECUZIONE SQL</th>
				</tr>
				<tr>
					<td class="content_center">
						<textarea rows="15" name="CODICE_SQL" id="CODICE_SQL" style="width:100%;"><%= request("CODICE_SQL") %></textarea>
					</td>
				</tr>
				<tr>
					<td class="note">
						Il comando VERIFICA esegue l'sql all'interno di una transazione della quale viene fatto il rollback.<br>
						Il comando ESEGUI esegue l'sql all'interno di una transazione unica.
					</td>
				</tr>
				<tr>
					<td class="footer">
						<input style="width:10%;" type="submit" class="button" name="ESEGUI" value="VERIFICA">
						<input style="width:10%;" type="submit" class="button" name="ESEGUI" value="ESEGUI">
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</body>
</html>

<% 
conn.close
set conn = nothing
%>