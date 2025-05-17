<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<% Server.ScriptTimeout = 100000 %>
<!--#INCLUDE FILE="Tools4DataBase.asp" -->
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<% 
'*****************************************************************************************************************
'verifica dei permessi
CALL VerificaPermessiUtente(true)
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
'imposta elenco di schemi da visualizzare
dim SchemaType(40), SchemaName(40), i, Conn, rs, field, j
SchemaType(0) = adSchemaAsserts
SchemaType(1) = adSchemaCatalogs
SchemaType(2) = adSchemaCharacterSets
SchemaType(3) = adSchemaCheckConstraints
SchemaType(4) = adSchemaCollations
SchemaType(5) = adSchemaColumnPrivileges
SchemaType(6) = adSchemaColumns
SchemaType(7) = adSchemaColumnsDomainUsage
SchemaType(8) = adSchemaConstraintColumnUsage
SchemaType(9) = adSchemaConstraintTableUsage
SchemaType(10) = adSchemaCubes
SchemaType(11) = adSchemaDBInfoKeywords
SchemaType(12) = adSchemaDBInfoLiterals
SchemaType(13) = adSchemaDimensions
SchemaType(14) = adSchemaForeignKeys
SchemaType(15) = adSchemaHierarchies
SchemaType(16) = adSchemaIndexes
SchemaType(17) = adSchemaKeyColumnUsage
SchemaType(18) = adSchemaLevels
SchemaType(19) = adSchemaMeasures
SchemaType(20) = adSchemaMembers
SchemaType(21) = adSchemaPrimaryKeys
SchemaType(22) = adSchemaProcedureColumns
SchemaType(23) = adSchemaProcedureParameters
SchemaType(24) = adSchemaProcedures
SchemaType(25) = adSchemaProperties
SchemaType(26) = adSchemaProviderSpecific
SchemaType(27) = adSchemaProviderTypes
SchemaType(28) = adSchemaReferentialConstraints
SchemaType(29) = adSchemaSchemata
SchemaType(30) = adSchemaSQLLanguages
SchemaType(31) = adSchemaStatistics
SchemaType(32) = adSchemaTableConstraints
SchemaType(33) = adSchemaTablePrivileges
SchemaType(34) = adSchemaTables
SchemaType(35) = adSchemaTranslations
SchemaType(36) = adSchemaTrustees
SchemaType(37) = adSchemaUsagePrivileges
SchemaType(38) = adSchemaViewColumnUsage
SchemaType(39) = adSchemaViews
SchemaType(40) = adSchemaViewTableUsage

SchemaName(0) = "adSchemaAsserts"
SchemaName(1) = "adSchemaCatalogs"
SchemaName(2) = "adSchemaCharacterSets"
SchemaName(3) = "adSchemaCheckConstraints"
SchemaName(4) = "adSchemaCollations"
SchemaName(5) = "adSchemaColumnPrivileges"
SchemaName(6) = "adSchemaColumns"
SchemaName(7) = "adSchemaColumnsDomainUsage"
SchemaName(8) = "adSchemaConstraintColumnUsage"
SchemaName(9) = "adSchemaConstraintTableUsage"
SchemaName(10) = "adSchemaCubes"
SchemaName(11) = "adSchemaDBInfoKeywords"
SchemaName(12) = "adSchemaDBInfoLiterals"
SchemaName(13) = "adSchemaDimensions"
SchemaName(14) = "adSchemaForeignKeys"
SchemaName(15) = "adSchemaHierarchies"
SchemaName(16) = "adSchemaIndexes"
SchemaName(17) = "adSchemaKeyColumnUsage"
SchemaName(18) = "adSchemaLevels"
SchemaName(19) = "adSchemaMeasures"
SchemaName(20) = "adSchemaMembers"
SchemaName(21) = "adSchemaPrimaryKeys"
SchemaName(22) = "adSchemaProcedureColumns"
SchemaName(23) = "adSchemaProcedureParameters"
SchemaName(24) = "adSchemaProcedures"
SchemaName(25) = "adSchemaProperties"
SchemaName(26) = "adSchemaProviderSpecific"
SchemaName(27) = "adSchemaProviderTypes"
SchemaName(28) = "adSchemaReferentialConstraints"
SchemaName(29) = "adSchemaSchemata"
SchemaName(30) = "adSchemaSQLLanguages"
SchemaName(31) = "adSchemaStatistics"
SchemaName(32) = "adSchemaTableConstraints"
SchemaName(33) = "adSchemaTablePrivileges"
SchemaName(34) = "adSchemaTables"
SchemaName(35) = "adSchemaTranslations"
SchemaName(36) = "adSchemaTrustees"
SchemaName(37) = "adSchemaUsagePrivileges"
SchemaName(38) = "adSchemaViewColumnUsage"
SchemaName(39) = "adSchemaViews"
SchemaName(40) = "adSchemaViewTableUsage"

set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application(request("ConnString")), "", ""
%>
<table width="740" cellspacing="0" cellpadding="0" border="0" class="tabella_madre">
	<caption style="border:0px;">
		<a style="float:right;" href="javascript:close();" class="menu" name="top">CHIUDI</a>
	</caption>
	<tr><th>SELEZIONARE IL TIPO DI ANALISI DA ESEGUIRE</th></tr>
	<tr>
		<td class="content">
			<ul style="margin-top:0px; margin-bottom:0px;">
				<li><a href="?ConnString=<%= request("ConnString") %>&ID=ALL">[ Analisi completa ]</a></li>
				<% for j = lbound(SchemaType) to ubound(SchemaType) %>
					<li><a href="?ConnString=<%= request("ConnString") %>&ID=<%= j %>">[ <%= SchemaName(j) %> ]</a></li>
				<%next %>
			</ul>
		</td>
	</tr>
</table>
	<%dim start, limit
	if isNumeric(request("ID")) AND request("ID")<>"" then
		start = CInteger(request("ID"))
		limit = start
	else
		if request("ID")<>"" then
			start = lbound(SchemaType)
			limit = ubound(SchemaType)
		else
			start = NULL
			limit = NULL
		end if
	end if
	if not isNull(start) then
		for i = start to limit %>
			<table cellspacing="0" cellpadding="0" border="0">
				<tr>
					<td style="padding-top:15px; padding-left:2px;">
						<a href="#top" name="<%= SchemaName(i) %>">
							[ top ]
						</a>
					</td>
				</tr>
				<tr>
					<td style="padding-top:3px;">
						<table cellspacing="1" cellpadding="0" class="tabella_madre">
							<caption><%= SchemaName(i) %></caption>
							<% on error resume next
							set rs = Conn.OpenSchema(SchemaType(i))
							
							if err.number<>0 then
								err.clear%>
								<tr><td class="content_b">Schema di analisi non supportato dal provider</td></tr>
							<%else
								on error goto 0
								'schema aperto correttamente %>
								<tr>
									<% for each field in rs.Fields %>
										<th class="center"><%= field.name %></th>
									<% next %>
								</tr>
								<%while not rs.eof %>
									<tr>
										<% for each field in rs.Fields %>
											<td class="content" style="padding-left:15px; padding-right:15px;" nowrap>
												<%= field.value %>
											</td>
										<% next %>
									</tr>
									<%rs.MoveNext 
								wend
							end if%>
						</table>
					</td>
				</tr>
			</table>
		<% next
	end if%>
<br>
<br>
<br>
</body>
</html>

<% 
conn.close
set conn = nothing
%>
