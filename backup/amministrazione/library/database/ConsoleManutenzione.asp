<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools4DataBase.asp" -->
<!--#INCLUDE FILE="../class_testata.asp" -->
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<% 
'*****************************************************************************************************************
'verifica dei permessi
CALL VerificaPermessiUtente(false)
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
<SCRIPT LANGUAGE="javascript" type="text/javascript">
    function AnalisiVariabiliAmbiente(){
        OpenPositionedScrollWindow('AnalisiVariabiliAmbiente.asp', 'AnalisiVariabili', 0, 0, 800, 600, true); 
        return void(0);
    }
</SCRIPT>
<body leftmargin=0 topmargin=0>
<!-- barra alta -->
<table width="740" cellspacing="0" cellpadding="0" border="0">
	<caption style="border:0px;">
		<table width="100%" cellspacing="0" cellpadding="0" border="0">
	  		<tr>
	  			<td style="text-align:right; padding-right:10px;">
					<a href="../logout.asp" class="menu" title="esci dall'appplicazione e torna all'area di login" <%= ACTIVE_STATUS %>>AMMINISTRAZIONE</a>
				</td>
	  		</tr>
		</table>
	</caption>
    <% CALL WriteChiusuraIntestazione("") %>
    <% 	dim header
    set header = New testata 
    header.iniz_sottosez(0)
    header.sezione = "gestione database"
    header.puls_new = "VARIABILI D'AMBIENTE"
    header.link_new = "javascript:AnalisiVariabiliAmbiente();"
    header.scrivi_con_sottosez() %>
</table>
<div id="content">	
	<% 
	dim var, ConnString_list, ConnString
	for each var in Application.Contents
		if instr(1, var, "ConnectionString", vbTextCompare)>0 then
			ConnString_list = ConnString_list & var & ";"
		end if
	next
	ConnString_list = split(ConnString_list, ";")
	
	dim NextWebVersion
	NextWebVersion = GetNextWebCurrentVersion(NULL, NULL)
    
    dim conn
    set conn = server.CreateObject("ADODB.Connection")
    for each ConnString in ConnString_list
        if ConnString<>"" then %>
		    <table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
			    <%'apertura connessione a database
				on error resume next
				conn.open Application(ConnString), "", ""
				on error goto 0
				
                'verifica se connessione aperta correttamente
				if conn.State <> adStateOpen  then %>
                     <caption class="border errore">connessione non valida "<%= ConnString %>"</caption>
                    <tr>
                        <td class="label errore" colspan="3"><%= Application(ConnString) %></td>
                    </tr>
                <% else
				    Select Case UCase(ConnString)
					    case "DATA_CONNECTIONSTRING" %>
						    <caption class="border ok">connessione dati principale
						<% case "DATA_ARCHIVE_CONNECTIONSTRING" %>
						    <caption class="border warning">connessione archivio dati
						<% case else %>
						    <caption class="border">connessione
                        <% end select %>
					    "<%= ConnString %>"
					</caption>
                    <tr>
                        <td class="label" colspan="2">stringa di connessione:</td>
                        <td class="content"><%= Application(ConnString) %></td>
                        <td class="content_center" style="width:18%;">
                            <a class="button_L2_block" href="javascript:void(0);" 
                    		   onclick="OpenPositionedScrollWindow('AnalisiConnessione.asp?ConnString=<%= Server.URLEncode(ConnString) %>', '_blank', 0, 0, 800,600, true);">
                    		    ANALIZZA CONNESSIONE
                    		</a>
                        </td>
                    </tr>
                    <tr>
                        <td class="label_no_width" rowspan="5" style="width:8%;">database:</td>
                        <td class="label_no_width" rowspan="2">nome:</td>
                        <td class="content_B" rowspan="2">
						    <%= GetDatabaseName(conn) %>
						    <%if NextWebVersion > 4 OR _
								 instr(1, GetDatabaseName(conn), "dblayers", vbTextCompare)>0 then
							    'database che contiene il next-web
								%>
								<span class="note visibile">&nbsp;(NEXT-web<%= IIF(NextWebVersion>3, " " & NextWebVersion, "") %>)</span>
							<% end if %>
						</td>
                        <td class="content_center">
                            <a class="button_L2_block" href="javascript:void(0);"
							   onclick="OpenPositionedScrollWindow('AnalisiDatabase.asp?ConnString=<%= Server.URLEncode(ConnString) %>', '_blank', 0, 0, 800,600, true);">
							    ANALIZZA DATABASE
							</a>
                        </td>
                    </tr>
                    <tr>
                        <td class="content_center">
                            <a class="button_L2_block" href="javascript:void(0);" 
							   onclick="OpenPositionedScrollWindow('ExecuteSQL.asp?ConnString=<%= Server.URLEncode(ConnString) %>', '_blank', 0, 0, 760,340, true);">
							    ESEGUI SQL
							</a>
                        </td>
                    </tr>
                    <tr>
                        <td style="vertical-align:middle;" class="label_no_width">versione:</td>
                        <td style="vertical-align:middle;" class="content_B">
							<%= ReadCurrentDbVersion(conn) %>
						</td>
                        <td class="content_center">
                            <% 
							dim HREF 
							if instr(1, ConnString, "DATA_ARCHIVE_ConnectionString", vbTextCompare)<1 then
            					select case DB_Type(conn)
            					    case DB_ACCESS 
            						    HREF = "Update_A_" & replace(GetDatabaseName(conn),"-","_") 
            						case DB_SQL
            						    HREF = "Update_S_" & replace(GetDatabaseName(conn),"-","_") 
            					    case else
            						    HREF = "Update_" & replace(GetDatabaseName(conn),"-","_") 
            					end select
                            else
                                HREF = "Archivio__UPDATE_DB"
                            end if 
							HREF = HREF + ".asp?ConnString=" + Server.URLEncode(ConnString)
							%>
                            <a class="button_block alert" style="line-height:20px;" href="javascript:void(0);" 
							   onclick="OpenPositionedScrollWindow('<%= HREF %>', '_blank', 0, 0, 760,500, true);">
							    AGGIORNA
							</a>
                        </td>
                    </tr>
                    <tr>
                        <td class="label_no_width" rowspan="2">dimensione:</td>
                        <td class="content" rowspan="2"><%= DatabaseSize(conn) %></td>
                        <td class="content_center">
                            <% if DB_Type(conn) = DB_SQL then%>
                                <a class="button_L2_block" href="javascript:void(0);" 
                                   title="Procedura che analizza la dimensione interna delle tabelle di SQL Server"
                                   onclick="OpenPositionedScrollWindow('AnalisiDimensioniTabelle.asp?ConnString=<%= Server.URLEncode(ConnString) %>', '_blank', 0, 0, 800,600, true);">
                                    DIMENSIONI TABELLE
                                </a>
                            <% else %>
                                <a class="button_L2_block_disabled" title="Funzione abilitata solo con SQL Server">DIMENSIONI TABELLE</a>
                            <% end if %>
                        </td>
                    </tr>
                    <tr>
                        <td class="content_center">
                            <a class="button_L2_block" href="javascript:void(0);" 
                               onclick="OpenAutoPositionedWindow('CompactDatabase.asp?ConnString=<%= Server.URLEncode(ConnString) %>', '_blank', 410,200);">
                                COMPATTA
                            </a>
                        </td>
                    </tr>
					<%if instr(1, ConnString, "DATA_ConnectionString", vbTextCompare)>0 then
                        if Application("DATA_ARCHIVE_ConnectionString")<>"" then%>
                            <tr>
                                <th colspan="4" class="L2">Gestione database di archivio</th>
                            </tr>
                            <tr>
						        <td class="label_no_width" colspan="3">
							        Archiviazione su database "DATA_ARCHIVE_ConnectionString" delle email inviate:
							    </td>
							    <td class="content_center">
                                    <a class="button_L2_block" href="javascript:void(0);" 
								       onclick="OpenPositionedScrollWindow('Archivio__IMPORT_email.asp', '_blank', 0, 0, 760,500, true);">
								        ARCHIVIA EMAIL
								    </a>
							    </td>
						    </tr>
							<tr>
						        <td class="label_no_width" colspan="3">
							        Archiviazione su database "DATA_ARCHIVE_ConnectionString" dei log_framework (escusi quelli dell'ultimo mese):
							    </td>
							    <td class="content_center">
                                    <a class="button_L2_block" href="javascript:void(0);" 
								       onclick="OpenPositionedScrollWindow('Archivio__IMPORT_log_framework.asp', '_blank', 0, 0, 760,500, true);">
								        ARCHIVIA LOG_FRAMEWORK
								    </a>
							    </td>
						    </tr>
                        <% end if 
                        if NextWebVersion >= 5 then %>
                            <tr>
                                <th colspan="4" class="L2">Gestione indice dei contenuti</th>
                            </tr>
                            <tr>
						        <td class="label_no_width" colspan="3">
                                    Verifica di tutti i contenuti presenti nell'indice:
                                </td>
                                <td class="content_center">
                                    <a class="button_L2_block" href="javascript:void(0);" 
								       onclick="OpenAutoPositionedScrollWindow('../IndexContent/VerificaContenuti.asp', '_blank', 760,500, true);">
								        VERIFICA CONTENUTI
								    </a>
							    </td>
                            </tr>
							<tr>
                                <th colspan="4" class="L2">Gestione applicativi del NEXT-framework</th>
                            </tr>
                            <tr>
						        <td class="label_no_width" colspan="3">
                                    Vai alle pagine di gestione degli applicativi:
                                </td>
                                <td class="content_center">
                                    <a class="button_L2_block" href="javascript:void(0);" 
								       onclick="OpenAutoPositionedScrollWindow('../Configuration/Applicazioni.asp', '_blank', 800, 600, true);">
								        GESTIONE APPLICATIVI
								    </a>
							    </td>
                            </tr>
							<tr>
                                <th colspan="4" class="L2">Aggiornamento BULK dell'indice</th>
                            </tr>
                            <tr>
						        <td class="label_no_width" colspan="3">
                                    Lancia script di aggiornamento dell'indice in maniera non ricorsiva.
                                </td>
                                <td class="content_center">
                                    <a class="button_L2_block" href="javascript:void(0);" 
								       onclick="OpenAutoPositionedScrollWindow('../IndexContent/RebuildIndex_Bulk.asp', '_blank', 800, 600, true);">
								        AGGIORNA
								    </a>
							    </td>
                            </tr>
						    <tr>
							     <td class="label_no_width" colspan="3">
                                    Lancia la query per eliminare gli spam.
                                </td>
							    <td class="content_center">
                                   <a class="button_L2_block" href="javascript:void(0);" 
							           onclick="OpenPositionedScrollWindow('CancellaSpam.asp?ConnString=<%= Server.URLEncode(ConnString) %>', '_blank', 0, 0, 743,335, true);">
							    PULIZZIA DATABASE
																			
                             </tr>							
						
                        <% end if
                    end if
                    conn.close
                end if%>
            </table>
        <% end if
    next 
	set conn = nothing%>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:20px;margin-top:20px;">
		<caption class="warning">Esecuzione scripts speciali</caption>
		<tr>
			<th colspan="2">ELENCO SCRIPT</th>
		</tr>
		<tr>
			<td class="content warning" colspan="2">
				ATTENZIONE: Gli script lanciati verranno eseguiti anche se il database non &egrave; corretto: sta allo script verificare correttezza di connessione e dati.
			</td>
		</tr>
		<% dim fso, scriptdir, scriptfile
		Set fso = CreateObject("Scripting.FileSystemObject")
		set scriptdir = fso.GetFolder(Server.MapPath("subscripts"))
		
		for each scriptfile in scriptdir.files %>
			<tr>
				<td class="label"><%= scriptfile.name %></td>
				<td class="content_right">
					<a class="button" href="subscripts/<%= scriptfile.name %>"><%= scriptfile.name %></a>
				</td>
			</tr>
		<% next
		set fso = nothing %>
	</table>
	
</div>
</body>
</html>
