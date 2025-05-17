<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/TOOLS.ASP" -->
<!--#INCLUDE FILE="../library/TOOLS4Admin.ASP" -->
<!--#INCLUDE FILE="Comunicazioni_Tools.asp" -->

<% 
dim PAGINA, LINGUA, nextWeb_Version
dim nextWeb_Conn, rs, sql, conn, messageType

'Apro la connessione al dilayers
set nextWeb_Conn = Server.CreateObject("ADODB.Connection")
nextWeb_Conn.open Application("l_conn_ConnectionString"),"",""

' Apro la connessione al DATA
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

messageType = cIntero(request("type"))

nextWeb_Version = GetNextWebCurrentVersion(conn, rs)
	
'imposta variabili per next-web
Application("AZ_ID") = Application("AZ_ID")
Session("is_NextCom_Page") = true		'variabile usata in PagineScheda.asp del nextWeb per 
										'modificare il flusso dell'applicazione
dim var

if cInteger(request("PAGINA"))>0 then
	'pagina del corpo del messaggio gia' generata
	sql = "SELECT * FROM tb_pages WHERE id_page=" & cIntero(request("PAGINA"))
	rs.open sql, nextWeb_Conn, adOpenKeyset, adLockOptimistic, adCmdText
	
	PAGINA = rs("id_page")	'recupera id pagina 
	LINGUA = rs("lingua")
	
else
	'generazione pagina per corpo del messaggio
	set rs = Server.CreateObject("ADODB.RecordSet")

	'cancella eventuali pagine residue precedenti
	sql = "DELETE FROM tb_pagineSito WHERE nome_ps_IT LIKE '" & "Composizione next-email - " & Session("LOGIN_4_LOG") & "' "
	CALL nextWeb_Conn.execute(sql, 0, adExecuteNoRecords)
	sql = "DELETE FROM tb_pages WHERE nomepage LIKE '" & "Composizione next-email - " & Session("LOGIN_4_LOG") & "' "
    CALL nextWeb_Conn.execute(sql, 0, adExecuteNoRecords)
	
    
    'inserisce la nuova pagina
    sql = "SELECT * FROM tb_pages WHERE nomepage LIKE '" &  "Composizione next-email - " & Session("LOGIN_4_LOG") & "' " &_
          " AND id_webs = " & Application("AZ_ID")
    rs.open sql, nextWeb_Conn, adOpenKeyset, adLockOptimistic, adCmdText
    	
    rs.AddNew
    rs("nomepage") = "Composizione next-email - " & Session("LOGIN_4_LOG")
    rs("template") = false
    rs("id_Template") = 0
    rs("lingua") = LINGUA_ITALIANO
    rs("id_webs") = Application("AZ_ID")
    rs("contatore") = 0
    rs("contRes") = Date
    
    if nextWeb_Version=5 then
        rs("sfondoColore") = "#FFFFFF"
        rs("page_modData") = Date
        rs("contUtenti") = 0
        rs("contCrawler") = 0
        rs("contAltro") = 0
            
    else
        rs("visipage") = false
        rs("sfondo") = "#FFFFFF"
        rs("External_ID") = Session("LOGIN_4_LOG")
    end if
    rs.Update
	
	PAGINA = rs("id_page")	'recupera id pagina 
	LINGUA = rs("lingua")
	rs.close
	
	
	'inserisce la nuova paginaSito
    sql = "SELECT * FROM tb_pagineSito WHERE id_pagDyn_IT = " & PAGINA
    rs.open sql, nextWeb_Conn, adOpenKeyset, adLockOptimistic, adCmdText
    rs.AddNew
	rs("id_web") = Application("AZ_ID")
	rs("archiviata") = true
	rs("riservata") = false
	rs("id_pagDyn_IT") = PAGINA
	rs("nome_ps_IT") = "Composizione next-email - " & Session("LOGIN_4_LOG")
	rs("ps_insData") = Now()
	rs("ps_insAdmin_id") = Session("ID_ADMIN")
	rs("ps_modData") = Now()
	rs("ps_modAdmin_id") = Session("ID_ADMIN")
	rs("indicizzabile") = true
	rs.Update
	
	'CALL Ceck_page_exists(conn, rs)
	
    %>
	
	<script language="JavaScript" type="text/javascript">
		//imposta visualizzazione della pagina nel frame e numero di pagina nell'input
		opener.SetPreview( <%= PAGINA %>);
	</script>
<%end if 

rs.close
    	
set rs = nothing
	
CALL ComunicazioniNew_Wizard_Session_AddField(messageType, "email_nuova_pagina", PAGINA)

%>


<script language="JavaScript" type="text/javascript">
	<% Select case lcase(request.querystring("operazione"))
		case "modifica" %>
			//apre la pagina corrente in modifica
			document.location = "../<%= GetNextWebDirectory(nextWeb_Version) %>/loadShock.asp?PAGINA=<%= PAGINA %>&NEXTEMAIL=1";
		<% case "copia" %>
			//apre lo strumento di copia della pagina
			document.location = "../<%= GetNextWebDirectory(nextWeb_Version) %>/SitoPagineCopia.asp?nextmail=true&ID_S=&ID_D=<%= PAGINA %>&azione=COPIA";
		<% case "template" %>
			//apre lo strumento di associazione dei template
			document.location = "../<%= GetNextWebDirectory(nextWeb_Version) %>/SitoPagineTemplate.asp?nextmail=true&ID_STAGE=<%= PAGINA %>&nome_lingua=<%= GetNomeLingua(LINGUA)%>";
		<% case "lingua" %>
			//apre la gestione della lingua della pagina
			document.location = "ComunicazioniNew_Wizard_2_loadShock_lingua.asp?PAGINA=<%= PAGINA %>";
	<% end select %>
</script>
<%
nextWeb_Conn.close
set nextWeb_Conn = nothing
%>