<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<% Server.ScriptTimeout = 1073741824 %>
<!--#INCLUDE FILE="../Tools4DataBase.asp" -->
<!--#INCLUDE FILE="../../tools.asp" -->
<!--#INCLUDE FILE="../../tools4Admin.asp" -->

<%
dim conn, rs, sql
set conn = server.createobject("adodb.connection")
set rs = server.createobject("adodb.recordset")
conn.open Application("L_conn_ConnectionString")
conn.CommandTimeout = 180
CALL SendEmailSupportEX("Prestigeinternational.com - maintenance - AVVIO", "avvio script manutenzione database prestigeinternational.com")

response.write "Inizio Aggiornamento<br>"
CALL WriteLogAdminHttpRaw(conn, "AA_versione", 0, "Manutenzione - avvio", "avvio script manutenzione database prestigeinternational.com", "")

response.write "Alleggerimento db<br>"
CALL WriteLogAdminHttpRaw(conn, "AA_versione", 0, "Manutenzione - alleggerimento", "alleggerimento database miami e new york", "")
sql = ReadFileContent(Server.MapPath("Maintenance_prestigeinternational_alleggerimento.sql"))
CALL ExecuteMultipleSql(conn, sql, true)


response.write "Ricostruzione indici<br>"
CALL WriteLogAdminHttpRaw(conn, "AA_versione", 0, "Manutenzione - indici", "Ricostruzione indici database", "")
sql = ReadFileContent(Server.MapPath("Maintenance_prestigeinternational_indici.sql"))
CALL ExecuteMultipleSql(conn, sql, true)

response.write "Compattazione database<br>"
CALL WriteLogAdminHttpRaw(conn, "AA_versione", 0, "Manutenzione - compattazione", "Compattazione log e database", "")
sql = ReadFileContent(Server.MapPath("Maintenance_prestigeinternational_compattazione.sql"))
CALL ExecuteMultipleSql(conn, sql, true)

CALL WriteLogAdminHttpRaw(conn, "AA_versione", 0, "Manutenzione - completato", "fine script manutenzione database prestigeinternational.com", "")

CALL SendEmailSupportEX("Prestigeinternational.com - maintenance - FINE", "fine script manutenzione database prestigeinternational.com")

conn.close()

response.write "Fine Aggiornamento<br>"

set rs = nothing
set conn = nothing

%>
