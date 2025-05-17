<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>


<!--#INCLUDE FILE="../library/InitSex.asp" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="../library/class_testata.asp" -->
<!--#INCLUDE FILE="Tools_Memo2.asp" -->


<%
dim rs, conn, sql, listaImpegno, idImpegno
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")

CALL Parametri.LoadNew(conn, rs, NextMemo2)

sql = " SELECT imp_id, DATEADD(n , imp_anticipo_avviso*(-1), imp_data_ora_inizio) AS inizio_data_avviso, imp_data_ora_inizio " & _
	  " FROM mtb_impegni WHERE imp_invia_avviso = 1 AND imp_data_ora_inizio > " & SQL_Now(conn) & _
	  " AND DATEADD(n , imp_anticipo_avviso*(-1), imp_data_ora_inizio) < " & SQL_Now(conn)
'response.write sql & "<br>"
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
conn.beginTrans
while not rs.eof
	sql = "SELECT las_id FROM mtb_log_avvisi_spediti WHERE las_impegno_id = " & rs("imp_id") & _
		  " AND las_data_spedizione > " &SQL_DateTime(conn,rs("inizio_data_avviso"))& " AND las_data_spedizione < " &SQL_DateTime(conn,rs("imp_data_ora_inizio"))
	if cString(GetValueList(conn, NULL, sql)) = "" then
		'response.write sql & "<br>"
		CALL SendAvvisoImpegno(conn,rs("imp_id"),cIntero(Session("ID_PAGINA_AVVISO")))
	end if
	rs.moveNext
wend
conn.commitTrans


%>