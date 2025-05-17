<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<%
dim conn, rs_s, rs_d, sql,field
dim contatore
sql = " SELECT * FROM tb_Circolari " + _
	  " WHERE (1=1) " + sql + _
	  " ORDER BY CI_Pubblicazione DESC, CI_Titolo"
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs_s = Server.CreateObject("ADODB.RecordSet")
set rs_d = Server.CreateObject("ADODB.RecordSet")
contatore = 0
rs_s.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

rs_d.open "mtb_documenti", conn, adOpenKeySet, adLockOptimistic
response.write "<h1>IMPORT</h1>" +vbCRLF
conn.begintrans
do while not rs_s.eof 

	contatore = contatore + 1
	response.write "<br>" & contatore

	rs_d.AddNew

	rs_d("doc_numero") = rs_s("CI_Numero")
	rs_d("doc_titolo_it") = rs_s("CI_Titolo")
	rs_d("doc_estratto_it") = rs_s("CI_Estratto")
	rs_d("doc_pubblicazione") = rs_s("CI_Pubblicazione")
	rs_d("doc_scadenza") = rs_s("CI_Scadenza")
	rs_d("doc_file_it") = rs_s("CI_File")
	rs_d("doc_visibile") = rs_s("CI_Visibile")
	rs_d("doc_protetto") = rs_s("CI_Protetto")

	if IsNull(rs_s("CI_idCategoria")) then
		dim sql_cat, rs_cat
		sql_cat="select top 1 * from mtb_documenti_categorie"
		set rs_cat = Server.CreateObject("ADODB.RecordSet")
		
		'necessario creare una categoria nel next memo2
		rs_cat.open sql_cat, conn, adOpenForwardOnly, adLockOptimistic, adCmdText	
		rs_d("doc_categoria_id") = rs_cat("catC_id")
		rs_cat.close
	else
		rs_d("doc_categoria_id") = rs_s("CI_idCategoria")
	end if

	'rs_d("") = rs_s("")


	rs_d.update
	rs_s.movenext

loop


'chiusura transazione di import
conn.committrans

rs_s.close
rs_d.close

%>