<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% Server.ScriptTimeout = 2147483647 %>
<% response.buffer = false %>
<% response.charset = "UTF-8" 

%>
<!--#INCLUDE FILE="intestazione.asp"-->
<!--#INCLUDE VIRTUAL="amministrazione/nextB2B/Tools4Save_B2B.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione articoli importati - import B2B articoli"
dicitura.puls_new = "FINE"
dicitura.link_new = "default.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, s_rs, d_rs, d_rsv, sql
dim value, art_id, rel_id
set conn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.RecordSet")
set s_rs = Server.CreateObject("ADODB.RecordSet")
set d_rs = Server.CreateObject("ADODB.RecordSet")
set d_rsv = Server.CreateObject("ADODB.RecordSet")
conn.open Application("DATA_ConnectionString")

dim rs_guest, ID_CATEGORIA, ID_MARCA
set rs_guest = Server.CreateObject("ADODB.RecordSet")


ID_CATEGORIA = 1242 'Ricambi - Lavorwash
ID_MARCA = 573 'Lavorwash

%>
<div id="content">
<%
sql = " select [Codice], [Descrizione], [Listino 2013] from _listino_lavor_2013 "
s_rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText

dim variante
set variante = New GestioneVariante 
set variante.conn = conn

conn.begintrans
%>
<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:1px">
	<caption class="border">IMPORTAZIONE ARTICOLI DA AREA DI IMPORT A NEXT-b2b</caption>
	<tr>
		<td class="label" colspan="2">n&deg; articoli da importare: </td>
		<td class="content" colspan="4"><%= s_rs.recordcount%></td>
	</tr>
	<tr>
		<th width="5%">&nbsp;</th>
		<th width="15%">CODICE</th>
		<th width="20%">PREZZO</th>
		<th>DESCRIZIONE</th>
		<th width="18%">RECORD</th>
	</tr>
</table>
<% while not s_rs.eof %>
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border-bottom:1px">
		<%
		'esegue inserimento articolo B2B per questo articolo
		sql = "SELECT * FROM gtb_articoli WHERE art_tipologia_id = " & ID_CATEGORIA & " AND art_cod_int LIKE '"&Trim(s_rs("Codice"))&"'"
		d_rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
		
		if d_rs.eof then
			'nuovo inserimento
			d_rs.addNew
			d_rs("art_NoVenSingola") = false
			d_rs("art_se_accessorio") = false
			d_rs("art_ha_accessori") = false
			d_rs("art_insData") = NOW()
			d_rs("art_modData") = NOW()
			d_rs("art_in_confezione") = false
			d_rs("art_se_bundle") = false
			d_rs("art_in_bundle") = false
			d_rs("art_se_confezione") = false
			d_rs("art_varianti") = false
			d_rs("art_cod_int") = Trim(s_rs("Codice"))
			d_rs("art_nome_it") = Trim(s_rs("Descrizione"))
			d_rs("art_disabilitato") = false
			d_rs("art_unico") = false

			d_rs("art_prezzo_base") = cReal(Trim(s_rs("Listino 2013")))
			
			d_rs("art_spedizione_id") = 1
			d_rs("art_applicativo_id") = 38

			d_rs("art_iva_id") = 1
			d_rs("art_giacenza_min") = 1
			d_rs("art_qta_min_ord") = 1
			d_rs("art_lotto_riordino") = 1
			d_rs("art_qta_max_ord") = 1
			
			d_rs("art_marca_id") = ID_MARCA
			d_rs("art_tipologia_id") = ID_CATEGORIA
			d_rs.update
			
			
			'inserisce dati variante di default
			variante.InsertUpdate d_rs("art_id"), 0, "", _
								  d_rs("art_cod_int"), d_rs("art_cod_pro"), d_rs("art_cod_alt"), _
								  d_rs("art_prezzo_base"), false, 0, 0, "", _
								  false, 1, 1, 1
			%>
			<tr>
				<td width="5%" class="content">INSERIMENTO</td>
			<%
		else
			'modifica articolo già esistente
			d_rs("art_prezzo_base") = cReal(Trim(s_rs("Listino 2013")))
			
			CALL AggiornaPrezziVarianti(conn, rs_guest, d_rs("art_id"))
			
			d_rs.update
			
			%>
			<tr>
				<td width="5%" class="content">MODIFICA</td>
			<%
		end if
		d_rs.close
		
		%>
			<td width="15%" class="content"><%= Trim(s_rs("Codice")) %></td>
			<td width="20%" class="content"><%= cReal(Trim(s_rs("Listino 2013"))) %></td>
			<td class="content"><%= Trim(s_rs("Descrizione")) %></td>
			<td width="18%" class="content_b" style="text-align:right;"><%= s_rs.absoluteposition %> / <%= s_rs.recordcount %></td>
		</tr>
	</table>
	<% 
	s_rs.movenext
wend
s_rs.close


conn.committrans
%>

	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<tr>
			<td class="footer" style="text-align:right;"><a class="button" href="default.asp">FINE</a></td>
		</tr>
	</table>
	<br>
</div>
</body>
</html>
<% 
conn.close
set rs = nothing
set s_rs = nothing
set d_rs = nothing
set d_rsv = nothing
set conn = nothing
%>
