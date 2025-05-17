<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata

dicitura.iniz_sottosez(3)
dicitura.sottosezioni(1) = "IMPEGNI"
dicitura.links(1) = "Impegni.asp"
dicitura.sottosezioni(2) = "TIPOLOGIE"
dicitura.links(2) = "ImpegniTipologie.asp"
dicitura.sottosezioni(3) = "CONFIGURAZIONE"
dicitura.links(3) = "AgendaConfigura.asp"

dicitura.sezione = "Gestione impegni/appuntamenti - elenco"
dicitura.puls_new = ""
dicitura.link_new = ""
dicitura.scrivi_con_sottosez() 



dim i, conn, rsd, rs, sql, Pager

set Pager = new PageNavigator

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.RecordSet")
set rsd = Server.CreateObject("ADODB.RecordSet")


sql = " SELECT * FROM mtb_log_avvisi_spediti INNER JOIN " & _
      " 	mtb_impegni ON mtb_log_avvisi_spediti.las_impegno_id = mtb_impegni.imp_id INNER JOIN " & _
	  " 	tb_Utenti ON mtb_log_avvisi_spediti.las_id_utente_destinatario = tb_Utenti.ut_ID " & _
	  "		INNER JOIN tb_Indirizzario ON tb_Utenti.ut_NextCom_ID = tb_Indirizzario.IDElencoIndirizzi " & _
	  " ORDER BY las_data_spedizione DESC "
CALL Pager.OpenSmartRecordset(conn, rs, sql, 25)

%>
<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>
			Log e-mail di avviso spedite
			<% if not rs.eof then %>
				- n&ordm; <%= rs.recordcount %> e-mail spedite
			<% end if %>
		</caption>
		<% if not rs.eof then%>
			<tr>
				<th width="18%">DATA e ORA INVIO</th>
				<th width="23%">MITTENTE</th>
				<th>NOME IMPEGNO</th>
				<th width="23%">UTENTE DESTINATARIO</th>
			</tr>
			<%rs.AbsolutePage = Pager.PageNo
			while not rs.eof and rs.AbsolutePage = Pager.PageNo %>
				<tr>
					<td class="content"><%= DateTimeIta(rs("las_data_spedizione")) %></td>
					<td class="content"><%= GetAdminName(conn, cInteger(rs("las_id_admin_mittente"))) %></td>
					<td class="content"><%= rs("imp_titolo_it") %></td>
					<td class="content"><%= ContactFullName(rs) %></td>
				</tr>
				<%rs.movenext
			wend%>
			
			<tr>
				<td colspan="4" class="footer" style="text-align:left;">
					<% 	CALL Pager.Render_GroupNavigator(10, Pager.PageCount, "", "button", "button_disabled")%>
				</td>
			</tr>
		<% else %>
			<tr><td class="noRecords">Nessun download effettuato</th></tr>
		<% end if %>
	</table>
	&nbsp;
	</form>
</div>
</body>
</html>
<% 
rs.close
conn.close 
set rs = nothing
set rsd = nothing
set conn = nothing
%>